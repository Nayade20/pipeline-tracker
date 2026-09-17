import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import time
import io
import os
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

# ─────────────────────────────────────────────
# CONFIGURACIÓN — pon aquí tu API Key
# ─────────────────────────────────────────────

# Lee configuración desde Streamlit Secrets (cloud) o directamente (local)
import streamlit as _st
try:
    HUBSPOT_API_KEY = _st.secrets["HUBSPOT_API_KEY"]
    APP_PASSWORD    = _st.secrets["APP_PASSWORD"]
    FTP_HOST        = _st.secrets.get("FTP_HOST", "")
    FTP_USER        = _st.secrets.get("FTP_USER", "")
    FTP_PASS        = _st.secrets.get("FTP_PASS", "")
    FTP_PATH        = _st.secrets.get("FTP_PATH", "/snapshots.json")
except Exception:
    HUBSPOT_API_KEY = "pat-eu1-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"  # ← solo para uso local
    APP_PASSWORD    = ""
    FTP_HOST = FTP_USER = FTP_PASS = ""
    FTP_PATH = "/snapshots.json"

OUTPUT_DIR = r"C:\Users\Administracion1\Zentralcom\Zentralcom S.L - Documentos\Administracion\NAYADE\CODIGOS PYTHON\informe automatico hubspot"

BASE_URL = "https://api.hubapi.com"

SNAPSHOTS_FILE = os.path.join(OUTPUT_DIR, "snapshots.json")

OWNER_NAMES = {
    "74753477":   "Iñigo Mangas",
    "75326441":   "Alejandro Alzas",
    "75326442":   "José R. Mendibil",
    "75631012":   "Rafa Quintanilla",
    "75887314":   "Onura Onura",
    "83853436":   "David Rivero",
    "91245262":   "Olatz Alkorta",
    "92982282":   "Administración Zentralcom",
    "102123176":  "Iban Ibañez",
    "1186290725": "Nayade Barrutieta",
}

PIPELINE_NAMES = {
    "default":   "CLIENTES",
    "664092197": "LEAD",
}

STAGE_LABELS = {
    "960918614":  "Objetivo clientes familias nuevas",
    "960918615":  "Presentadas nuevas familias — venta cruzada",
    "960918616":  "Facturas recibidas (cliente)",
    "960918617":  "Petición más info (cliente)",
    "960918618":  "Solicitud presupuesto proveedor (cliente)",
    "960918619":  "Oferta recibida proveedor (cliente)",
    "960918620":  "Informe preparado (cliente)",
    "960918621":  "Informe presentado (cliente)",
    "990804832":  "Negocio ganado",
    "1077494083": "Seguimiento subcontratación",
    "996161568":  "Negocio perdido",
    "974964831":  "Nuevo",
    "974964832":  "Empresa cualificada",
    "974964833":  "Contacto clave identificado",
    "974964834":  "Contactado sin feedback",
    "991276872":  "Contactado sin interés",
    "974964835":  "Interesado — sin reunión",
    "1077424678": "Seguimiento subcontratación (lead)",
    "1185458868": "Seguimiento Onura",
    "1185458869": "Seguimiento Rafa",
    "974964836":  "Reunión realizada",
    "974964837":  "Facturas recibidas (lead)",
    "975190032":  "Petición más info (lead)",
    "975190033":  "Solicitud presupuesto proveedor (lead)",
    "975190034":  "Oferta recibida proveedor (lead)",
    "990996102":  "Informe preparado (lead)",
    "990996103":  "Informe presentado (lead)",
    "991037466":  "Cierre ganado",
    "991037467":  "Cierre perdido",
    "1397904426": "Descartado subcontratación (lead)",
}

WON_STAGES  = {"990804832", "991037466"}
LOST_STAGES = {"996161568", "991037467"}
STALE_DAYS  = [14, 30, 60]

# Zona horaria de España — todas las fechas del dashboard se muestran en hora peninsular
TZ_ESPANA = "Europe/Madrid"

def ahora_espana():
    """Fecha y hora actual en España (sin zona horaria, para comparar con los datos)."""
    return pd.Timestamp.now(tz=TZ_ESPANA).tz_localize(None).to_pydatetime()

# ─────────────────────────────────────────────
# CLIENTE API HUBSPOT
# ─────────────────────────────────────────────

def hs_get(api_key, url, params=None):
    headers = {"Authorization": f"Bearer {api_key}"}
    resp = requests.get(url, headers=headers, params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()

def hs_post(api_key, url, payload):
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    resp = requests.post(url, headers=headers, json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json()

def test_connection(api_key):
    try:
        hs_get(api_key, f"{BASE_URL}/crm/v3/owners")
        return True
    except Exception:
        return False

def get_all_deals(api_key):
    props = [
        "dealname", "dealstage", "pipeline", "hubspot_owner_id",
        "amount", "createdate", "closedate",
        "hs_lastmodifieddate", "hs_last_activity_date", "notes_last_updated",
    ]
    all_deals = []
    params = {"limit": 100, "properties": ",".join(props)}
    while True:
        data = hs_get(api_key, f"{BASE_URL}/crm/v3/objects/deals", params)
        all_deals.extend(data.get("results", []))
        after = data.get("paging", {}).get("next", {}).get("after")
        if not after:
            break
        params["after"] = after
        time.sleep(0.1)
    return all_deals

def get_real_activity(api_key, since_ts_ms, until_ts_ms):
    """
    Obtiene actividad real: llamadas, emails, reuniones, notas y tareas
    asociadas a negocios en el período indicado.
    Devuelve: {deal_id: {"owner_id": str, "types": set}}
    """
    activity_endpoints = {
        "calls":    "/crm/v3/objects/calls",
        "emails":   "/crm/v3/objects/emails",
        "meetings": "/crm/v3/objects/meetings",
        "notes":    "/crm/v3/objects/notes",
        "tasks":    "/crm/v3/objects/tasks",
    }
    type_props = {
        "calls":    ["hs_timestamp", "hubspot_owner_id"],
        "emails":   ["hs_timestamp", "hubspot_owner_id"],
        "meetings": ["hs_timestamp", "hubspot_owner_id"],
        "notes":    ["hs_timestamp", "hubspot_owner_id"],
        "tasks":    ["hs_timestamp", "hubspot_owner_id"],
    }
    deal_activity = {}

    for act_type, url in activity_endpoints.items():
        # 1. Recoger actividades en el rango de fechas
        in_range = []
        after = None
        while True:
            params = {"limit": 100, "properties": ",".join(type_props[act_type])}
            if after:
                params["after"] = after
            try:
                data = hs_get(api_key, url, params)
            except Exception:
                break
            for item in data.get("results", []):
                ts_raw = item.get("properties", {}).get("hs_timestamp")
                if not ts_raw:
                    continue
                try:
                    ts_ms = int(ts_raw)
                except (ValueError, TypeError):
                    try:
                        ts_ms = int(pd.to_datetime(ts_raw, utc=True).timestamp() * 1000)
                    except Exception:
                        continue
                if since_ts_ms <= ts_ms <= until_ts_ms:
                    in_range.append(item)
            paging = data.get("paging", {})
            after = paging.get("next", {}).get("after")
            if not after:
                break
            time.sleep(0.1)

        if not in_range:
            continue

        # 2. Para cada actividad, buscar los negocios asociados
        owner_map = {
            item["id"]: item.get("properties", {}).get("hubspot_owner_id", "")
            for item in in_range
        }
        ids = list(owner_map.keys())

        for i in range(0, len(ids), 100):
            batch = ids[i:i + 100]
            assoc_url = f"{BASE_URL}/crm/v4/associations/{act_type}/deals/batch/read"
            payload = {"inputs": [{"id": aid} for aid in batch]}
            try:
                assoc_data = hs_post(api_key, assoc_url, payload)
                for result in assoc_data.get("results", []):
                    act_id   = str(result.get("from", {}).get("id", ""))
                    owner_id = owner_map.get(act_id, "")
                    for assoc in result.get("to", []):
                        deal_id = str(assoc.get("toObjectId") or assoc.get("id", ""))
                        if deal_id not in deal_activity:
                            deal_activity[deal_id] = {"owner_id": owner_id, "types": set()}
                        deal_activity[deal_id]["types"].add(act_type)
            except Exception:
                pass
            time.sleep(0.15)

    return deal_activity

def get_stage_history_bulk(api_key, deal_ids):
    results = {}
    for i in range(0, len(deal_ids), 10):
        batch = deal_ids[i:i + 10]
        payload = {
            "inputs": [{"id": str(did)} for did in batch],
            "propertiesWithHistory": ["dealstage"],
            "properties": ["dealname", "hubspot_owner_id", "pipeline"],
        }
        try:
            data = hs_post(api_key, f"{BASE_URL}/crm/v3/objects/deals/batch/read", payload)
            for item in data.get("results", []):
                did = item["id"]
                results[did] = {
                    "dealname":     item["properties"].get("dealname", ""),
                    "owner_id":     item["properties"].get("hubspot_owner_id", ""),
                    "pipeline":     item["properties"].get("pipeline", ""),
                    "stage_history": item.get("propertiesWithHistory", {}).get("dealstage", []),
                }
        except Exception:
            pass
        time.sleep(0.2)
    return results

def get_engagement_types(api_key, deal_ids):
    """
    Obtiene el tipo de última actividad por negocio usando el endpoint de engagements.
    Igual que el código de referencia: /engagements/v1/engagements/associated/deal/{id}
    """
    TIPO_MAP = {
        "MEETING":         "Reunión",
        "EMAIL":           "Email",
        "INCOMING_EMAIL":  "Email",   # correos recibidos
        "FORWARDED_EMAIL": "Email",   # correos reenviados
        "CALL":            "Llamada",
        "NOTE":            "Nota",
        "TASK":            "Tarea",
    }
    results = {}
    for deal_id in deal_ids:
        url = f"https://api.hubapi.com/engagements/v1/engagements/associated/deal/{deal_id}/paged?limit=100"
        try:
            data = hs_get(api_key, url)
            ultima_ts, tipo = None, "—"
            # Comparación en epoch ms: no depende de la zona horaria del servidor
            ahora_ms = int(time.time() * 1000)
            for item in data.get("results", []):
                eng = item.get("engagement", {})
                t   = eng.get("type", "")
                ts  = eng.get("timestamp")
                if t in TIPO_MAP and ts:
                    # Actividades futuras (reuniones agendadas) → no cuentan
                    if ts <= ahora_ms and (ultima_ts is None or ts > ultima_ts):
                        ultima_ts, tipo = ts, TIPO_MAP[t]
            results[str(deal_id)] = tipo
        except Exception:
            results[str(deal_id)] = "—"
        time.sleep(0.1)
    return results


# ─────────────────────────────────────────────
# ANÁLISIS DE DATOS
# ─────────────────────────────────────────────

def parse_dt(val):
    if not val:
        return pd.NaT
    try:
        # HubSpot devuelve fechas en UTC — se convierten a hora española
        return pd.to_datetime(val, utc=True).tz_convert(TZ_ESPANA).tz_localize(None)
    except Exception:
        return pd.NaT

def parse_deals(raw_deals):
    rows = []
    for d in raw_deals:
        p = d.get("properties", {})
        owner_id    = str(p.get("hubspot_owner_id") or "")
        pipeline_id = str(p.get("pipeline") or "")
        rows.append({
            "deal_id":            d["id"],
            "dealname":           p.get("dealname", ""),
            "dealstage":          p.get("dealstage", ""),
            "pipeline_id":        pipeline_id,
            "pipeline":           PIPELINE_NAMES.get(pipeline_id, pipeline_id),
            "owner_id":           owner_id,
            "owner":              OWNER_NAMES.get(owner_id, f"Owner {owner_id}"),
            "amount":             float(p.get("amount") or 0),
            "createdate":         parse_dt(p.get("createdate")),
            "closedate":          parse_dt(p.get("closedate")),
            "last_modified":      parse_dt(p.get("hs_lastmodifieddate")),
            "last_activity_date": parse_dt(p.get("hs_last_activity_date")),
            "notes_last_updated": parse_dt(p.get("notes_last_updated")),
        })
    return pd.DataFrame(rows)

def activity_by_owner(df, week_start, week_end):
    """
    Usa notes_last_updated — misma propiedad que usa HubSpot para 'Última actividad'.
    La tasa de actividad se calcula sobre negocios ABIERTOS (excluye ganados/perdidos).
    """
    active_mask = (
        df["notes_last_updated"].notna() &
        (df["notes_last_updated"] >= week_start) &
        (df["notes_last_updated"] <= week_end)
    )
    active    = df[active_mask]
    # df ya solo contiene negocios abiertos
    total_open = df.groupby("owner").size().reset_index(name="open_deals")
    act        = active.groupby("owner").size().reset_index(name="active_this_week")

    result = total_open.merge(act, on="owner", how="left").fillna(0)
    result["open_deals"]       = result["open_deals"].astype(int)
    result["active_this_week"] = result["active_this_week"].astype(int)
    result["total_deals"]      = result["open_deals"]  # compatibilidad con resto del código
    result["sin_actividad"]    = result["open_deals"] - result["active_this_week"]
    result["sin_actividad"]    = result["sin_actividad"].clip(lower=0)
    result["activity_rate"]    = (
        result["active_this_week"] / result["open_deals"] * 100
    ).where(result["open_deals"] > 0, 0).round(1)
    return result.sort_values("active_this_week", ascending=False)

def activity_detail_by_owner(df, real_activity):
    """Desglose de tipos de actividad por comercial (solo si se cargó actividad real)."""
    rows = []
    for deal_id, info in real_activity.items():
        owner_id = str(info.get("owner_id", ""))
        owner    = OWNER_NAMES.get(owner_id, f"Owner {owner_id}")
        deal_row = df[df["deal_id"] == deal_id]
        if not deal_row.empty:
            owner = deal_row.iloc[0]["owner"]
        for act_type in info.get("types", set()):
            rows.append({"owner": owner, "activity_type": act_type})
    if not rows:
        return pd.DataFrame(columns=["owner", "activity_type"])
    return pd.DataFrame(rows)

def stage_activity_summary(df, week_start, week_end):
    """Actividad por etapa usando notes_last_updated."""
    active = df[
        df["notes_last_updated"].notna() &
        (df["notes_last_updated"] >= week_start) &
        (df["notes_last_updated"] <= week_end)
    ]
    total  = df.groupby("dealstage").size().reset_index(name="total_deals")
    act    = active.groupby("dealstage").size().reset_index(name="active_this_week")
    result = total.merge(act, on="dealstage", how="left").fillna(0)
    result["active_this_week"] = result["active_this_week"].astype(int)
    result["stage_label"]      = result["dealstage"].map(STAGE_LABELS).fillna(result["dealstage"])
    return result.sort_values("active_this_week", ascending=False)

def stage_transitions(history_data, week_start, week_end):
    rows = []
    for deal_id, info in history_data.items():
        hist     = info.get("stage_history", [])
        owner    = OWNER_NAMES.get(str(info.get("owner_id", "")), f"Owner {info.get('owner_id','')}")
        pipeline = PIPELINE_NAMES.get(str(info.get("pipeline", "")), info.get("pipeline", ""))
        for i, entry in enumerate(hist):
            ts = parse_dt(entry.get("timestamp"))
            if ts and week_start <= ts <= week_end and i + 1 < len(hist):
                from_stage_id = hist[i+1].get("value","")
                to_stage_id   = entry.get("value","")

                # Ignorar si no hay etapa anterior (negocio recién creado)
                if not from_stage_id or not from_stage_id.strip():
                    continue

                # Ignorar si etapa anterior y nueva son iguales
                if from_stage_id == to_stage_id:
                    continue

                from_label = STAGE_LABELS.get(from_stage_id, from_stage_id)
                to_label   = STAGE_LABELS.get(to_stage_id, to_stage_id)

                rows.append({
                    "deal_id":    deal_id,
                    "dealname":   info.get("dealname", ""),
                    "owner":      owner,
                    "pipeline":   pipeline,
                    "from_stage": from_label,
                    "to_stage":   to_label,
                    "changed_at": ts,
                })
    cols = ["deal_id","dealname","owner","pipeline","from_stage","to_stage","changed_at"]
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=cols)

def new_deals_by_stage(df, week_start, week_end):
    new    = df[(df["createdate"] >= week_start) & (df["createdate"] <= week_end)]
    result = new.groupby("dealstage").size().reset_index(name="new_deals")
    result["stage_label"] = result["dealstage"].map(STAGE_LABELS).fillna(result["dealstage"])
    return result.sort_values("new_deals", ascending=False)

def owner_stage_matrix(df):
    df2 = df.copy()
    df2["stage_label"] = df2["dealstage"].map(STAGE_LABELS).fillna(df2["dealstage"])
    pivot = df2.pivot_table(index="owner", columns="stage_label",
                            values="deal_id", aggfunc="count", fill_value=0)
    pivot["TOTAL"] = pivot.sum(axis=1)
    return pivot.reset_index()

# Etapas cerradas — excluidas del dashboard principal
CLOSED_STAGES = WON_STAGES | LOST_STAGES | {
    "991037466",  # Cierre ganado
    "991037467",  # Cierre perdido
    "990804832",  # Negocio ganado
    "996161568",  # Negocio perdido
    "991276872",  # Contactado sin interés (ambos pipelines)
}

def stale_deals(df, thresholds, reference_date):
    """Negocios abiertos (sin etapas de cierre) sin actividad desde hace X días."""
    open_df = df.copy()  # df_all ya solo contiene negocios abiertos
    open_df["ref_date"]      = open_df["notes_last_updated"].fillna(open_df["last_modified"])
    open_df["days_inactive"] = (reference_date - open_df["ref_date"]).dt.days
    results = {}
    for days in thresholds:
        stale = open_df[open_df["days_inactive"] >= days].copy()
        stale["stage_label"] = stale["dealstage"].map(STAGE_LABELS).fillna(stale["dealstage"])
        results[days] = stale.sort_values("days_inactive", ascending=False)
    return results

# ─────────────────────────────────────────────
# PERSISTENCIA DE SNAPSHOTS
# ─────────────────────────────────────────────

import ftplib
import json

def _snapshots_to_json(snapshots: dict) -> str:
    data = {}
    date_cols = ["createdate", "closedate", "last_modified", "last_activity_date", "notes_last_updated"]
    for name, snap in snapshots.items():
        # Marcar las fechas con la zona horaria española antes de guardar,
        # para que al recargar el snapshot las horas no se desplacen.
        df_save = snap["df_week"].copy()
        for col in date_cols:
            if col in df_save.columns and pd.api.types.is_datetime64_any_dtype(df_save[col]):
                try:
                    df_save[col] = df_save[col].dt.tz_localize(TZ_ESPANA, ambiguous="NaT", nonexistent="NaT")
                except Exception:
                    pass
        data[name] = {
            "week_label": snap["week_label"],
            "df_week": df_save.to_json(date_format="iso"),
            "act_df":  snap["act_df"].to_json(date_format="iso") if snap.get("act_df") is not None else None,
            # Tipo de actividad congelado en el momento de guardar el snapshot
            "eng_types": snap.get("eng_types") or {},
        }
    return json.dumps(data, ensure_ascii=False, indent=2)

def _json_to_snapshots(raw: str) -> dict:
    """
    Convierte el JSON guardado en snapshots.

    Versión robusta: si el archivo remoto/local no tiene el formato esperado
    de snapshots, no rompe la app; devuelve {} para arrancar sin snapshots.
    Esto evita errores cuando snapshots.json contiene por accidente un DataFrame
    u otro JSON, por ejemplo {"deal_id": {...}, "dealname": {...}}.
    """
    if not raw or not str(raw).strip():
        return {}

    try:
        data = json.loads(raw)
    except Exception:
        return {}

    if not isinstance(data, dict):
        return {}

    snapshots = {}
    for name, snap in data.items():
        # Cada snapshot válido debe ser un dict con estas claves.
        # Si no, se ignora en vez de lanzar KeyError.
        if not isinstance(snap, dict):
            continue
        if "week_label" not in snap or "df_week" not in snap:
            continue

        try:
            df_week = pd.read_json(io.StringIO(snap["df_week"]))
        except Exception:
            # Compatibilidad por si pandas acepta directamente string en versiones antiguas
            try:
                df_week = pd.read_json(snap["df_week"])
            except Exception:
                continue

        for col in ["createdate", "closedate", "last_modified", "last_activity_date", "notes_last_updated"]:
            if col in df_week.columns:
                df_week[col] = pd.to_datetime(df_week[col], utc=True, errors="coerce").dt.tz_convert(TZ_ESPANA).dt.tz_localize(None)

        # deal_id como texto — evita desajustes str/int al comparar snapshots
        if "deal_id" in df_week.columns:
            df_week["deal_id"] = df_week["deal_id"].astype(str).str.replace(r"\.0$", "", regex=True)

        # Traducir etapas — compatibilidad con snapshots guardados antes de esta corrección.
        # pd.read_json puede convertir los IDs a números (960918614.0), por eso
        # se normaliza a texto y se elimina el sufijo ".0" antes de mapear.
        if "dealstage" in df_week.columns:
            df_week["dealstage"] = df_week["dealstage"].astype(str).str.replace(r"\.0$", "", regex=True)
            df_week["stage_label"] = df_week["dealstage"].map(STAGE_LABELS).fillna(df_week["dealstage"])

        act_df = None
        if snap.get("act_df"):
            try:
                act_df = pd.read_json(io.StringIO(snap["act_df"]))
            except Exception:
                try:
                    act_df = pd.read_json(snap["act_df"])
                except Exception:
                    act_df = None

        snapshots[str(name)] = {
            "week_label": snap.get("week_label", str(name)),
            "df_week": df_week,
            "act_df": act_df,
            # Snapshots antiguos no guardaban el tipo de actividad → {}
            "eng_types": snap.get("eng_types") if isinstance(snap.get("eng_types"), dict) else {},
        }

    return snapshots

def save_snapshots_to_file(snapshots: dict):
    """Guarda snapshots en FTP si está configurado, si no en disco local."""
    try:
        json_str = _snapshots_to_json(snapshots)
        if FTP_HOST and FTP_USER:
            # Guardar en FTP
            buf = io.BytesIO(json_str.encode("utf-8"))
            with ftplib.FTP(timeout=5) as ftp:
                ftp.connect(FTP_HOST, timeout=5)
                ftp.login(FTP_USER, FTP_PASS)
                ftp.storbinary(f"STOR {FTP_PATH}", buf)
        else:
            # Fallback: disco local
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            with open(SNAPSHOTS_FILE, "w", encoding="utf-8") as f:
                f.write(json_str)
        return True
    except Exception as e:
        return str(e)

def load_snapshots_from_file() -> dict:
    """Carga snapshots desde FTP si está configurado, si no desde disco local."""
    try:
        if FTP_HOST and FTP_USER:
            buf = io.BytesIO()
            with ftplib.FTP(timeout=5) as ftp:
                ftp.connect(FTP_HOST, timeout=5)
                ftp.login(FTP_USER, FTP_PASS)
                try:
                    ftp.retrbinary(f"RETR {FTP_PATH}", buf.write)
                except ftplib.error_perm:
                    return {}  # archivo no existe aún
            buf.seek(0)
            raw = buf.read().decode("utf-8")
            if not raw.strip():
                return {}
            return _json_to_snapshots(raw)
        else:
            if not os.path.exists(SNAPSHOTS_FILE):
                return {}
            with open(SNAPSHOTS_FILE, "r", encoding="utf-8") as f:
                return _json_to_snapshots(f.read())
    except Exception as e:
        st.session_state["ftp_load_error"] = str(e)
        return {}

def delete_snapshot_from_file(name: str, snapshots: dict):
    if name in snapshots:
        del snapshots[name]
        save_snapshots_to_file(snapshots)
    return snapshots


# ─────────────────────────────────────────────
# EXPORTACIÓN EXCEL
# ─────────────────────────────────────────────

HEADER_FILL  = PatternFill("solid", start_color="1A3A5C")
HEADER_FONT  = Font(bold=True, color="FFFFFF", name="Arial", size=11)
SUBHEAD_FILL = PatternFill("solid", start_color="2E75B6")
SUBHEAD_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
ACCENT_FILL  = PatternFill("solid", start_color="DEEAF1")
ALT_FILL     = PatternFill("solid", start_color="F5F9FD")
NORMAL_FONT  = Font(name="Arial", size=10)
BOLD_FONT    = Font(name="Arial", size=10, bold=True)
GREEN_FILL   = PatternFill("solid", start_color="E2EFDA")
RED_FILL     = PatternFill("solid", start_color="FCE4D6")
ORANGE_FILL  = PatternFill("solid", start_color="FFF2CC")
THIN         = Side(style="thin", color="CCCCCC")
THIN_BORDER  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def xl_header(ws, title, subtitle, cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
    ws["A1"] = title
    ws["A1"].font = Font(name="Arial", size=14, bold=True, color="1A3A5C")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=cols)
    ws["A2"] = subtitle
    ws["A2"].font = Font(name="Arial", size=10, color="595959")
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 6

def xl_style_headers(ws, row, ncols):
    for c in range(1, ncols+1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws.row_dimensions[row].height = 32

def xl_style_data(ws, start_row, end_row, ncols):
    for r in range(start_row, end_row+1):
        fill = ALT_FILL if r % 2 == 0 else PatternFill()
        for c in range(1, ncols+1):
            cell = ws.cell(r, c)
            cell.font   = NORMAL_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="left", vertical="center")
            if fill.fill_type:
                cell.fill = fill

def xl_auto_width(ws, min_w=10, max_w=40):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
