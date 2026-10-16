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
    "960918614":  "Objetivo clientes familias nuevas (antigua)",
    "960918615":  "Presentadas nuevas familias — venta cruzada (antigua)",
    "960918616":  "Facturas recibidas (cliente)",
    "960918617":  "Petición más info (cliente)",
    "960918618":  "Solicitud presupuesto proveedor (cliente)",
    "960918619":  "Oferta recibida proveedor (cliente)",
    "960918620":  "Informe preparado (cliente)",
    "960918621":  "Informe presentado (cliente)",
    "990804832":  "Negocio ganado",
    "1077494083": "Seguimiento subcontratación",
    "996161568":  "Negocio perdido",
    # Etapas nuevas del pipeline LEAD (octubre 2026)
    "1452888180": "BBDD",
    "1452887693": "Selección",
    "1452880567": "Reunión",
    # Etapas antiguas del pipeline LEAD (ya vacías; se mantienen para el historial)
    "974964831":  "Nuevo (antigua)",
    "974964832":  "Empresa cualificada (antigua)",
    "974964833":  "Contacto clave identificado (antigua)",
    "974964834":  "Contactado sin feedback (antigua)",
    "991276872":  "Contactado sin interés (antigua)",
    "1185458868": "Seguimiento Onura (antigua)",
    "1185458869": "Seguimiento Rafa (antigua)",
    "974964836":  "Reunión realizada (antigua)",
    # Etapas que siguen en uso
    "974964835":  "Interesado — sin reunión",
    "1077424678": "Seguimiento subcontratación (lead)",
    "974964837":  "Facturas recibidas (lead)",
    "975190032":  "Petición más info (lead)",
    "975190033":  "Solicitud presupuesto proveedor (lead)",
    "975190034":  "Oferta recibida proveedor (lead)",
    "990996102":  "Informe preparado (lead)",
    "990996103":  "Informe presentado (lead)",
    "991037466":  "Cierre ganado",
    "991037467":  "Cierre perdido",
    "1397904426": "Descartado",
}

WON_STAGES  = {"990804832", "991037466"}
LOST_STAGES = {"996161568", "991037467"}
STALE_DAYS  = [14, 30, 60]

# ─────────────────────────────────────────────
# NUEVA ESTRUCTURA LEAD — CUPO SEMANAL
# ─────────────────────────────────────────────
LEAD_PIPELINE_ID = "664092197"
ST_BBDD        = "1452888180"
ST_SELECCION   = "1452887693"
ST_REUNION     = "1452880567"
ST_DESCARTADO  = "1397904426"
ST_PRESUPUESTO = "974964837"   # Facturas recibidas (lead) — cuenta como presupuesto
ST_GANADO_LEAD = "991037466"   # Cierre ganado


# Fecha en la que se reorganizó el pipeline LEAD: ese día se movieron
# negocios en bloque a BBDD, Reunión y Descartado. Esas entradas NO cuentan
# como reuniones ni descartes reales en los KPIs.
MIGRACION_DIA = "2026-10-06"

# Valores por defecto (se pueden cambiar en la barra lateral)
CUPO_POR_DEFECTO      = 20
SEMANAS_CADUCIDAD_DEF = 3

# ─────────────────────────────────────────────
# ETAPAS CLAVE DE CADA PIPELINE (LEAD y CLIENTES)
# ─────────────────────────────────────────────
# Para cada pipeline: qué etapa es BBDD, Selección, Reunión (= visita),
# Facturas recibidas (= presupuesto), ganado (= cliente) y descartado.
# Si un id es None, la app lo busca en HubSpot por el NOMBRE de la etapa
# (así reconoce etapas nuevas sin tocar el código).
# "migracion" = día en que se reorganizó el pipeline moviendo negocios en bloque:
# las entradas de ese día NO cuentan como visitas, presupuestos ni clientes.
DATE_COLS = ["f_bbdd", "f_seleccion", "f_reunion", "f_descartado", "f_presupuesto", "f_ganado"]

PIPE_CONFIG = {
    LEAD_PIPELINE_ID: {
        "nombre": "LEAD", "migracion": "2026-10-06",
        "ids":    {"f_bbdd": ST_BBDD, "f_seleccion": ST_SELECCION, "f_reunion": ST_REUNION,
                   "f_descartado": ST_DESCARTADO, "f_presupuesto": ST_PRESUPUESTO, "f_ganado": ST_GANADO_LEAD},
        "nombres": {},
    },
    "default": {
        "nombre": "CLIENTES", "migracion": "2026-10-14",
        "ids":    {"f_bbdd": None, "f_seleccion": None, "f_reunion": None,
                   "f_descartado": "996161568",   # Negocio perdido
                   "f_presupuesto": "960918616",  # Facturas recibidas cliente
                   "f_ganado": "990804832"},      # Negocio ganado
        "nombres": {"f_bbdd": ["BBDD"], "f_seleccion": ["SELECCION"], "f_reunion": ["REUNION"],
                    "f_presupuesto": ["FACTURAS RECIBIDAS CLIENTE", "FACTURAS RECBIDAS CLIENTE"],
                    "f_ganado": ["NEGOCIO GANADO"], "f_descartado": ["NEGOCIO PERDIDO"]},
    },
}

# Se rellenan al arrancar, después de leer los pipelines de HubSpot
STAGE_MAP     = {}       # {pipeline_id: {"f_seleccion": id, ...}}
SELECCION_IDS = {ST_SELECCION}
BBDD_IDS      = {ST_BBDD}

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

def get_pipelines(api_key):
    """Etapas de cada pipeline en HubSpot: {pipeline_id: [(stage_id, nombre, metadata), ...]}."""
    try:
        data = hs_get(api_key, f"{BASE_URL}/crm/v3/pipelines/deals")
    except Exception as e:
        st.session_state["_pipelines_error"] = str(e)
        return {}
    return {
        str(pl["id"]): [(str(stg["id"]), stg.get("label", ""), stg.get("metadata", {}) or {})
                        for stg in pl.get("stages", [])]
        for pl in data.get("results", [])
    }

def _normaliza(txt):
    import unicodedata
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return " ".join(txt.upper().split())

_ARREGLOS = {"RECBIDAS": "RECIBIDAS", "PETCION": "PETICIÓN", "PETCIÓN": "PETICIÓN", "PETICION": "PETICIÓN",
             "INFORMACION": "INFORMACIÓN", "SELECCION": "SELECCIÓN", "REUNION": "REUNIÓN",
             "SUBCONTRATACION": "SUBCONTRATACIÓN", "INTERES": "INTERÉS", "CUALIFICADA ": "CUALIFICADA"}

def nombre_bonito(nombre):
    """'FACTURAS RECBIDAS CLIENTE' → 'Facturas recibidas cliente' (corrige erratas y tildes)."""
    palabras = [_ARREGLOS.get(w, w) for w in str(nombre).strip().upper().split()]
    txt = " ".join(palabras).capitalize()
    return txt.replace("Bbdd", "BBDD").replace("bbdd", "BBDD").replace("onura", "Onura").replace("rafa", "Rafa")

def resolver_etapas(pipes):
    """
    Decide qué etapa de cada pipeline es Selección, Reunión, etc.
    Usa el id fijo si existe en HubSpot; si no, busca la etapa por su nombre.
    También añade a STAGE_LABELS las etapas que el código todavía no conoce.
    """
    global STAGE_MAP, SELECCION_IDS, BBDD_IDS
    # Nombres de etapa: SIEMPRE los que tiene ahora HubSpot (así una etapa renombrada
    # sale con su nombre nuevo en todas las pestañas). Los nombres escritos en
    # STAGE_LABELS solo se usan para etapas ya borradas o si HubSpot no responde.
    nombres_lead = {_normaliza(n) for sid, n, *_ in pipes.get(LEAD_PIPELINE_ID, [])}
    for pid_, etapas in pipes.items():
        for sid, nombre, *resto in etapas:
            # Etapas marcadas como cerradas en HubSpot → se suman a las cerradas del código
            meta = resto[0] if resto else {}
            if str(meta.get("isClosed", "")).lower() == "true":
                CLOSED_STAGES.add(sid)
                try:
                    prob = float(meta.get("probability"))
                except (TypeError, ValueError):
                    prob = None
                if prob is not None and prob >= 1:
                    WON_STAGES.add(sid)
                elif prob is not None and prob <= 0:
                    LOST_STAGES.add(sid)
            if not nombre:
                continue
            base = nombre_bonito(nombre)
            # Mismo nombre que una etapa de LEAD → la de CLIENTES lleva "(cliente)"
            if pid_ == "default" and _normaliza(nombre) in nombres_lead:
                base += " (cliente)"
            STAGE_LABELS[sid] = base
    mapa = {}
    for pid, cfg in PIPE_CONFIG.items():
        etapas = pipes.get(pid, [])
        ids_hs = {sid for sid, *_ in etapas}
        por_nombre = {_normaliza(n): sid for sid, n, *_ in etapas}
        res = {}
        for col in DATE_COLS:
            fijo = cfg["ids"].get(col)
            if fijo and (not etapas or fijo in ids_hs):
                res[col] = fijo
                continue
            res[col] = None
            for nombre in cfg["nombres"].get(col, []):
                if _normaliza(nombre) in por_nombre:
                    res[col] = por_nombre[_normaliza(nombre)]
                    break
        mapa[pid] = res
    STAGE_MAP     = mapa
    SELECCION_IDS = {m["f_seleccion"] for m in mapa.values() if m.get("f_seleccion")} or {ST_SELECCION}
    BBDD_IDS      = {m["f_bbdd"] for m in mapa.values() if m.get("f_bbdd")} or {ST_BBDD}
    return mapa

def get_all_deals(api_key, stage_ids=()):
    props = [
        "dealname", "dealstage", "pipeline", "hubspot_owner_id",
        "amount", "createdate", "closedate",
        "hs_lastmodifieddate", "hs_last_activity_date", "notes_last_updated",
        "hs_v2_date_entered_current_stage",
    ]
    # Fecha de entrada en las etapas clave (versión nueva y antigua de HubSpot)
    for _sid in stage_ids:
        props.append(f"hs_v2_date_entered_{_sid}")
        props.append(f"hs_date_entered_{_sid}")
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
        "calls":    f"{BASE_URL}/crm/v3/objects/calls",
        "emails":   f"{BASE_URL}/crm/v3/objects/emails",
        "meetings": f"{BASE_URL}/crm/v3/objects/meetings",
        "notes":    f"{BASE_URL}/crm/v3/objects/notes",
        "tasks":    f"{BASE_URL}/crm/v3/objects/tasks",
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

def get_meetings(api_key, since_ts_ms, until_ts_ms):
    """
    Reuniones registradas en HubSpot (actividad tipo reunión) entre dos fechas.
    Se usa la búsqueda por fecha: el número de reuniones es pequeño, así que
    no hay riesgo de que HubSpot corte resultados (como sí pasaba con emails).
    No cuenta las reuniones marcadas como canceladas o "no presentado".
    """
    url = f"{BASE_URL}/crm/v3/objects/meetings/search"
    filas, after = [], None
    while True:
        payload = {
            "filterGroups": [{"filters": [{
                "propertyName": "hs_timestamp", "operator": "BETWEEN",
                "value": str(since_ts_ms), "highValue": str(until_ts_ms),
            }]}],
            "properties": ["hs_timestamp", "hubspot_owner_id", "hs_meeting_title", "hs_meeting_outcome"],
            "limit": 100,
        }
        if after:
            payload["after"] = after
        try:
            data = hs_post(api_key, url, payload)
        except Exception as e:
            st.session_state["_meetings_error"] = str(e)
            break
        for item in data.get("results", []):
            p = item.get("properties", {})
            outcome = (p.get("hs_meeting_outcome") or "").upper()
            if outcome in {"CANCELED", "CANCELLED", "NO_SHOW"}:
                continue
            filas.append({
                "meeting_id": item["id"],
                "fecha":      parse_dt(p.get("hs_timestamp")),
                "owner_id":   str(p.get("hubspot_owner_id") or ""),
                "titulo":     p.get("hs_meeting_title") or "Reunión",
            })
        after = data.get("paging", {}).get("next", {}).get("after")
        if not after or len(filas) >= 9900:
            break
        time.sleep(0.15)
    df = pd.DataFrame(filas, columns=["meeting_id", "fecha", "owner_id", "titulo"])
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df["owner"] = df["owner_id"].map(OWNER_NAMES).fillna("Owner " + df["owner_id"])
    return df

def get_stage_history_bulk(api_key, deal_ids):
    results = {}
    for i in range(0, len(deal_ids), 50):
        batch = deal_ids[i:i + 50]
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

def get_engagement_types(api_key, deal_ids, notes_last_updated_map=None):
    """
    Determina, para cada negocio, qué tipo de actividad corresponde a su fecha
    de "Última actividad" (notes_last_updated).

    Historial de intentos anteriores (por si hace falta recordar por qué está
    hecho así):
    1) Endpoint clásico de HubSpot (engagements v1): no siempre acertaba con
       la actividad realmente más reciente (mostraba "Nota" o "Tarea" cuando
       en realidad era un email).
    2) Listar TODO el historial de actividades de la cuenta y filtrar por
       fecha aquí mismo: con miles de emails/notas era demasiado lento y se
       quedaba a medias.
    3) Buscar en HubSpot (Search API) solo las actividades dentro de un rango
       de fechas de toda la cuenta, y luego mirar a qué negocio pertenecía
       cada una: comprobado con datos reales que, con cuentas grandes,
       HubSpot puede no devolver todas las actividades del rango (demasiadas
       páginas / límites de la API), así que algunas actividades —sobre todo
       emails, que son las más numerosas— se perdían y el negocio se quedaba
       sin ese dato (aparecía "—" o el tipo equivocado, por ejemplo "Tarea"
       cuando la última actividad real era un email).

    Solución actual (4): en vez de preguntar "¿qué actividades hay en este
    rango de fechas, en toda la cuenta?", se pregunta directamente
    "¿qué actividades tiene ASOCIADAS cada uno de estos negocios?" — mismo
    enfoque que usa la propia pantalla de HubSpot al abrir un negocio. Al
    partir de la lista (pequeña) de negocios en vez de la lista (enorme) de
    actividades de toda la cuenta, no hace falta limitar por fechas ni hay
    riesgo de que se queden actividades sin descargar. Comprobado con datos
    reales de varios negocios: la fecha del email/nota/tarea asociado
    coincide exactamente (al milisegundo) con "Última actividad".

    notes_last_updated_map: diccionario {deal_id (str): epoch_ms} con la fecha
    de "Última actividad" real de cada negocio (el mismo valor que se muestra
    en el panel). Es necesario para poder saber qué tipo corresponde a esa
    fecha — sin él, se devuelve "—" para todos los negocios.
    """
    if not notes_last_updated_map:
        return {str(d): "—" for d in deal_ids}

    TOLERANCIA_MS = 2 * 60 * 1000  # 2 minutos de margen, por posibles redondeos

    activity_types = {
        "calls":    "Llamada",
        "emails":   "Email",
        "meetings": "Reunión",
        "notes":    "Nota",
        "tasks":    "Tarea",
    }

    deal_ids_str = [str(d) for d in deal_ids]

    # deal_activity_ids[deal_id][tipo_label] = [ids de actividades de ese tipo asociadas]
    deal_activity_ids = {}
    # activity_ts[(act_type, activity_id)] = fecha (hs_timestamp) de esa actividad
    activity_ts = {}
    # errores_tipo[tipo_label] = primer mensaje de error visto para ese tipo,
    # por si HubSpot rechaza la petición (p. ej. por falta de permisos en la
    # API Key). Se enseña luego en la pantalla como aviso, para poder
    # diagnosticar sin necesidad de mirar el código.
    errores_tipo = {}

    for act_type, tipo_label in activity_types.items():
        # 1. Preguntar, para NUESTROS negocios, qué actividades de este tipo
        #    tienen asociadas (en vez de buscar entre todas las de la cuenta).
        activity_ids_needed = set()
        assoc_url = f"{BASE_URL}/crm/v4/associations/deals/{act_type}/batch/read"
        for i in range(0, len(deal_ids_str), 100):
            batch   = deal_ids_str[i:i + 100]
            payload = {"inputs": [{"id": d} for d in batch]}
            try:
                data = hs_post(api_key, assoc_url, payload)
            except Exception as e:
                if tipo_label not in errores_tipo:
                    detalle = getattr(getattr(e, "response", None), "text", "") or str(e)
                    errores_tipo[tipo_label] = f"asociaciones negocio→{act_type}: {detalle[:300]}"
                continue
            for result in data.get("results", []):
                deal_id = str(result.get("from", {}).get("id", ""))
                for assoc in result.get("to", []):
                    act_id = str(assoc.get("toObjectId") or assoc.get("id", ""))
                    if not act_id:
                        continue
                    deal_activity_ids.setdefault(deal_id, {}).setdefault(tipo_label, []).append(act_id)
                    activity_ids_needed.add(act_id)
            time.sleep(0.1)

        if not activity_ids_needed:
            continue

        # 2. Traer la fecha (hs_timestamp) de cada actividad encontrada
        ids       = list(activity_ids_needed)
        batch_url = f"{BASE_URL}/crm/v3/objects/{act_type}/batch/read"
        for i in range(0, len(ids), 100):
            batch   = ids[i:i + 100]
            payload = {"properties": ["hs_timestamp"], "inputs": [{"id": aid} for aid in batch]}
            try:
                data = hs_post(api_key, batch_url, payload)
            except Exception as e:
                if tipo_label not in errores_tipo:
                    detalle = getattr(getattr(e, "response", None), "text", "") or str(e)
                    errores_tipo[tipo_label] = f"detalle de {act_type}: {detalle[:300]}"
                continue
            for item in data.get("results", []):
                ts_raw = item.get("properties", {}).get("hs_timestamp")
                if not ts_raw:
                    continue
                try:
                    ts_ms = int(pd.to_datetime(ts_raw, utc=True).timestamp() * 1000)
                except Exception:
                    continue
                activity_ts[(act_type, item["id"])] = ts_ms
            time.sleep(0.1)

    # Guardamos los errores (si los hay) para poder mostrarlos como aviso en
    # pantalla — así se puede ver el motivo exacto sin tocar el código.
    try:
        st.session_state["_engagement_debug_errors"] = errores_tipo
    except Exception:
        pass

    # 3. Para cada negocio, elegir el tipo cuya fecha coincide con "Última actividad"
    results = {}
    for deal_id in deal_ids_str:
        notes_ts  = notes_last_updated_map.get(deal_id)
        tipos_ids = deal_activity_ids.get(deal_id, {})

        tipos = {}
        for act_type, tipo_label in activity_types.items():
            mejor_ts = None
            for act_id in tipos_ids.get(tipo_label, []):
                ts = activity_ts.get((act_type, act_id))
                if ts is not None and (mejor_ts is None or ts > mejor_ts):
                    mejor_ts = ts
            if mejor_ts is not None:
                tipos[tipo_label] = mejor_ts

        if notes_ts is None or not tipos:
            results[deal_id] = "—"
            continue
        # Si varias actividades caen dentro del margen de tolerancia (p. ej.
        # una nota y un email casi al mismo minuto), nos quedamos con la que
        # coincide MÁS EXACTAMENTE (menor diferencia), no con la primera que
        # se encuentre al recorrer los tipos.
        coincide, mejor_diff = None, None
        for tipo, ts in tipos.items():
            diff = abs(ts - notes_ts)
            if diff <= TOLERANCIA_MS and (mejor_diff is None or diff < mejor_diff):
                coincide, mejor_diff = tipo, diff
        # IMPORTANTE: si ningún tipo coincide EXACTAMENTE (con el margen de
        # tolerancia) con la fecha de "Última actividad", se muestra "—" en
        # vez de "adivinar" cogiendo el tipo más reciente encontrado. Antes
        # se hacía esa suposición y, cuando fallaba la detección de un tipo
        # (por ejemplo Email o Llamada), el negocio se quedaba mostrando
        # "Tarea" de forma incorrecta con solo con que hubiera una tarea
        # cualquiera asociada. Mejor decir "no lo sé" que decir algo mal.
        results[deal_id] = coincide if coincide else "—"
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

def parse_deals(raw_deals, stage_map=None):
    stage_map = stage_map or {}
    rows = []
    for d in raw_deals:
        p = d.get("properties", {})
        owner_id    = str(p.get("hubspot_owner_id") or "")
        pipeline_id = str(p.get("pipeline") or "")
        ids_pipe = stage_map.get(pipeline_id, {})
        stage_dates = {}
        for col in DATE_COLS:
            sid = ids_pipe.get(col)
            stage_dates[col] = (parse_dt(p.get(f"hs_v2_date_entered_{sid}") or p.get(f"hs_date_entered_{sid}"))
                                if sid else pd.NaT)
        rows.append({
            **stage_dates,
            "dia_migracion":      PIPE_CONFIG.get(pipeline_id, {}).get("migracion", ""),
            "f_etapa_actual":     parse_dt(p.get("hs_v2_date_entered_current_stage")),
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
    df = pd.DataFrame(rows)
    # Si una columna de fecha viene vacía en todos los negocios, asegurar tipo fecha
    for col in DATE_COLS + ["f_etapa_actual"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df

def completar_fechas_etapas(df, stage_map, history_data=None):
    """
    HubSpot NO crea la propiedad "fecha de entrada" para las etapas nuevas
    (BBDD, Selección, Reunión...), así que esas fechas llegan vacías.
    Se rellenan:
      1) con el historial de etapas del negocio (última vez que entró en esa etapa);
      2) si no hay historial, con la fecha de entrada en la etapa actual
         cuando el negocio sigue en esa etapa.
    """
    df = df.copy()
    history_data = history_data or {}
    for col in DATE_COLS:
        if col not in df.columns:
            continue
        sid_por_pipe = {pid: m.get(col) for pid, m in stage_map.items() if m.get(col)}
        faltan = df[col].isna() & df["pipeline_id"].isin(sid_por_pipe.keys())
        for idx in df.index[faltan]:
            sid = sid_por_pipe[df.at[idx, "pipeline_id"]]
            fecha = pd.NaT
            hist = history_data.get(str(df.at[idx, "deal_id"]), {}).get("stage_history", [])
            fechas = [parse_dt(h.get("timestamp")) for h in hist if str(h.get("value")) == str(sid)]
            fechas = [f for f in fechas if pd.notna(f)]
            if fechas:
                fecha = max(fechas)
            elif str(df.at[idx, "dealstage"]) == str(sid):
                fecha = df.at[idx, "f_etapa_actual"]
            df.at[idx, col] = fecha
        df[col] = pd.to_datetime(df[col], errors="coerce")
    return df

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
    "991276872",  # Contactado sin interés (antigua)
    "1397904426", # Descartado
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
# CUPO SEMANAL Y EMBUDO (pipeline LEAD)
# ─────────────────────────────────────────────

def _no_migracion(df, col):
    """True si la entrada en la etapa NO fue el día en que se reorganizó su pipeline."""
    dia = df["dia_migracion"].fillna("") if "dia_migracion" in df.columns else MIGRACION_DIA
    return df[col].dt.strftime("%Y-%m-%d").fillna("") != dia

def _en_periodo(serie, inicio, fin, excluir_migracion=False):
    """True si la fecha está dentro del período (opcionalmente sin el día de la migración)."""
    mask = serie.notna() & (serie >= inicio) & (serie <= fin)
    if excluir_migracion:
        mask &= serie.dt.strftime("%Y-%m-%d") != MIGRACION_DIA
    return mask

def cupo_semanal(df_lead, inicio, fin, hoy, cupo, semanas_caducidad):
    """
    Calcula el cupo de cada comercial en el pipeline LEAD.

    - En Selección: negocios que están ahora en la etapa Selección (ocupan hueco).
    - Nuevos: de los anteriores, los que entraron en Selección dentro del período.
    - Arrastrados: los que entraron en Selección antes del período.
    - Reuniones: negocios que entraron en Reunión dentro del período.
    - Descartados: negocios que entraron en Descartado dentro del período.
    - Huecos libres: cupo − negocios en Selección.
    Devuelve (resumen_por_comercial, negocios_en_seleccion).
    """
    df = df_lead.copy()
    en_sel = df[df["dealstage"].isin(SELECCION_IDS)].copy()
    fecha_sel = en_sel["f_seleccion"].fillna(en_sel["f_etapa_actual"])
    en_sel["fecha_seleccion"] = fecha_sel
    en_sel["semanas_en_seleccion"] = ((hoy - fecha_sel).dt.days // 7).fillna(0).astype(int)
    en_sel["es_nuevo"] = fecha_sel.notna() & (fecha_sel >= inicio) & (fecha_sel <= fin)

    def _estado(semanas):
        if semanas >= semanas_caducidad:
            return "🔴 Caducado"
        if semanas >= semanas_caducidad - 1:
            return "🟡 A punto de caducar"
        return "🟢 En plazo"
    en_sel["estado"] = en_sel["semanas_en_seleccion"].apply(_estado)

    reuniones   = df[_en_periodo(df["f_reunion"],    inicio, fin) & _no_migracion(df, "f_reunion")]
    descartados = df[_en_periodo(df["f_descartado"], inicio, fin) & _no_migracion(df, "f_descartado")]

    owners = sorted(set(en_sel["owner"]) | set(reuniones["owner"]) | set(descartados["owner"]))
    filas = []
    for o in owners:
        sel_o = en_sel[en_sel["owner"] == o]
        n_sel = len(sel_o)
        filas.append({
            "owner":        o,
            "en_seleccion": n_sel,
            "nuevos":       int(sel_o["es_nuevo"].sum()),
            "arrastrados":  int((~sel_o["es_nuevo"]).sum()),
            "reuniones":    int((reuniones["owner"] == o).sum()),
            "descartados":  int((descartados["owner"] == o).sum()),
            "caducados":    int((sel_o["estado"] == "🔴 Caducado").sum()),
            "huecos_libres": max(cupo - n_sel, 0),
            "exceso":        max(n_sel - cupo, 0),
        })
    cols = ["owner", "en_seleccion", "nuevos", "arrastrados", "reuniones",
            "descartados", "caducados", "huecos_libres", "exceso"]
    resumen = pd.DataFrame(filas, columns=cols)
    if not resumen.empty:
        resumen = resumen.sort_values(["reuniones", "en_seleccion"], ascending=False)
    return resumen, en_sel, reuniones, descartados

def en_seleccion_al_cierre(history_data, sel_ids, momento):
    """
    Negocios que estaban en la etapa Selección en un momento concreto (p. ej. el domingo
    a las 23:59 de una semana pasada). Se reconstruye con el historial de etapas de HubSpot:
    la etapa de cada negocio en ese momento es la última a la que entró antes de esa hora.
    """
    ids = set()
    for did, info in history_data.items():
        cambios = []
        for e in info.get("stage_history", []):
            ts = parse_dt(e.get("timestamp"))
            if pd.notna(ts):
                cambios.append((ts, str(e.get("value", ""))))
        etapa = None
        for ts, valor in sorted(cambios):
            if ts <= momento:
                etapa = valor
            else:
                break
        if etapa in sel_ids:
            ids.add(str(did))
    return ids

MESES_ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
KPI_ORDEN = ["Seleccionados", "Visitas", "Reuniones", "Presupuestos", "Clientes"]
KPI_DESCRIPCION = {
    "Seleccionados": "Negocios que entran en la etapa Selección (pasan al cupo)",
    "Visitas":       "Negocios que entran en la etapa Reunión del pipeline LEAD",
    "Reuniones":     "Reuniones registradas como actividad en HubSpot (sin canceladas)",
    "Presupuestos":  "Negocios que entran en Facturas recibidas (LEAD)",
    "Clientes":      "Negocios que entran en Cierre ganado (LEAD)",
}

def etiqueta_semana(fecha):
    """'Sem 41' según el calendario (semana ISO, empieza en lunes)."""
    iso = pd.Timestamp(fecha).isocalendar()
    return f"Sem {iso.week:02d}"

def estilo_degradado(tabla, rgb=(24, 95, 165), por_fila=True):
    """Colorea las celdas de una tabla de números (más intenso = valor más alto) sin usar matplotlib."""
    cols = [c for c in tabla.columns if c != "Total"]
    def _colorear(bloque):
        out = pd.DataFrame("", index=bloque.index, columns=bloque.columns)
        for idx in bloque.index:
            fila = bloque.loc[idx, cols]
            maximo = fila.max() if por_fila else bloque[cols].values.max()
            for c in cols:
                v = fila[c]
                if maximo and v > 0:
                    a = 0.12 + 0.55 * (v / maximo)
                    out.loc[idx, c] = f"background-color: rgba({rgb[0]},{rgb[1]},{rgb[2]},{a:.2f})"
        if "Total" in bloque.columns:
            out["Total"] = "font-weight: bold"
        return out
    return tabla.style.apply(_colorear, axis=None)

def kpi_eventos(df_lead, meetings_df=None):
    """
    Lista de 'eventos' (una fila por cada visita, presupuesto, cliente...),
    con su fecha, comercial y negocio. Todos los KPIs se calculan a partir de aquí.
    No incluye los movimientos en bloque del día de la reorganización.
    """
    fuentes = [
        ("Seleccionados", "f_seleccion",   False),
        ("Visitas",       "f_reunion",     True),
        # En la reorganización solo se movieron negocios en bloque a BBDD, Selección,
        # Reunión y Descartado: por eso solo las visitas descartan el día de la migración.
        ("Presupuestos",  "f_presupuesto", False),
        ("Clientes",      "f_ganado",      False),
    ]
    partes = []
    for kpi, col, excl in fuentes:
        if col not in df_lead.columns:
            continue
        sub = df_lead[df_lead[col].notna()]
        if excl:
            sub = sub[_no_migracion(sub, col)]
        if sub.empty:
            continue
        partes.append(pd.DataFrame({
            "fecha": sub[col].values, "kpi": kpi,
            "owner": sub["owner"].values, "detalle": sub["dealname"].values,
            "deal_id": sub["deal_id"].astype(str).values,
        }))
    if meetings_df is not None and not meetings_df.empty:
        m = meetings_df[meetings_df["fecha"].notna()]
        partes.append(pd.DataFrame({
            "fecha": m["fecha"].values, "kpi": "Reuniones",
            "owner": m["owner"].values, "detalle": m["titulo"].values,
        }))
    if not partes:
        return pd.DataFrame(columns=["fecha", "kpi", "owner", "detalle", "semana", "lunes", "mes", "mes_inicio"])
    ev = pd.concat(partes, ignore_index=True)
    ev["fecha"]      = pd.to_datetime(ev["fecha"])
    ev["lunes"]      = (ev["fecha"] - pd.to_timedelta(ev["fecha"].dt.weekday, unit="D")).dt.normalize()
    ev["semana"]     = ev["fecha"].apply(etiqueta_semana)
    ev["mes_inicio"] = ev["fecha"].dt.to_period("M").dt.to_timestamp()
    ev["mes"]        = ev["mes_inicio"].apply(lambda d: f"{MESES_ES[d.month - 1]} {d.year}")
    return ev

def periodos_kpi(frecuencia, n_periodos, hoy):
    """Lista de (etiqueta, inicio, fin) de los últimos n períodos, del más antiguo al actual."""
    hoy = pd.Timestamp(hoy)
    out = []
    if frecuencia == "Semana":
        lunes = (hoy - pd.Timedelta(days=hoy.weekday())).normalize()
        for i in range(n_periodos - 1, -1, -1):
            ini = lunes - pd.Timedelta(weeks=i)
            fin = ini + pd.Timedelta(days=7) - pd.Timedelta(microseconds=1)
            out.append((f"{etiqueta_semana(ini)} · {ini.strftime('%d/%m')}", ini, fin))
    else:
        primero = hoy.normalize().replace(day=1)
        for i in range(n_periodos - 1, -1, -1):
            ini = primero - pd.DateOffset(months=i)
            fin = ini + pd.DateOffset(months=1) - pd.Timedelta(microseconds=1)
            out.append((f"{MESES_ES[ini.month - 1]} {ini.year}", ini, fin))
    return out

def tabla_kpis(eventos, periodos, kpis=KPI_ORDEN, owner=None):
    """Tabla tipo cuadro de mando: una fila por KPI y una columna por período (+ Total)."""
    ev = eventos if owner is None else eventos[eventos["owner"] == owner]
    data = {}
    for etiqueta, ini, fin in periodos:
        sub = ev[(ev["fecha"] >= ini) & (ev["fecha"] <= fin)]
        cuenta = sub.groupby("kpi").size()
        data[etiqueta] = [int(cuenta.get(k, 0)) for k in kpis]
    tabla = pd.DataFrame(data, index=kpis)
    tabla["Total"] = tabla.sum(axis=1)
    tabla.index.name = "KPI"
    return tabla

def tabla_comercial_periodo(eventos, periodos, kpi):
    """Comerciales × períodos para un KPI concreto."""
    filas = {}
    owners = sorted(eventos["owner"].dropna().unique())
    for etiqueta, ini, fin in periodos:
        sub = eventos[(eventos["kpi"] == kpi) & (eventos["fecha"] >= ini) & (eventos["fecha"] <= fin)]
        cuenta = sub.groupby("owner").size()
        filas[etiqueta] = [int(cuenta.get(o, 0)) for o in owners]
    tabla = pd.DataFrame(filas, index=owners)
    tabla["Total"] = tabla.sum(axis=1)
    tabla = tabla[tabla["Total"] > 0].sort_values("Total", ascending=False)
    tabla.index.name = "Comercial"
    return tabla

def excel_semana(tabla1, tabla2, semana, hoy, detalle_df=None):
    """Excel sencillo con las dos tablas de la pestaña 'Esta semana'."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Esta semana"
    xl_header(ws, f"Seguimiento semanal — {semana}", f"Generado el {hoy.strftime('%d/%m/%Y %H:%M')}", 6)
    r = 4
    ws.cell(r, 1, "¿Cómo va cada comercial?").font = Font(name="Arial", size=11, bold=True, color="1A3A5C")
    r += 1
    cols1 = list(tabla1.columns)
    for ci, h in enumerate(cols1, 1):
        ws.cell(r, ci, h)
    xl_style_headers(ws, r, len(cols1))
    ini = r + 1
    for _, fila in tabla1.iterrows():
        r += 1
        for ci, c in enumerate(cols1, 1):
            v = fila[c]
            ws.cell(r, ci, int(v) if c != "Comercial" else str(v))
    if r >= ini:
        xl_style_data(ws, ini, r, len(cols1))
        for ci in range(1, len(cols1) + 1):
            ws.cell(r, ci).font = Font(name="Arial", size=10, bold=True)

    r += 3
    ws.cell(r, 1, "¿Cómo vamos semana a semana?").font = Font(name="Arial", size=11, bold=True, color="1A3A5C")
    r += 1
    cols2 = ["KPI"] + list(tabla2.columns)
    for ci, h in enumerate(cols2, 1):
        ws.cell(r, ci, h)
    xl_style_headers(ws, r, len(cols2))
    ini = r + 1
    for kpi, fila in tabla2.iterrows():
        r += 1
        ws.cell(r, 1, kpi)
        for ci, v in enumerate(fila.values, 2):
            ws.cell(r, ci, int(v))
    xl_style_data(ws, ini, r, len(cols2))
    xl_auto_width(ws)

    # Hoja con los negocios que hay detrás de cada número
    if detalle_df is not None and not detalle_df.empty:
        wd = wb.create_sheet("Negocios")
        cols3 = list(detalle_df.columns)
        xl_header(wd, "Qué negocios son", f"Visitas, presupuestos y clientes de las últimas 4 semanas · {semana}", len(cols3))
        for ci, h in enumerate(cols3, 1):
            wd.cell(4, ci, h)
        xl_style_headers(wd, 4, len(cols3))
        r3 = 4
        for _, fila in detalle_df.iterrows():
            r3 += 1
            for ci, c in enumerate(cols3, 1):
                wd.cell(r3, ci, str(fila[c]))
        xl_style_data(wd, 5, r3, len(cols3))
        xl_auto_width(wd, max_w=50)
        wd.freeze_panes = "A5"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def excel_kpis(eventos, periodos_sem, periodos_mes, titulo_rango):
    """Excel con los KPIs: semanal, mensual, por comercial y detalle de cada evento."""
    wb = Workbook()
    wb.remove(wb.active)

    def _cabecera(ws, titulo, subtitulo, cols):
        xl_header(ws, titulo, subtitulo, len(cols))
        for ci, h in enumerate(cols, 1):
            ws.cell(4, ci, h)
        xl_style_headers(ws, 4, len(cols))
        ws.freeze_panes = "B5"

    def _hoja_tabla(nombre, titulo, tabla):
        ws = wb.create_sheet(nombre)
        cols = [tabla.index.name or ""] + list(tabla.columns)
        _cabecera(ws, titulo, titulo_rango, cols)
        r = 5
        for idx, fila in tabla.iterrows():
            ws.cell(row=r, column=1, value=str(idx))
            for j, v in enumerate(fila.values, start=2):
                ws.cell(row=r, column=j, value=int(v))
            r += 1
        xl_style_data(ws, 5, r - 1, len(cols))
        for rr in range(5, r):
            ws.cell(row=rr, column=len(cols)).font = Font(bold=True)
        xl_auto_width(ws)
        return ws

    _hoja_tabla("KPIs semanales", "KPIs por semana — pipeline LEAD", tabla_kpis(eventos, periodos_sem))
    _hoja_tabla("KPIs mensuales", "KPIs por mes — pipeline LEAD",   tabla_kpis(eventos, periodos_mes))

    # Por comercial: una fila por comercial y semana
    ws = wb.create_sheet("Por comercial")
    cols = ["Comercial", "Semana"] + KPI_ORDEN
    _cabecera(ws, "KPIs por comercial y semana", titulo_rango, cols)
    r = 5
    for owner in sorted(eventos["owner"].dropna().unique()):
        for etiqueta, ini, fin in periodos_sem:
            sub = eventos[(eventos["owner"] == owner) & (eventos["fecha"] >= ini) & (eventos["fecha"] <= fin)]
            cuenta = sub.groupby("kpi").size()
            valores = [int(cuenta.get(k, 0)) for k in KPI_ORDEN]
            if sum(valores) == 0:
                continue
            ws.cell(row=r, column=1, value=owner)
            ws.cell(row=r, column=2, value=etiqueta)
            for j, v in enumerate(valores, start=3):
                ws.cell(row=r, column=j, value=v)
            r += 1
    xl_style_data(ws, 5, max(r - 1, 5), len(cols))
    xl_auto_width(ws)

    # Detalle
    ws = wb.create_sheet("Detalle")
    cols = ["Fecha", "Semana", "Mes", "KPI", "Comercial", "Negocio / reunión"]
    _cabecera(ws, "Detalle de cada evento", titulo_rango, cols)
    ini_total = min(p[1] for p in periodos_sem + periodos_mes)
    det = eventos[eventos["fecha"] >= ini_total].sort_values("fecha", ascending=False)
    r = 5
    for _, e in det.iterrows():
        ws.cell(row=r, column=1, value=e["fecha"].strftime("%d/%m/%Y"))
        ws.cell(row=r, column=2, value=e["semana"])
        ws.cell(row=r, column=3, value=e["mes"])
        ws.cell(row=r, column=4, value=e["kpi"])
        ws.cell(row=r, column=5, value=e["owner"])
        ws.cell(row=r, column=6, value=str(e["detalle"]))
        r += 1
    xl_style_data(ws, 5, max(r - 1, 5), len(cols))
    xl_auto_width(ws, max_w=60)

    # Definiciones
    ws = wb.create_sheet("Definiciones")
    _cabecera(ws, "Qué cuenta cada KPI", f"Semanas según calendario (lunes a domingo). No incluye los movimientos en bloque del {pd.Timestamp(MIGRACION_DIA).strftime('%d/%m/%Y')}.", ["KPI", "Qué cuenta"])
    for i, k in enumerate(KPI_ORDEN, start=5):
        ws.cell(row=i, column=1, value=k)
        ws.cell(row=i, column=2, value=KPI_DESCRIPCION[k])
    xl_style_data(ws, 5, 4 + len(KPI_ORDEN), 2)
    xl_auto_width(ws, max_w=80)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


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
        # Las fechas de entrada en etapas no hacen falta en los snapshots
        _extra = [c for c in DATE_COLS + ["f_etapa_actual", "dia_migracion"] if c in df_save.columns]
        if _extra:
            df_save = df_save.drop(columns=_extra)
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
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
        ws.column_dimensions[col_letter].width = max(min_w, min(max_len+2, max_w))

def xl_sheet_resumen(wb, week_label, act_df, stage_df, trans_df, new_df, stale_data):
    ws = wb.create_sheet("📊 Resumen semanal")
    xl_header(ws, "Resumen semanal del pipeline", f"Semana: {week_label}", 4)
    r = 4
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws.cell(r, 1, "INDICADORES CLAVE").fill = SUBHEAD_FILL
    ws.cell(r, 1).font = SUBHEAD_FONT
    r += 1
    total   = int(act_df["total_deals"].sum())
    active  = int(act_df["active_this_week"].sum())
    rate    = round(active/total*100, 1) if total else 0
    stale14 = len(stale_data.get(14, pd.DataFrame()))
    for label, val in [
        ("Total negocios",              total),
        ("Con actividad esta semana",   active),
        ("Tasa de actividad",           f"{rate}%"),
        ("Cambios de etapa detectados", len(trans_df)),
        ("Negocios nuevos",             int(new_df["new_deals"].sum()) if not new_df.empty else 0),
        ("Negocios estancados >14 días",stale14),
    ]:
        ws.cell(r, 1, label).font = NORMAL_FONT
        ws.cell(r, 2, val).font   = BOLD_FONT
        r += 1
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    ws.cell(r, 1, "ACTIVIDAD POR COMERCIAL").fill = SUBHEAD_FILL
    ws.cell(r, 1).font = SUBHEAD_FONT
    r += 1
    for ci, h in enumerate(["Comercial","Total negocios","Con actividad","Tasa %"], 1):
        ws.cell(r, ci, h)
    xl_style_headers(ws, r, 4)
    ds = r+1
    for _, row in act_df.iterrows():
        r += 1
        ws.cell(r, 1, row["owner"])
        ws.cell(r, 2, int(row["total_deals"]))
        ws.cell(r, 3, int(row["active_this_week"]))
        ws.cell(r, 4, f"{row['activity_rate']}%")
        fill = GREEN_FILL if row["activity_rate"] >= 20 else (ORANGE_FILL if row["activity_rate"] >= 10 else RED_FILL)
        ws.cell(r, 3).fill = fill
    xl_style_data(ws, ds, r, 4)
    xl_auto_width(ws)
    ws.freeze_panes = "A5"

def xl_sheet_actividad(wb, week_label, act_df):
    ws = wb.create_sheet("👤 Actividad comercial")
    xl_header(ws, "Actividad por comercial", f"Semana: {week_label}", 5)
    r = 4
    for ci, h in enumerate(["Comercial","Total negocios","Con actividad","Sin actividad","Tasa %"], 1):
        ws.cell(r, ci, h)
    xl_style_headers(ws, r, 5)
    ds = r+1
    for _, row in act_df.iterrows():
        r += 1
        ws.cell(r, 1, row["owner"])
        ws.cell(r, 2, int(row["total_deals"]))
        ws.cell(r, 3, int(row["active_this_week"]))
        ws.cell(r, 4, int(row["sin_actividad"]))
        ws.cell(r, 5, row["activity_rate"])
        fill = GREEN_FILL if row["activity_rate"] >= 20 else (ORANGE_FILL if row["activity_rate"] >= 10 else RED_FILL)
        ws.cell(r, 3).fill = fill
    xl_style_data(ws, ds, r, 5)
    chart = BarChart()
    chart.title = "Actividad esta semana por comercial"
    chart.height, chart.width = 12, 20
    data_ref = Reference(ws, min_col=3, max_col=3, min_row=ds-1, max_row=r)
    cats     = Reference(ws, min_col=1, min_row=ds, max_row=r)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, f"A{r+3}")
    xl_auto_width(ws)
    ws.freeze_panes = "A5"

def xl_sheet_etapas(wb, week_label, stage_df):
    ws = wb.create_sheet("📋 Etapas — actividad")
    xl_header(ws, "Actividad por etapa", f"Semana: {week_label}", 3)
    r = 4
    for ci, h in enumerate(["Etapa","Total negocios","Con actividad esta semana"], 1):
        ws.cell(r, ci, h)
    xl_style_headers(ws, r, 3)
    ds = r+1
    for _, row in stage_df.iterrows():
        r += 1
        ws.cell(r, 1, row.get("stage_label", row["dealstage"]))
        ws.cell(r, 2, int(row["total_deals"]))
        ws.cell(r, 3, int(row["active_this_week"]))
    xl_style_data(ws, ds, r, 3)
    xl_auto_width(ws)
    ws.freeze_panes = "A5"

def xl_sheet_cambios(wb, week_label, trans_df):
    ws = wb.create_sheet("🔄 Cambios de etapa")
    xl_header(ws, "Cambios de etapa detectados", f"Semana: {week_label}", 7)
    r = 4
    if trans_df.empty:
        ws.cell(r, 1, "No se detectaron cambios de etapa en este período.").font = NORMAL_FONT
        return
    for ci, h in enumerate(["Negocio","ID","Comercial","Pipeline","Etapa anterior","Etapa nueva","Fecha"], 1):
        ws.cell(r, ci, h)
    xl_style_headers(ws, r, 7)
    ds = r+1
    for _, row in trans_df.iterrows():
        r += 1
        ws.cell(r, 1, row.get("dealname",""))
        ws.cell(r, 2, row.get("deal_id",""))
        ws.cell(r, 3, row.get("owner",""))
        ws.cell(r, 4, row.get("pipeline",""))
        ws.cell(r, 5, row.get("from_stage",""))
        ws.cell(r, 6, row.get("to_stage",""))
        ts = row.get("changed_at")
        ws.cell(r, 7, str(ts)[:16] if ts else "")
    xl_style_data(ws, ds, r, 7)
    xl_auto_width(ws)
    ws.freeze_panes = "A5"

def xl_sheet_nuevos(wb, week_label, new_df):
    ws = wb.create_sheet("✨ Nuevos negocios")
    xl_header(ws, "Negocios creados esta semana", f"Semana: {week_label}", 2)
    r = 4
    if new_df.empty:
        ws.cell(r, 1, "No se crearon negocios nuevos en este período.").font = NORMAL_FONT
        return
    for ci, h in enumerate(["Etapa inicial","Negocios nuevos"], 1):
        ws.cell(r, ci, h)
    xl_style_headers(ws, r, 2)
    ds = r+1
    for _, row in new_df.iterrows():
        r += 1
        ws.cell(r, 1, row.get("stage_label", row["dealstage"]))
        ws.cell(r, 2, int(row["new_deals"]))
    xl_style_data(ws, ds, r, 2)
    xl_auto_width(ws)
    ws.freeze_panes = "A5"

def xl_sheet_estancados(wb, stale_data):
    ws = wb.create_sheet("⚠️ Estancados")
    xl_header(ws, "Negocios sin actividad", "Clasificados por días de inactividad", 5)
    r = 4
    for days, df_stale in stale_data.items():
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.cell(r, 1, f"SIN ACTIVIDAD MÁS DE {days} DÍAS ({len(df_stale)} negocios)").fill = SUBHEAD_FILL
        ws.cell(r, 1).font = SUBHEAD_FONT
        r += 1
        if df_stale.empty:
            ws.cell(r, 1, "Ninguno").font = NORMAL_FONT
            r += 2
            continue
        for ci, h in enumerate(["Negocio","Comercial","Etapa actual","Pipeline","Días sin actividad"], 1):
            ws.cell(r, ci, h)
        xl_style_headers(ws, r, 5)
        ds = r+1
        for _, row in df_stale.head(50).iterrows():
            r += 1
            ws.cell(r, 1, row.get("dealname",""))
            ws.cell(r, 2, row.get("owner",""))
            ws.cell(r, 3, row.get("stage_label",""))
            ws.cell(r, 4, row.get("pipeline",""))
            d = int(row.get("days_inactive", 0))
            ws.cell(r, 5, d)
            ws.cell(r, 5).fill = RED_FILL if d >= 60 else (ORANGE_FILL if d >= 30 else PatternFill())
        xl_style_data(ws, ds, r, 5)
        r += 2
    xl_auto_width(ws)
    ws.freeze_panes = "A5"

def xl_sheet_matriz(wb, matrix_df):
    ws = wb.create_sheet("🗃️ Matriz comercial-etapa")
    xl_header(ws, "Distribución total por comercial y etapa", "Todos los negocios", len(matrix_df.columns))
    r = 4
    for ci, col in enumerate(matrix_df.columns, 1):
        ws.cell(r, ci, col)
    xl_style_headers(ws, r, len(matrix_df.columns))
    for _, row in matrix_df.iterrows():
        r += 1
        for ci, val in enumerate(row, 1):
            cell = ws.cell(r, ci, val)
            cell.font      = BOLD_FONT if (ci == 1 or ci == len(row)) else NORMAL_FONT
            cell.border    = THIN_BORDER
            cell.alignment = Alignment(horizontal="center" if ci > 1 else "left", vertical="center")
            if ci == len(row):
                cell.fill = ACCENT_FILL
    xl_auto_width(ws, min_w=8, max_w=30)
    ws.freeze_panes = "B5"

def xl_sheet_comparativa(wb, snap_a_name, snap_b_name, dfa, dfb, act_a, act_b):
    """Pestaña única de comparativa con tabla unificada y tendencia."""

    ids_a = set(dfa["deal_id"])
    ids_b = set(dfb["deal_id"])

    # Tabla unificada
    all_df = pd.concat([dfa, dfb]).drop_duplicates(subset="deal_id").copy()
    all_df["dealstage"] = all_df["dealstage"].astype(str).str.replace(r"\.0$", "", regex=True)
    all_df["stage_label"] = all_df["dealstage"].map(STAGE_LABELS).fillna(all_df["dealstage"])
    all_df[snap_a_name[:15]] = all_df["deal_id"].apply(lambda x: "SI" if x in ids_a else "NO")
    all_df[snap_b_name[:15]] = all_df["deal_id"].apply(lambda x: "SI" if x in ids_b else "NO")
    all_df["Tendencia"] = all_df["deal_id"].apply(
        lambda x: "Sube"     if x in ids_b and x not in ids_a
             else "Baja"     if x in ids_a and x not in ids_b
             else "Mantiene"
    )
    all_df = all_df.sort_values("Tendencia")

    # ── Pestaña 1: Tabla unificada ────────────
    ws1 = wb.create_sheet("📊 Comparativa negocios")
    xl_header(ws1, f"Comparativa: {snap_a_name} vs {snap_b_name}",
              f"Sube: {len(ids_b-ids_a)} | Baja: {len(ids_a-ids_b)} | Mantiene: {len(ids_a&ids_b)}", 7)
    r = 4
    headers = ["Negocio","Comercial","Etapa","Pipeline", snap_a_name[:15], snap_b_name[:15], "Tendencia"]
    for ci, h in enumerate(headers, 1):
        ws1.cell(r, ci, h)
    xl_style_headers(ws1, r, 7)
    ds = r + 1
    for _, row in all_df.iterrows():
        r += 1
        ws1.cell(r, 1, row.get("dealname",""))
        ws1.cell(r, 2, row.get("owner",""))
        ws1.cell(r, 3, row.get("stage_label",""))
        ws1.cell(r, 4, row.get("pipeline",""))
        ws1.cell(r, 5, row[snap_a_name[:15]])
        ws1.cell(r, 6, row[snap_b_name[:15]])
        tend = row["Tendencia"]
        ws1.cell(r, 7, tend)
        if tend == "Sube":
            ws1.cell(r, 7).fill = GREEN_FILL
        elif tend == "Baja":
            ws1.cell(r, 7).fill = RED_FILL
        else:
            ws1.cell(r, 7).fill = ACCENT_FILL
    xl_style_data(ws1, ds, r, 7)
    xl_auto_width(ws1)
    ws1.freeze_panes = "A5"

    # ── Pestaña 2: Diferencia por comercial ──
    ws2 = wb.create_sheet("👤 Diferencia comercial")
    xl_header(ws2, "Comparativa de actividad por comercial", f"{snap_a_name} vs {snap_b_name}", 4)
    r = 4
    for ci, h in enumerate(["Comercial", snap_a_name[:20], snap_b_name[:20], "Diferencia"], 1):
        ws2.cell(r, ci, h)
    xl_style_headers(ws2, r, 4)
    ds = r + 1
    comp = act_a.merge(act_b, on="owner", how="outer", suffixes=("_a","_b")).fillna(0)
    comp["active_this_week_a"] = comp["active_this_week_a"].astype(int)
    comp["active_this_week_b"] = comp["active_this_week_b"].astype(int)
    comp["diff"] = comp["active_this_week_b"] - comp["active_this_week_a"]
    comp = comp.sort_values("active_this_week_b", ascending=False)
    for _, row in comp.iterrows():
        r += 1
        d = int(row["diff"])
        ws2.cell(r, 1, row["owner"])
        ws2.cell(r, 2, int(row["active_this_week_a"]))
        ws2.cell(r, 3, int(row["active_this_week_b"]))
        ws2.cell(r, 4, f"+{d}" if d > 0 else str(d))
        ws2.cell(r, 4).fill = GREEN_FILL if d > 0 else (RED_FILL if d < 0 else PatternFill())
    xl_style_data(ws2, ds, r, 4)
    xl_auto_width(ws2)
    ws2.freeze_panes = "A5"


def generate_excel(week_label, act_df, stage_df, trans_df, new_df, matrix_df, stale_data,
                   comparativa=None):
    wb = Workbook()
    wb.remove(wb.active)
    xl_sheet_resumen(wb, week_label, act_df, stage_df, trans_df, new_df, stale_data)
    xl_sheet_actividad(wb, week_label, act_df)
    xl_sheet_etapas(wb, week_label, stage_df)
    xl_sheet_cambios(wb, week_label, trans_df)
    xl_sheet_nuevos(wb, week_label, new_df)
    xl_sheet_estancados(wb, stale_data)
    xl_sheet_matriz(wb, matrix_df)
    if comparativa:
        xl_sheet_comparativa(wb, **comparativa)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ─────────────────────────────────────────────
# HELPERS DE VISUALIZACIÓN
# ─────────────────────────────────────────────

# Contador global para claves únicas
_table_counter = [0]

def show_deal_table(df, title="Negocios", eng_types=None):
    """Muestra tabla de negocios con filtros por columna y buscador."""
    if df.empty:
        st.info("No hay negocios que mostrar.")
        return
    df2 = df.copy()
    df2["stage_label"]      = df2["dealstage"].map(STAGE_LABELS).fillna(df2["dealstage"])
    df2["Última actividad"] = df2["notes_last_updated"].dt.strftime("%d/%m/%Y %H:%M").fillna("—")

    ACT_EMOJI = {
        "Reunión": "📅 Reunión",
        "Email":   "📧 Email",
        "Llamada": "📞 Llamada",
        "Nota":    "📝 Nota",
        "Tarea":   "✅ Tarea",
        "—":       "—",
    }
    if eng_types:
        df2["Tipo actividad"] = df2["deal_id"].apply(
            lambda x: ACT_EMOJI.get(eng_types.get(str(x), "—"), "—")
        )
        cols   = ["dealname","owner","stage_label","pipeline","Tipo actividad","Última actividad"]
        rename = {"dealname":"Negocio","owner":"Comercial","stage_label":"Etapa","pipeline":"Pipeline"}
    else:
        cols   = ["dealname","owner","stage_label","pipeline","Última actividad"]
        rename = {"dealname":"Negocio","owner":"Comercial","stage_label":"Etapa","pipeline":"Pipeline"}

    show = df2[cols].rename(columns=rename)

    # ── Filtros ──────────────────────────────────
    # Contador incremental para garantizar claves únicas
    _table_counter[0] += 1
    uid = str(_table_counter[0])
    # Se usa un checkbox en vez de st.expander porque esta tabla puede estar
    # dentro de otro expander, y Streamlit no permite anidar expanders.
    ver_filtros = st.checkbox("🔍 Filtros", value=False, key=f"verfiltros_{uid}")
    if ver_filtros:
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            buscar = st.text_input("🔎 Buscar negocio", key=f"buscar_{uid}", placeholder="Escribe para buscar...")
        with fc2:
            owners_opts = sorted(show["Comercial"].dropna().unique().tolist())
            owner_sel   = st.multiselect("👤 Comercial", options=owners_opts, key=f"owner_{uid}", placeholder="Todos")
        with fc3:
            etapas_opts = sorted(show["Etapa"].dropna().unique().tolist())
            etapa_sel   = st.multiselect("📋 Etapa", options=etapas_opts, key=f"etapa_{uid}", placeholder="Todas")

        fc4, fc5 = st.columns(2)
        with fc4:
            pipe_opts = sorted(show["Pipeline"].dropna().unique().tolist())
            pipe_sel  = st.multiselect("🏗️ Pipeline", options=pipe_opts, key=f"pipe_{uid}", placeholder="Todos")
        with fc5:
            if "Tipo actividad" in show.columns:
                tipo_opts = sorted(show["Tipo actividad"].dropna().unique().tolist())
                tipo_sel  = st.multiselect("⚡ Tipo actividad", options=tipo_opts, key=f"tipo_{uid}", placeholder="Todos")
            else:
                tipo_sel = []
    else:
        # Filtros ocultos → valores por defecto (sin filtrar)
        buscar, owner_sel, etapa_sel, pipe_sel, tipo_sel = "", [], [], [], []

    # Aplicar filtros
    filtered = show.copy()
    if buscar:
        filtered = filtered[filtered["Negocio"].str.contains(buscar, case=False, na=False)]
    if owner_sel:
        filtered = filtered[filtered["Comercial"].isin(owner_sel)]
    if etapa_sel:
        filtered = filtered[filtered["Etapa"].isin(etapa_sel)]
    if pipe_sel:
        filtered = filtered[filtered["Pipeline"].isin(pipe_sel)]
    if tipo_sel and "Tipo actividad" in filtered.columns:
        filtered = filtered[filtered["Tipo actividad"].isin(tipo_sel)]

    st.caption(f"**{len(filtered)}** negocios" + (f" (de {len(show)} totales)" if len(filtered) != len(show) else ""))
    st.dataframe(
        filtered,
        hide_index=True,
        use_container_width=True,
        height=min(450, 55 + len(filtered) * 35),
        column_config={
            "Negocio":         st.column_config.TextColumn("Negocio", width="large"),
            "Comercial":       st.column_config.TextColumn("Comercial"),
            "Etapa":           st.column_config.TextColumn("Etapa", width="large"),
            "Pipeline":        st.column_config.TextColumn("Pipeline", width="small"),
            "Última actividad":st.column_config.TextColumn("Última actividad"),
        }
    )

def semaforo_color(rate):
    if rate >= 20:   return "#1D9E75"
    elif rate >= 10: return "#F4A835"
    else:            return "#E05252"

def semaforo_emoji(rate):
    if rate >= 20:   return "🟢"
    elif rate >= 10: return "🟡"
    else:            return "🔴"

PLOT_LAYOUT = dict(
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Arial", size=12),
)

# ─────────────────────────────────────────────
# APP STREAMLIT
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="Pipeline Tracker — HubSpot",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 1.8rem; font-weight: 700; }
[data-testid="stMetricLabel"] { font-size: 0.85rem; color: #666; }
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
div[data-testid="metric-container"] {
    background: #f8f9fb;
    border: 1px solid #e8eaf0;
    border-radius: 10px;
    padding: 16px;
}
</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=3600, show_spinner=False)
def cached_pipelines(api_key):
    return get_pipelines(api_key)

@st.cache_data(ttl=3600, show_spinner=False)
def cached_all_deals(api_key, stage_ids=(), stage_map_json="{}"):
    return parse_deals(get_all_deals(api_key, stage_ids), json.loads(stage_map_json))

@st.cache_data(ttl=3600, show_spinner=False)
def cached_stage_history(api_key, deal_ids_tuple):
    return get_stage_history_bulk(api_key, list(deal_ids_tuple))

@st.cache_data(ttl=3600, show_spinner=False)
def cached_real_activity(api_key, since_ts_ms, until_ts_ms):
    return get_real_activity(api_key, since_ts_ms, until_ts_ms)

@st.cache_data(ttl=3600, show_spinner=False)
def cached_meetings(api_key, since_ts_ms, until_ts_ms):
    return get_meetings(api_key, since_ts_ms, until_ts_ms)

@st.cache_data(ttl=3600, show_spinner=False)
def cached_engagement_types(api_key, deal_ids_tuple, notes_ts_tuple=None):
    notes_map = dict(notes_ts_tuple) if notes_ts_tuple else None
    return get_engagement_types(api_key, list(deal_ids_tuple), notes_map)

# ── Snapshots persistentes ────────────────────
if "snapshots" not in st.session_state:
    st.session_state.snapshots = load_snapshots_from_file()
    st.session_state["ftp_ok"] = len(st.session_state.snapshots) >= 0  # marca que intentó cargar

# ── Barra lateral ──────────────────────────────

api_key = HUBSPOT_API_KEY

st.sidebar.title("📊 Pipeline Tracker")
st.sidebar.caption("Seguimiento semanal de HubSpot")

if st.sidebar.button("🔄 Actualizar datos de HubSpot", use_container_width=True,
                     help="Descarta la caché y vuelve a descargar todos los datos. La carga tardará más."):
    st.cache_data.clear()
    st.rerun()

st.sidebar.divider()

st.sidebar.subheader("Período de análisis")
col1, col2 = st.sidebar.columns(2)
with col1:
    start_date = st.date_input("Desde", value=ahora_espana().date() - timedelta(days=7))
with col2:
    end_date = st.date_input("Hasta", value=ahora_espana().date())

st.sidebar.divider()

load_history = st.sidebar.checkbox(
    "Cargar historial de etapas", value=True,
    help="Necesario para ver cambios de etapa.",
)
load_real_activity = st.sidebar.checkbox(
    "Desglose por tipo de actividad", value=False,
    help="Carga llamadas, emails, reuniones, notas y tareas. Más lento.",
)
pipeline_filter = st.sidebar.multiselect(
    "Filtrar pipeline", options=list(PIPELINE_NAMES.values()),
    default=list(PIPELINE_NAMES.values()),
)
owner_filter = st.sidebar.multiselect(
    "Filtrar comercial", options=sorted(OWNER_NAMES.values()),
    default=[], placeholder="Todos",
)
incluir_bbdd = st.sidebar.checkbox(
    "Incluir BBDD en actividad y estancados", value=False,
    help="La etapa BBDD es la reserva de empresas sin trabajar. Por defecto no cuenta "
         "como negocio abierto para que no distorsione la tasa de actividad ni los estancados.",
)

# Cupo fijo de momento (versión sencilla: sin ajustes en la barra lateral)
cupo_comercial    = CUPO_POR_DEFECTO
semanas_caducidad = SEMANAS_CADUCIDAD_DEF

st.sidebar.divider()
st.sidebar.subheader("📸 Snapshots semanales")
if FTP_HOST and FTP_USER:
    if st.sidebar.button("🔌 Probar conexión FTP", use_container_width=True):
        try:
            import ftplib as _ftplib
            with _ftplib.FTP(timeout=5) as _ftp:
                _ftp.connect(FTP_HOST, timeout=5)
                _ftp.login(FTP_USER, FTP_PASS)
                st.sidebar.success(f"✅ FTP conectado: {FTP_HOST}")
        except Exception as _e:
            st.sidebar.error(f"❌ FTP error: {_e}")
snapshot_name = st.sidebar.text_input("Nombre del snapshot", value=week_label if 'week_label' in dir() else "", placeholder="Ej: Semana 28/05 — 04/06")
if st.sidebar.button("💾 Guardar snapshot actual", use_container_width=True):
    st.session_state["pending_snapshot"] = snapshot_name

if st.session_state.get("ftp_load_error"):
    st.sidebar.warning(f"⚠️ Error al cargar FTP: {st.session_state['ftp_load_error']}")

if st.session_state.snapshots:
    st.sidebar.caption(f"📁 {len(st.session_state.snapshots)} snapshot(s) guardado(s):")
    for sn in list(st.session_state.snapshots.keys()):
        col_s1, col_s2 = st.sidebar.columns([3, 1])
        col_s1.caption(f"• {sn}")
        if col_s2.button("🗑", key=f"del_{sn}", help=f"Eliminar {sn}"):
            st.session_state.snapshots = delete_snapshot_from_file(sn, st.session_state.snapshots)
            st.rerun()

# ── Ver snapshot guardado ─────────────────────
if st.session_state.snapshots:
    st.sidebar.divider()
    snap_keys_all = ["— Ver período actual —"] + list(st.session_state.snapshots.keys())
    snap_ver = st.sidebar.selectbox("👁️ Visualizar snapshot:", options=snap_keys_all, key="snap_ver")
    if snap_ver != "— Ver período actual —":
        st.session_state["modo_snapshot"] = snap_ver
    else:
        st.session_state["modo_snapshot"] = None

st.sidebar.divider()
st.sidebar.caption("🟢 Actividad ≥20%  🟡 10-20%  🔴 <10%")
st.sidebar.caption("Actividad = notes_last_updated. Tipos: reuniones, emails, llamadas, notas y tareas.")

# ── Login ──────────────────────────────────────

if APP_PASSWORD:
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if not st.session_state.authenticated:
        st.markdown("<br>" * 3, unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 1.5, 1])
        with col2:
            st.markdown("""
            <div style="text-align:center; margin-bottom:24px;">
                <div style="font-size:2.5rem">📊</div>
                <div style="font-size:1.4rem; font-weight:700; color:#1A3A5C">Pipeline Tracker</div>
                <div style="font-size:13px; color:#888; margin-top:4px">Zentralcom · HubSpot</div>
            </div>
            """, unsafe_allow_html=True)
            pwd = st.text_input("Contraseña", type="password", placeholder="Introduce la contraseña...")
            if st.button("Entrar", use_container_width=True):
                if pwd == APP_PASSWORD:
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta")
        st.stop()

# ── Comprobación API ───────────────────────────

if not api_key or api_key.startswith("pat-eu1-xxxxxxxx"):
    st.error("⚠️ Pon tu API Key real en la variable HUBSPOT_API_KEY al inicio del código.")
    st.stop()

if not test_connection(api_key):
    st.error("❌ No se pudo conectar con HubSpot. Revisa tu API Key.")
    st.stop()

# ── Carga de datos ─────────────────────────────

week_start = datetime.combine(start_date, datetime.min.time())
week_end   = datetime.combine(end_date,   datetime.max.time())
week_label = f"{start_date.strftime('%d/%m/%Y')} — {end_date.strftime('%d/%m/%Y')}"
today      = ahora_espana()

with st.spinner("Cargando negocios de HubSpot..."):
    _pipes = cached_pipelines(api_key)
    resolver_etapas(_pipes)
    _stage_ids = tuple(sorted({sid for m in STAGE_MAP.values() for sid in m.values() if sid}))
    df_all = cached_all_deals(api_key, _stage_ids, json.dumps(STAGE_MAP, sort_keys=True))

with st.spinner("Calculando fechas de entrada en las etapas nuevas..."):
    # Solo hace falta el historial de los negocios tocados desde la reorganización
    _desde = pd.Timestamp(min(c["migracion"] for c in PIPE_CONFIG.values())) - pd.Timedelta(days=1)
    _ids_hist = df_all.loc[
        df_all["pipeline_id"].isin(STAGE_MAP.keys())
        & df_all[DATE_COLS].isna().any(axis=1)
        & (df_all["last_modified"] >= _desde),
        "deal_id"].astype(str)
    _hist_fechas = cached_stage_history(api_key, tuple(sorted(_ids_hist))) if len(_ids_hist) else {}
    df_all = completar_fechas_etapas(df_all, STAGE_MAP, _hist_fechas)

if owner_filter:
    df_all = df_all[df_all["owner"].isin(owner_filter)]

# Todos los negocios (abiertos y cerrados) de todos los pipelines — para la pestaña "Esta semana",
# que tiene su propio selector de pipeline (no depende del filtro de la barra lateral)
df_kpi_base = df_all.copy()
df_lead = df_kpi_base[df_kpi_base["pipeline_id"] == LEAD_PIPELINE_ID].copy()

if pipeline_filter:
    df_all = df_all[df_all["pipeline"].isin(pipeline_filter)]

# A partir de aquí todo el dashboard trabaja SOLO con negocios abiertos
df_closed = df_all[df_all["dealstage"].isin(CLOSED_STAGES)].copy()  # guardamos cerrados solo para cierres KPI
df_all    = df_all[~df_all["dealstage"].isin(CLOSED_STAGES)].copy()
n_bbdd    = int(df_all["dealstage"].isin(BBDD_IDS).sum())
if not incluir_bbdd:
    # BBDD = reserva de empresas sin trabajar → fuera de actividad y estancados
    df_all = df_all[~df_all["dealstage"].isin(BBDD_IDS)].copy()

# df_week basado en notes_last_updated — misma propiedad que usa HubSpot UI
df_week = df_all[
    df_all["notes_last_updated"].notna() &
    (df_all["notes_last_updated"] >= week_start) &
    (df_all["notes_last_updated"] <= week_end)
].copy()

trans_df = pd.DataFrame(columns=["deal_id","dealname","owner","pipeline","from_stage","to_stage","changed_at"])
if load_history and not df_week.empty:
    with st.spinner(f"Cargando historial de etapas para {len(df_week)} negocios..."):
        history_data = cached_stage_history(api_key, tuple(df_week["deal_id"].tolist()))
    trans_df = stage_transitions(history_data, week_start, week_end)

activity_detail_df = pd.DataFrame(columns=["owner","tipo"])
engagement_types   = {}
if not df_week.empty:
    # Fecha real de "Última actividad" de cada negocio (notes_last_updated,
    # convertida a epoch ms en UTC) — se usa para saber si esa fecha
    # corresponde de verdad al vencimiento de una tarea.
    notes_ts_map = {}
    for _, row in df_week.iterrows():
        nlu = row["notes_last_updated"]
        if pd.notna(nlu):
            try:
                notes_ts_map[str(row["deal_id"])] = int(pd.Timestamp(nlu).tz_localize(TZ_ESPANA).timestamp() * 1000)
            except Exception:
                pass
    notes_ts_tuple = tuple(sorted(notes_ts_map.items()))
    with st.spinner(f"Cargando tipo de actividad para {len(df_week)} negocios..."):
        engagement_types = cached_engagement_types(api_key, tuple(df_week["deal_id"].tolist()), notes_ts_tuple)
    # Aviso técnico temporal: si HubSpot ha rechazado alguna petición (por
    # ejemplo por falta de permisos en la API Key para leer Emails o
    # Llamadas), se enseña aquí el motivo exacto para poder diagnosticarlo.
    _debug_errors = st.session_state.get("_engagement_debug_errors")
    if _debug_errors:
        with st.expander("⚠️ Aviso técnico: algún tipo de actividad no se ha podido leer (pulsa para ver el detalle)"):
            for _tipo, _msg in _debug_errors.items():
                st.write(f"**{_tipo}**: {_msg}")
    rows = []
    for _, row in df_week.iterrows():
        tipo = engagement_types.get(str(row["deal_id"]), "—")
        if tipo != "—":
            rows.append({"owner": row["owner"], "tipo": tipo})
    activity_detail_df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=["owner","tipo"])

act_df     = activity_by_owner(df_all, week_start, week_end)
# ── Modo visualización de snapshot ────────────
modo_snap = st.session_state.get("modo_snapshot")
if modo_snap and modo_snap in st.session_state.snapshots:
    snap_data  = st.session_state.snapshots[modo_snap]
    df_week    = snap_data["df_week"]
    if snap_data.get("act_df") is not None:
        act_df = snap_data["act_df"]
    # Tipos de actividad congelados al guardar el snapshot.
    # Snapshots antiguos no los guardaban → la columna no se mostrará.
    engagement_types = snap_data.get("eng_types") or {}
    week_label = f"📸 SNAPSHOT: {modo_snap}"

# Guardar snapshot si se solicitó
if st.session_state.get("pending_snapshot"):
    name = st.session_state.pop("pending_snapshot")
    st.session_state.snapshots[name] = {
        "week_label": week_label,
        "df_week":    df_week.copy(),
        "act_df":     act_df.copy(),
        "eng_types":  dict(engagement_types),
    }
    result = save_snapshots_to_file(st.session_state.snapshots)
    if result is True:
        if FTP_HOST and FTP_USER:
            st.sidebar.success(f"✅ Guardado en FTP: {name}")
        else:
            st.sidebar.success(f"✅ Guardado en disco: {name}")
    else:
        st.sidebar.error(f"❌ Error al guardar: {result}")
        st.sidebar.info("El snapshot queda guardado en memoria hasta que cierres la app.")
stage_df   = stage_activity_summary(df_all, week_start, week_end)
new_df     = new_deals_by_stage(df_all, week_start, week_end)
matrix_df  = owner_stage_matrix(df_all)
stale_data = stale_deals(df_all, STALE_DAYS, today)

# ── CABECERA ───────────────────────────────────

st.title("📊 Pipeline Tracker — HubSpot")
_bbdd_txt = "" if incluir_bbdd else f" (sin contar {n_bbdd:,} en BBDD)"
st.caption(f"Período: **{week_label}** · {len(df_all):,} negocios abiertos{_bbdd_txt} · Actualizado: {today.strftime('%d/%m/%Y %H:%M')}")

# ── KPIs ───────────────────────────────────────

total      = len(df_all)  # df_all ya es solo abiertos
active     = int(act_df["active_this_week"].sum())
rate       = round(active/total*100, 1) if total else 0
# Cierres = negocios cerrados con fecha de cierre en el período (usando df_closed)
won  = len(df_closed[
    df_closed["dealstage"].isin(WON_STAGES) &
    df_closed["closedate"].notna() &
    (df_closed["closedate"] >= week_start) &
    (df_closed["closedate"] <= week_end)
])
lost = len(df_closed[
    df_closed["dealstage"].isin(LOST_STAGES) &
    df_closed["closedate"].notna() &
    (df_closed["closedate"] >= week_start) &
    (df_closed["closedate"] <= week_end)
])
new_total = int(new_df["new_deals"].sum()) if not new_df.empty else 0
stale14   = len(stale_data.get(14, pd.DataFrame()))
stale30   = len(stale_data.get(30, pd.DataFrame()))

c = st.columns(7)
# KPIs clickables
kpi_data = [
    ("Negocios abiertos",   f"{total:,}",          None,                     "total"),
    ("Con actividad",       f"{active:,}",          f"{rate}% del total",     "activos"),
    ("Tasa actividad",      f"{rate}%",             None,                     "activos"),
    ("Cambios de etapa",    f"{len(trans_df):,}",   None,                     "cambios"),
    ("Nuevos en período",   f"{new_total:,}",        None,                    "nuevos"),
    ("Cierres ganados",     f"{won:,}",             f"{lost} perdidos en período","ganados"),
    ("Estancados >14d",     f"{stale14:,}",         f"De {total} abiertos · {stale30} >30d", "estancados"),
]
for col, (label, val, delta, key) in zip(c, kpi_data):
    col.metric(label, val, delta)

# Desplegables KPIs
nuevos_df = df_all[(df_all["createdate"] >= week_start) & (df_all["createdate"] <= week_end)]
ganados_df = df_closed[
    df_closed["dealstage"].isin(WON_STAGES) &
    df_closed["closedate"].notna() &
    (df_closed["closedate"] >= week_start) &
    (df_closed["closedate"] <= week_end)
]
perdidos_df = df_closed[
    df_closed["dealstage"].isin(LOST_STAGES) &
    df_closed["closedate"].notna() &
    (df_closed["closedate"] >= week_start) &
    (df_closed["closedate"] <= week_end)
]

with st.expander(f"🔵 **Negocios abiertos** — {total:,} en el pipeline"):
    show_deal_table(df_all)

with st.expander(f"✅ **Con actividad** — {active:,} en el período seleccionado"):
    show_deal_table(df_week, eng_types=engagement_types)

with st.expander(f"✨ **Nuevos en período** — {new_total:,} creados"):
    show_deal_table(nuevos_df)

with st.expander(f"🏆 **Cierres** — {won:,} ganados / {lost:,} perdidos en el período"):
    if not ganados_df.empty:
        st.markdown("**✅ Ganados**")
        show_deal_table(ganados_df, eng_types=engagement_types)
    if not perdidos_df.empty:
        st.markdown("**❌ Perdidos**")
        show_deal_table(perdidos_df, eng_types=engagement_types)
    if ganados_df.empty and perdidos_df.empty:
        st.info("No hay cierres en el período seleccionado.")

with st.expander(f"⚠️ **Estancados >14 días** — {stale14:,} negocios"):
    show_deal_table(stale_data.get(14, pd.DataFrame()))

st.divider()

# ── PESTAÑAS ───────────────────────────────────

tabs = st.tabs([
    "📋 Esta semana",
    "👤 Comerciales",
    "📋 Etapas",
    "🔄 Cambios de etapa",
    "✨ Nuevos negocios",
    "⚠️ Estancados",
    "🌡️ Mapa de calor",
    "📊 Comparativa semanal",
])

# ── TAB 0: ESTA SEMANA (versión sencilla) ──────

# Los tres KPIs (iguales en LEAD y CLIENTES; las etapas exactas están en PIPE_CONFIG):
#   Visitas      = el negocio entra en la etapa Reunión
#   Presupuestos = el negocio entra en Facturas recibidas
#   Clientes     = el negocio entra en ganado (Cierre ganado en LEAD, Negocio ganado en CLIENTES)
# Cada negocio cuenta una sola vez, en la semana en que entra en la etapa.
KPIS_SEMANA = ["Visitas", "Presupuestos", "Clientes"]

with tabs[0]:
    lunes_hoy   = (pd.Timestamp(today) - pd.Timedelta(days=today.weekday())).normalize()
    semana_hoy  = etiqueta_semana(lunes_hoy)
    st.subheader(f"📋 Esta semana — {semana_hoy} (desde el lunes {lunes_hoy.strftime('%d/%m')})")

    pipe_nombre = st.radio("Pipeline", ["LEAD", "CLIENTES"], horizontal=True, key="esta_pipe")
    pipe_id   = LEAD_PIPELINE_ID if pipe_nombre == "LEAD" else "default"
    ids_pipe  = STAGE_MAP.get(pipe_id, {})
    df_tab    = df_kpi_base[df_kpi_base["pipeline_id"] == pipe_id].copy()
    nom_etapa = lambda col: STAGE_LABELS.get(ids_pipe.get(col) or "", "—")
    faltan    = [n for col, n in [("f_seleccion", "Selección"), ("f_reunion", "Reunión"),
                                  ("f_presupuesto", "Facturas recibidas"), ("f_ganado", "ganado")]
                 if not ids_pipe.get(col)]
    if faltan:
        st.warning(f"No encuentro en HubSpot estas etapas del pipeline {pipe_nombre}: {', '.join(faltan)}. "
                   f"Revisa que se llamen así o avisa para ajustarlo.")
    if st.session_state.get("_pipelines_error"):
        st.warning(f"No se han podido leer los pipelines de HubSpot: {st.session_state['_pipelines_error']}")

    if df_tab.empty:
        st.info(f"No hay negocios del pipeline {pipe_nombre} con los filtros actuales.")
    else:
        eventos = kpi_eventos(df_tab)   # visitas, presupuestos y clientes con su fecha

        # ── Tabla 1: cómo va cada comercial ──
        st.markdown("##### ¿Cómo va cada comercial?")

        # Selector de semana (la semana en curso por defecto; hasta 8 semanas atrás)
        semanas_sel = periodos_kpi("Semana", 8, today)[::-1]
        def _nombre_sem(i):
            etq, ini, fin = semanas_sel[i]
            txt = f"{etq.split(' · ')[0]} ({ini.strftime('%d/%m')} – {fin.strftime('%d/%m')})"
            return txt + " · en curso" if i == 0 else txt
        i_sem = st.selectbox("Semana", list(range(len(semanas_sel))), format_func=_nombre_sem,
                             key=f"sem_t1_{pipe_nombre}")
        es_actual = i_sem == 0
        sem_sel   = semanas_sel[i_sem][0].split(" · ")[0]
        ini_sel, fin_sel = semanas_sel[i_sem][1], semanas_sel[i_sem][2]
        ev_semana = eventos[(eventos["fecha"] >= ini_sel) & (eventos["fecha"] <= fin_sel)]
        cupo_df, _sel, _reu, _desc = cupo_semanal(
            df_tab, lunes_hoy.to_pydatetime(), today, today, cupo_comercial, semanas_caducidad
        )
        if es_actual:
            en_sel = dict(zip(cupo_df["owner"], cupo_df["en_seleccion"])) if not cupo_df.empty else {}
            sel_cierre = pd.DataFrame()
        else:
            # Semana pasada: cuántos tenía cada comercial en Selección el domingo a las 23:59,
            # reconstruido con el historial de etapas de los negocios que han pasado por Selección
            candidatos = df_tab[df_tab["f_seleccion"].notna()]
            ids_cierre = set()
            if not candidatos.empty:
                with st.spinner("Reconstruyendo cómo estaba cada comercial al cierre de la semana..."):
                    hist = cached_stage_history(api_key, tuple(sorted(candidatos["deal_id"].astype(str))))
                ids_cierre = en_seleccion_al_cierre(hist, SELECCION_IDS, fin_sel)
            sel_cierre = df_tab[df_tab["deal_id"].astype(str).isin(ids_cierre)]
            en_sel = sel_cierre.groupby("owner").size().to_dict()

        owners_t1 = sorted(set(en_sel) | set(ev_semana.loc[ev_semana["kpi"].isin(KPIS_SEMANA), "owner"]))

        col_trab  = "Trabajando ahora" if es_actual else "Trabajando al cierre"
        col_huec  = "Huecos libres" if es_actual else "Huecos libres al cierre"
        cols_t1 = ["Comercial", col_trab, "Visitas", "Presupuestos", "Clientes", col_huec]
        filas = []
        for o in owners_t1:
            ev_o = ev_semana[ev_semana["owner"] == o]
            n_sel = int(en_sel.get(o, 0))
            filas.append({
                "Comercial":        o,
                col_trab:           n_sel,
                "Visitas":          int((ev_o["kpi"] == "Visitas").sum()),
                "Presupuestos":     int((ev_o["kpi"] == "Presupuestos").sum()),
                "Clientes":         int((ev_o["kpi"] == "Clientes").sum()),
                col_huec:           max(cupo_comercial - n_sel, 0),
            })
        tabla1 = pd.DataFrame(filas)[cols_t1] if filas else pd.DataFrame(columns=cols_t1)
        if tabla1.empty:
            st.info("Ningún comercial tiene negocios en Selección ni visitas, presupuestos o clientes en esta semana."
                    if es_actual else f"En la {sem_sel} no hubo visitas, presupuestos ni clientes.")
        else:
            total = {"Comercial": "TOTAL"}
            total.update({c: int(tabla1[c].sum()) for c in cols_t1[1:]})
            tabla1 = pd.concat([tabla1, pd.DataFrame([total])], ignore_index=True)
            st.dataframe(
                tabla1.style.apply(
                    lambda fila: ["font-weight: bold" if fila["Comercial"] == "TOTAL" else "" for _ in fila],
                    axis=1,
                ),
                hide_index=True, use_container_width=True,
            )
        if es_actual:
            st.caption(
                f"Visitas, presupuestos y clientes cuentan desde el lunes. "
                f"**Trabajando ahora** = negocios en Selección. "
                f"**Huecos libres** = {cupo_comercial} − trabajando ahora: cuántos negocios nuevos se le pueden dar."
            )
        else:
            st.caption(
                f"Visitas, presupuestos y clientes de la {sem_sel}, de lunes a domingo. "
                f"**Trabajando al cierre** = negocios que estaban en Selección el domingo {fin_sel.strftime('%d/%m')} a las 23:59. "
                f"**Huecos libres al cierre** = {cupo_comercial} − trabajando al cierre."
            )

        # ── Qué negocios tiene cada comercial y por qué cuentan ──
        if owners_t1:
            with st.expander("🔎 Ver los negocios de cada comercial"):
                com = st.selectbox("Comercial", owners_t1, key=f"com_{pipe_nombre}")
                etapa_hoy = dict(zip(df_tab["deal_id"].astype(str), df_tab["dealstage"]))

                # 1) Negocios en Selección: ahora (semana en curso) o al cierre (semana pasada)
                sel_c = _sel[_sel["owner"] == com].sort_values("fecha_seleccion")
                if es_actual:
                    st.markdown(f"**Trabajando ahora ({len(sel_c)})** — están en la etapa Selección")
                else:
                    cierre_c = sel_cierre[sel_cierre["owner"] == com] if not sel_cierre.empty else sel_cierre
                    st.markdown(f"**Trabajando al cierre ({len(cierre_c)})** — estaban en Selección el domingo {fin_sel.strftime('%d/%m')}")
                    if cierre_c.empty:
                        st.caption("No tenía negocios en Selección al cierre de esa semana.")
                    else:
                        st.dataframe(pd.DataFrame({
                            "Negocio":      cierre_c["dealname"],
                            "Etapa actual": cierre_c["dealstage"].map(STAGE_LABELS).fillna("—"),
                        }), hide_index=True, use_container_width=True,
                            height=min(400, 55 + len(cierre_c) * 35))
                if not es_actual:
                    pass
                elif sel_c.empty:
                    st.caption("No tiene negocios en Selección.")
                else:
                    st.dataframe(pd.DataFrame({
                        "Negocio":              sel_c["dealname"],
                        "En Selección desde":   sel_c["fecha_seleccion"].dt.strftime("%d/%m/%Y").fillna("—"),
                        "Semanas en Selección": sel_c["semanas_en_seleccion"],
                        "Última actividad":     sel_c["notes_last_updated"].dt.strftime("%d/%m/%Y").fillna("—"),
                    }), hide_index=True, use_container_width=True,
                        height=min(400, 55 + len(sel_c) * 35))

                # 2) Visitas, presupuestos y clientes de esta semana
                ev_c = ev_semana[(ev_semana["owner"] == com) & ev_semana["kpi"].isin(KPIS_SEMANA)].sort_values("fecha")
                st.markdown(f"**{sem_sel} ({len(ev_c)})** — visitas, presupuestos y clientes"
                            + (" desde el lunes" if es_actual else ""))
                if ev_c.empty:
                    st.caption("No ha sumado visitas, presupuestos ni clientes en esta semana.")
                else:
                    motivo = {
                        "Visitas":      f"Entró en {nom_etapa('f_reunion')}",
                        "Presupuestos": f"Entró en {nom_etapa('f_presupuesto')}",
                        "Clientes":     f"Entró en {nom_etapa('f_ganado')}",
                    }
                    st.dataframe(pd.DataFrame({
                        "KPI":          ev_c["kpi"],
                        "Negocio":      ev_c["detalle"],
                        "Por qué cuenta": ev_c["kpi"].map(motivo) + " el " + ev_c["fecha"].dt.strftime("%d/%m"),
                        "Etapa actual": ev_c["deal_id"].map(etapa_hoy).map(STAGE_LABELS).fillna("—"),
                    }), hide_index=True, use_container_width=True)

        # ── Tabla 2: semana a semana ──
        st.markdown("##### ¿Cómo vamos semana a semana?")
        periodos4 = periodos_kpi("Semana", 4, today)
        tabla2 = tabla_kpis(eventos, periodos4, kpis=KPIS_SEMANA).drop(columns=["Total"])
        tabla2.columns = [c.split(" · ")[0] for c in tabla2.columns]
        st.dataframe(tabla2, use_container_width=True)
        st.caption(
            f"Pipeline {pipe_nombre}: **Visitas** = entran en *{nom_etapa('f_reunion')}* · "
            f"**Presupuestos** = entran en *{nom_etapa('f_presupuesto')}* · "
            f"**Clientes** = entran en *{nom_etapa('f_ganado')}*. "
            f"La columna {semana_hoy} es la semana en curso. "
            f"No cuentan como visitas los negocios movidos en bloque a Reunión el día de la reorganización "
            f"({pd.Timestamp(PIPE_CONFIG[pipe_id]['migracion']).strftime('%d/%m/%Y')})."
        )

        # ── Qué negocios hay detrás de cada número ──
        ini4 = periodos4[0][1]
        det = eventos[(eventos["kpi"].isin(KPIS_SEMANA)) & (eventos["fecha"] >= ini4)].copy()
        etapa_actual = dict(zip(df_tab["deal_id"].astype(str), df_tab["dealstage"]))
        if "deal_id" in det.columns:
            det["etapa"] = det["deal_id"].map(etapa_actual).map(STAGE_LABELS).fillna("—")
        else:
            det["etapa"] = "—"
        det = det.sort_values(["fecha"], ascending=False)
        detalle_df = pd.DataFrame({
            "Semana":       det["semana"],
            "Fecha":        det["fecha"].dt.strftime("%d/%m/%Y"),
            "KPI":          det["kpi"],
            "Negocio":      det["detalle"],
            "Comercial":    det["owner"],
            "Etapa actual": det["etapa"],
        })
        with st.expander(f"🔎 Ver qué negocios son ({len(detalle_df)})"):
            if detalle_df.empty:
                st.info("No hay visitas, presupuestos ni clientes en estas 4 semanas.")
            else:
                f1, f2 = st.columns(2)
                with f1:
                    kpi_f = st.selectbox("KPI", ["Todos"] + KPIS_SEMANA, key=f"det_kpi_{pipe_nombre}")
                with f2:
                    sem_f = st.selectbox("Semana", ["Todas"] + list(tabla2.columns)[::-1], key=f"det_sem_{pipe_nombre}")
                ver = detalle_df
                if kpi_f != "Todos":
                    ver = ver[ver["KPI"] == kpi_f]
                if sem_f != "Todas":
                    ver = ver[ver["Semana"] == sem_f]
                st.caption(f"**{len(ver)}** negocios")
                st.dataframe(ver, hide_index=True, use_container_width=True,
                             height=min(450, 55 + len(ver) * 35))

        st.download_button(
            "⬇️ Descargar en Excel",
            data=excel_semana(tabla1, tabla2, f"{sem_sel} · {pipe_nombre}", today, detalle_df),
            file_name=f"Seguimiento_semanal_{pipe_nombre}_{sem_sel.replace(' ', '')}_{today.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ── Comprobar qué papel tiene cada etapa de HubSpot en la app ──
    with st.expander(f"🔧 Comprobar etapas del pipeline {pipe_nombre}"):
        st.caption("Lista de las etapas que tiene ahora este pipeline en HubSpot y para qué las usa la app. "
                   "Sirve para comprobar que, tras cambiar etapas, cada una cuenta donde debe.")
        papeles = {}
        for col, papel in [("f_bbdd", "BBDD: reserva, fuera de actividad y estancados"),
                           ("f_seleccion", "Selección: cupo (Trabajando / Huecos libres)"),
                           ("f_reunion", "Visitas"),
                           ("f_presupuesto", "Presupuestos"),
                           ("f_ganado", "Clientes")]:
            if ids_pipe.get(col):
                papeles[ids_pipe[col]] = papel
        cuenta = df_kpi_base[df_kpi_base["pipeline_id"] == pipe_id].groupby("dealstage").size()
        filas_d = []
        for sid, nombre, *_ in _pipes.get(pipe_id, []):
            papel = papeles.get(sid, "")
            if sid in WON_STAGES:
                papel = (papel + " · " if papel else "") + "Cerrada (ganada)"
            elif sid in CLOSED_STAGES:
                papel = (papel + " · " if papel else "") + "Cerrada (no cuenta como abierta)"
            filas_d.append({"Etapa en HubSpot": nombre, "Nombre en la app": STAGE_LABELS.get(sid, sid),
                            "Papel en la app": papel or "Abierta", "Negocios ahora": int(cuenta.get(sid, 0)),
                            "ID": sid})
        if filas_d:
            st.dataframe(pd.DataFrame(filas_d), hide_index=True, use_container_width=True)
        else:
            st.warning("No se han podido leer las etapas de este pipeline en HubSpot.")
        ids_hs = {sid for sid, *_ in _pipes.get(pipe_id, [])}
        huerfanos = {sid: n for sid, n in cuenta.items() if ids_hs and sid not in ids_hs}
        if huerfanos:
            st.warning("Hay negocios en etapas que ya no existen en HubSpot: "
                       + ", ".join(f"{STAGE_LABELS.get(k, k)} ({v})" for k, v in huerfanos.items()))

# ── TAB 1: COMERCIALES ─────────────────────────

with tabs[1]:
    st.subheader("Actividad semanal por comercial")

    # Tarjetas semáforo
    cols = st.columns(max(len(act_df), 1))
    for col, (_, row) in zip(cols, act_df.iterrows()):
        color = semaforo_color(row["activity_rate"])
        emoji = semaforo_emoji(row["activity_rate"])
        pct   = row["activity_rate"]
        bg    = "#f0faf6" if pct >= 20 else "#fffbf0" if pct >= 10 else "#fdf2f2"
        with col:
            st.markdown(f"""
            <div style="border:1.5px solid {color}; border-radius:12px; padding:14px 10px;
                        text-align:center; background:{bg}">
                <div style="font-size:1.5rem">{emoji}</div>
                <div style="font-weight:700; font-size:0.9rem; margin:4px 0">{row['owner'].split()[0]}</div>
                <div style="font-size:1.6rem; font-weight:800; color:{color}">{pct}%</div>
                <div style="font-size:0.78rem; color:#888">{row['active_this_week']} / {row['open_deals']} abiertos</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Gráfico barras apiladas
    fig = go.Figure()
    fig.add_bar(
        name="Sin actividad (abiertos)",
        x=act_df["owner"], y=act_df["sin_actividad"],
        marker_color="#E8EAF0", marker_line_width=0,
    )
    fig.add_bar(
        name="Con actividad esta semana",
        x=act_df["owner"], y=act_df["active_this_week"],
        marker_color=[semaforo_color(r) for r in act_df["activity_rate"]],
        marker_line_width=0,
        text=act_df["active_this_week"], textposition="outside",
    )
    fig.update_layout(
        barmode="stack", height=380,
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        **PLOT_LAYOUT,
    )
    fig.update_xaxes(tickangle=-20)
    st.plotly_chart(fig, use_container_width=True)

    # Barras de progreso con expander clickable
    st.markdown("##### Detalle — haz clic en un comercial para ver sus negocios")
    for _, row in act_df.iterrows():
        pct   = row["activity_rate"]
        color = semaforo_color(pct)
        bar_w = min(int(pct * 3), 300)
        emoji = semaforo_emoji(row["activity_rate"])

        with st.expander(
            f"{emoji} **{row['owner']}** — {row['active_this_week']} activos / {row['total_deals']} total ({pct}%)",
            expanded=False
        ):
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown(f"""
                <div style="background:#f8f9fb; border-radius:8px; padding:12px 16px;">
                    <div style="display:flex; align-items:center; gap:10px;">
                        <div style="flex:1; background:#E8EAF0; border-radius:4px; height:14px;">
                            <div style="width:{bar_w}px; max-width:100%; background:{color};
                                        height:14px; border-radius:4px;"></div>
                        </div>
                        <div style="font-weight:700; color:{color}">{pct}%</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            tab_a, tab_b = st.tabs(["✅ Con actividad en el período", "💤 Sin actividad en el período"])
            with tab_a:
                deals_active = df_week[df_week["owner"] == row["owner"]]
                show_deal_table(deals_active, f"Con actividad — {row['owner']}", eng_types=engagement_types)
            with tab_b:
                deals_inactive = df_all[
                    (df_all["owner"] == row["owner"]) &
                    ~(df_all["deal_id"].isin(df_week["deal_id"]))
                ]
                show_deal_table(deals_inactive, f"Sin actividad — {row['owner']}")

    # Desglose por tipo (solo si se activó)
    if load_real_activity:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("##### Desglose por tipo de actividad")
        if activity_detail_df.empty:
            st.info("No se encontró actividad real asociada a negocios en este período.")
        else:
            TYPE_LABELS = {
                "calls":    "📞 Llamadas",
                "emails":   "📧 Emails",
                "meetings": "📅 Reuniones",
                "notes":    "📝 Notas",
                "tasks":    "✅ Tareas",
            }
            if "activity_type" in activity_detail_df.columns:
                activity_detail_df["tipo"] = activity_detail_df["activity_type"].map(TYPE_LABELS).fillna(activity_detail_df["activity_type"])
            pivot_act = activity_detail_df.groupby(["owner","tipo"]).size().reset_index(name="n")
            fig2 = px.bar(
                pivot_act, x="owner", y="n", color="tipo",
                barmode="stack",
                labels={"owner":"Comercial","n":"Actividades","tipo":"Tipo"},
                color_discrete_map={
                    "📞 Llamadas":  "#185FA5",
                    "📧 Emails":    "#1D9E75",
                    "📅 Reuniones": "#F4A835",
                    "📝 Notas":     "#9B59B6",
                    "✅ Tareas":    "#E05252",
                },
                height=340,
            )
            fig2.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.02), **PLOT_LAYOUT)
            fig2.update_xaxes(tickangle=-20)
            st.plotly_chart(fig2, use_container_width=True)

            type_totals = activity_detail_df["tipo"].value_counts().reset_index()
            type_totals.columns = ["Tipo","Total"]
            c_cols = st.columns(len(type_totals))
            for col, (_, row) in zip(c_cols, type_totals.iterrows()):
                col.metric(row["Tipo"], int(row["Total"]))

# ── TAB 2: ETAPAS ──────────────────────────────

with tabs[2]:
    st.subheader("Actividad por etapa esta semana")
    top = stage_df[stage_df["active_this_week"] > 0].head(20)
    if top.empty:
        st.info("No hay actividad en ninguna etapa en el período seleccionado.")
    else:
        fig = px.bar(
            top, x="active_this_week", y="stage_label", orientation="h",
            color="active_this_week",
            color_continuous_scale=["#B5D4F4","#185FA5"],
            text="active_this_week",
            labels={"active_this_week":"Negocios con actividad","stage_label":""},
            height=520,
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(coloraxis_showscale=False, margin=dict(t=20,b=20,l=260,r=60), **PLOT_LAYOUT)
        st.plotly_chart(fig, use_container_width=True)

    with st.expander("Ver tabla completa de etapas"):
        st.dataframe(
            stage_df[["stage_label","total_deals","active_this_week"]].rename(columns={
                "stage_label":"Etapa","total_deals":"Total negocios","active_this_week":"Con actividad"
            }),
            hide_index=True, use_container_width=True,
        )

    st.markdown("##### Haz clic en una etapa para ver sus negocios")
    for _, srow in stage_df[stage_df["active_this_week"] > 0].head(20).iterrows():
        label = srow.get("stage_label", srow["dealstage"])
        with st.expander(f"📋 **{label}** — {srow['active_this_week']} con actividad / {srow['total_deals']} total"):
            tab_s1, tab_s2 = st.tabs(["✅ Con actividad en el período", "💤 Todos en esta etapa"])
            with tab_s1:
                deals_stage_active = df_week[df_week["dealstage"] == srow["dealstage"]]
                show_deal_table(deals_stage_active, eng_types=engagement_types)
            with tab_s2:
                deals_stage_all = df_all[df_all["dealstage"] == srow["dealstage"]]
                show_deal_table(deals_stage_all)

# ── TAB 3: CAMBIOS DE ETAPA ────────────────────

with tabs[3]:
    st.subheader("Cambios de etapa esta semana")
    if trans_df.empty:
        st.info("No se detectaron cambios de etapa. Activa 'Cargar historial de etapas' en la barra lateral.")
    else:
        st.caption(f"**{len(trans_df)}** cambios detectados en el período")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### 🟢 Etapas con más entradas")
            to_c = trans_df.groupby("to_stage").size().reset_index(name="n").sort_values("n", ascending=False).head(10)
            fig = px.bar(to_c, x="n", y="to_stage", orientation="h", text="n",
                         color_discrete_sequence=["#1D9E75"],
                         labels={"n":"","to_stage":""}, height=320)
            fig.update_traces(textposition="outside")
            fig.update_layout(margin=dict(t=10,b=10,l=230,r=50), **PLOT_LAYOUT)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            st.markdown("##### 🔴 Etapas con más salidas")
            from_c = trans_df.groupby("from_stage").size().reset_index(name="n").sort_values("n", ascending=False).head(10)
            fig = px.bar(from_c, x="n", y="from_stage", orientation="h", text="n",
                         color_discrete_sequence=["#E05252"],
                         labels={"n":"","from_stage":""}, height=320)
            fig.update_traces(textposition="outside")
            fig.update_layout(margin=dict(t=10,b=10,l=230,r=50), **PLOT_LAYOUT)
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("##### Detalle de todos los cambios")

        # Agrupar movimientos por negocio en una sola línea
        trans_sorted = trans_df.sort_values(["dealname","changed_at"]).reset_index(drop=True)

        # Construir recorrido completo por negocio
        deals_flow = {}
        for _, row in trans_sorted.iterrows():
            did = row["deal_id"]
            if did not in deals_flow:
                deals_flow[did] = {
                    "dealname": row["dealname"],
                    "owner":    row["owner"],
                    "pipeline": row["pipeline"],
                    "fecha":    str(row["changed_at"])[:16] if row["changed_at"] else "—",
                    "stages":   [row["from_stage"]],
                }
            deals_flow[did]["stages"].append(row["to_stage"])
            # Actualizar fecha con la más reciente
            if row["changed_at"]:
                deals_flow[did]["fecha"] = str(row["changed_at"])[:16]

        for did, info in deals_flow.items():
            stages    = info["stages"]
            n_moves   = len(stages) - 1
            badge     = f'<span style="background:#E8EAF0;border-radius:10px;padding:2px 8px;font-size:11px;color:#555;margin-left:8px">{n_moves} movimiento{"s" if n_moves>1 else ""}</span>'

            # Construir la cadena de etapas con flechas
            stages_html = ""
            for i, stage in enumerate(stages):
                if i == 0:
                    color_bg, color_txt = "#FCE4D6", "#7B3B00"
                elif i == len(stages) - 1:
                    color_bg, color_txt = "#E2EFDA", "#375623"
                else:
                    color_bg, color_txt = "#E6F1FB", "#0C447C"
                stages_html += f'<span style="background:{color_bg};color:{color_txt};padding:3px 10px;border-radius:12px;font-size:12px;white-space:nowrap">{stage}</span>'
                if i < len(stages) - 1:
                    stages_html += '<span style="font-size:16px;color:#185FA5;margin:0 4px">→</span>'

            st.markdown(f"""
            <div style="margin-top:10px; padding:10px 14px; background:#f8f9fb;
                        border-radius:8px; border-left:3px solid #185FA5;">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                    <span style="font-weight:700; font-size:0.9rem">
                        {info['dealname']}{badge}
                    </span>
                    <span style="font-size:11px; color:#888">{info['owner']} · {info['pipeline']} · {info['fecha']}</span>
                </div>
                <div style="display:flex; align-items:center; gap:4px; flex-wrap:wrap;">
                    {stages_html}
                </div>
            </div>
            """, unsafe_allow_html=True)

# ── TAB 4: NUEVOS NEGOCIOS ─────────────────────

with tabs[4]:
    st.subheader(f"Negocios creados del {start_date.strftime('%d/%m')} al {end_date.strftime('%d/%m')}")
    st.caption("Negocios con fecha de creación dentro del período de la barra lateral, y la etapa en la que están hoy. "
               "Solo negocios abiertos" + ("" if incluir_bbdd else " y fuera de BBDD") + ".")
    if new_df.empty:
        st.info("No se crearon negocios nuevos en el período seleccionado.")
    else:
        c1, c2 = st.columns([1, 2])
        with c1:
            st.metric("Total nuevos en el período", int(new_df["new_deals"].sum()))
            st.dataframe(
                new_df[["stage_label","new_deals"]].rename(columns={
                    "stage_label":"Etapa actual","new_deals":"Negocios"
                }),
                hide_index=True, use_container_width=True,
            )
        with c2:
            fig = px.pie(
                new_df, values="new_deals", names="stage_label", hole=0.45,
                color_discrete_sequence=px.colors.sequential.Blues_r,
            )
            fig.update_traces(textposition="inside", textinfo="percent+label")
            fig.update_layout(showlegend=False, height=380, **PLOT_LAYOUT)
            st.plotly_chart(fig, use_container_width=True)

        # ── Quiénes son y por qué están aquí ──
        lista = nuevos_df.sort_values("createdate", ascending=False).copy()
        lista["stage_label"] = lista["dealstage"].map(STAGE_LABELS).fillna(lista["dealstage"])
        with st.expander(f"🔎 Ver qué negocios son ({len(lista)})"):
            n1, n2 = st.columns(2)
            with n1:
                etapa_nf = st.selectbox("Etapa actual", ["Todas"] + sorted(lista["stage_label"].unique()), key="nuevos_etapa")
            with n2:
                com_nf = st.selectbox("Comercial", ["Todos"] + sorted(lista["owner"].unique()), key="nuevos_com")
            if etapa_nf != "Todas":
                lista = lista[lista["stage_label"] == etapa_nf]
            if com_nf != "Todos":
                lista = lista[lista["owner"] == com_nf]
            st.caption(f"**{len(lista)}** negocios")
            st.dataframe(pd.DataFrame({
                "Negocio":           lista["dealname"],
                "Comercial":         lista["owner"],
                "Pipeline":          lista["pipeline"],
                "Por qué está aquí": "Creado el " + lista["createdate"].dt.strftime("%d/%m/%Y"),
                "Etapa actual":      lista["stage_label"],
            }), hide_index=True, use_container_width=True, height=min(450, 55 + len(lista) * 35))

# ── TAB 5: ESTANCADOS ──────────────────────────

with tabs[5]:
    st.subheader("⚠️ Negocios sin actividad")
    st.caption("Negocios abiertos (no ganados ni perdidos) sin actividad real desde hace X días.")

    threshold = st.radio(
        "Mostrar negocios sin actividad en más de:",
        options=[14, 30, 60],
        format_func=lambda x: f"{x} días",
        horizontal=True,
    )
    df_stale = stale_data.get(threshold, pd.DataFrame())

    if df_stale.empty:
        st.success(f"✅ No hay negocios estancados más de {threshold} días.")
    else:
        st.warning(f"**{len(df_stale)} negocios** llevan más de {threshold} días sin actividad.")
        by_owner = df_stale.groupby("owner").size().reset_index(name="estancados").sort_values("estancados", ascending=False)
        fig = px.bar(
            by_owner, x="owner", y="estancados", text="estancados",
            color="estancados", color_continuous_scale=["#FFF2CC","#E05252"],
            labels={"owner":"Comercial","estancados":"Negocios estancados"},
            height=320,
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(coloraxis_showscale=False, **PLOT_LAYOUT)
        st.plotly_chart(fig, use_container_width=True)

        def color_dias(val):
            if val >= 60:  return "background-color: #fde8e8; color: #a00"
            elif val >= 30: return "background-color: #fff3cd; color: #7a5000"
            return ""

        show_df = df_stale[["dealname","owner","stage_label","pipeline","days_inactive"]].rename(columns={
            "dealname":"Negocio","owner":"Comercial","stage_label":"Etapa actual",
            "pipeline":"Pipeline","days_inactive":"Días sin actividad",
        })
        try:
            styled = show_df.style.map(color_dias, subset=["Días sin actividad"])
        except AttributeError:
            styled = show_df.style.applymap(color_dias, subset=["Días sin actividad"])
        st.dataframe(
            styled,
            hide_index=True, use_container_width=True, height=350,
        )

        st.markdown("##### Detalle por comercial")
        for owner_name, grp in df_stale.groupby("owner"):
            with st.expander(f"👤 **{owner_name}** — {len(grp)} negocios estancados"):
                show_deal_table(grp)

# ── TAB 6: MAPA DE CALOR ───────────────────────

with tabs[6]:
    st.subheader("🌡️ Mapa de calor — Comercial × Etapa")
    st.caption("Intensidad = número de negocios. Detecta dónde se acumulan y posibles cuellos de botella.")

    df_heat = df_all.copy()
    df_heat["stage_label"] = df_heat["dealstage"].map(STAGE_LABELS).fillna(df_heat["dealstage"])
    stage_counts = df_heat.groupby("stage_label")["deal_id"].count()
    top_stages   = stage_counts[stage_counts > 0].nlargest(15).index.tolist()
    df_heat      = df_heat[df_heat["stage_label"].isin(top_stages)]
    pivot = df_heat.pivot_table(index="owner", columns="stage_label",
                                values="deal_id", aggfunc="count", fill_value=0)
    pivot_total = pivot.copy()
    totals = pivot_total.sum(axis=0)
    totals.name = "TOTAL"
    pivot_with_total = pd.concat([pivot_total, totals.to_frame().T])

    fig = px.imshow(
        pivot_with_total,
        color_continuous_scale=["#EBF4FB","#185FA5"],
        aspect="auto", text_auto=True,
        labels={"color":"Negocios"},
        height=460,
    )
    fig.update_xaxes(tickangle=-35, tickfont_size=11)
    fig.update_yaxes(tickfont_size=11)
    fig.update_layout(
        coloraxis_showscale=True,
        margin=dict(t=30, b=120, l=150, r=20),
        **PLOT_LAYOUT,
    )
    fig.add_hline(
        y=len(pivot_with_total) - 1.5,
        line_dash="dash", line_color="#185FA5", line_width=1.5,
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("##### Total negocios por etapa")
    totals_df = pd.DataFrame({
        "Etapa": totals.index,
        "Total negocios": totals.values.astype(int)
    }).sort_values("Total negocios", ascending=False)
    st.dataframe(totals_df, hide_index=True, use_container_width=True, height=300)



# ── TAB 7: COMPARATIVA ────────────────────────

with tabs[7]:
    st.subheader("📊 Comparativa entre semanas")
    st.caption("Guarda snapshots de distintas semanas desde la barra lateral y compáralas aquí.")

    if len(st.session_state.snapshots) < 1:
        st.info("Todavía no hay snapshots guardados. Selecciona un período, carga los datos y pulsa 'Guardar snapshot actual' en la barra lateral.")
    else:
        snap_keys = list(st.session_state.snapshots.keys())

        c1, c2 = st.columns(2)
        with c1:
            snap_a = st.selectbox("Semana A", options=snap_keys, key="snap_a")
        with c2:
            snap_b = st.selectbox("Semana B", options=snap_keys, key="snap_b",
                                  index=min(1, len(snap_keys)-1))

        if snap_a and snap_b and snap_a != snap_b:
            sa = st.session_state.snapshots[snap_a]
            sb = st.session_state.snapshots[snap_b]

            label_a = sa["week_label"]
            label_b = sb["week_label"]
            if snap_a > snap_b:
                sa, sb = sb, sa
                snap_a, snap_b = snap_b, snap_a
                label_a, label_b = label_b, label_a

            dfa = sa["df_week"]
            dfb = sb["df_week"]

            ids_a = set(dfa["deal_id"])
            ids_b = set(dfb["deal_id"])

            st.divider()
            kc1, kc2, kc3, kc4 = st.columns(4)
            kc1.metric(f"Activos {snap_a}", len(dfa))
            kc2.metric(f"Activos {snap_b}", len(dfb), delta=f"{len(dfb)-len(dfa):+d}")
            kc3.metric("↑ Suben actividad", len(ids_b - ids_a), help=f"Activos en {snap_b} pero no en {snap_a}")
            kc4.metric("↓ Bajan actividad", len(ids_a - ids_b), help=f"Activos en {snap_a} pero no en {snap_b}")

            st.divider()
            st.markdown("##### Actividad por comercial")
            act_a = dfa.groupby("owner").size().reset_index(name="n_a")
            act_b = dfb.groupby("owner").size().reset_index(name="n_b")
            comp_df = act_a.merge(act_b, on="owner", how="outer").fillna(0)
            comp_df["n_a"] = comp_df["n_a"].astype(int)
            comp_df["n_b"] = comp_df["n_b"].astype(int)
            comp_df = comp_df.sort_values("n_b", ascending=False)

            fig = go.Figure()
            fig.add_bar(name=snap_a, x=comp_df["owner"], y=comp_df["n_a"],
                        marker_color="#B5D4F4", marker_line_width=0)
            fig.add_bar(name=snap_b, x=comp_df["owner"], y=comp_df["n_b"],
                        marker_color="#185FA5", marker_line_width=0)
            fig.update_layout(barmode="group", height=360,
                              legend=dict(orientation="h", yanchor="bottom", y=1.02),
                              **PLOT_LAYOUT)
            fig.update_xaxes(tickangle=-20)
            st.plotly_chart(fig, use_container_width=True)

            st.divider()
            st.markdown("##### Detalle de negocios — actividad por período")

            all_ids = ids_a | ids_b
            df_all_snap = pd.concat([dfa, dfb]).drop_duplicates(subset="deal_id")
            df_all_snap = df_all_snap[df_all_snap["deal_id"].isin(all_ids)].copy()
            df_all_snap["dealstage"] = df_all_snap["dealstage"].astype(str).str.replace(r"\.0$", "", regex=True)

            # Siempre con los nombres de etapa actuales de HubSpot (el snapshot puede ser antiguo)
            if True:
                df_all_snap["stage_label"] = df_all_snap["dealstage"].map(STAGE_LABELS).fillna(df_all_snap["dealstage"])
            else:
                mask = (
                    df_all_snap["stage_label"].isna() |
                    df_all_snap["stage_label"].astype(str).str.match(r'^\d+(\.0)?$', na=False)
                )
                df_all_snap.loc[mask, "stage_label"] = df_all_snap.loc[mask, "dealstage"].map(STAGE_LABELS).fillna(df_all_snap.loc[mask, "dealstage"])

            df_all_snap[snap_a] = df_all_snap["deal_id"].apply(lambda x: "✅" if x in ids_a else "❌")
            df_all_snap[snap_b] = df_all_snap["deal_id"].apply(lambda x: "✅" if x in ids_b else "❌")
            df_all_snap["Tendencia"] = df_all_snap["deal_id"].apply(
                lambda x: "✅ Nueva actividad"   if x in ids_b and x not in ids_a
                     else "📉 Perdió actividad"  if x in ids_a and x not in ids_b
                     else "🔄 Activo en ambos"
            )

            show_comp = df_all_snap[["dealname","owner","stage_label","pipeline", snap_a, snap_b, "Tendencia"]].rename(columns={
                "dealname":    "Negocio",
                "owner":       "Comercial",
                "stage_label": "Etapa",
                "pipeline":    "Pipeline",
            }).sort_values("Tendencia")

            tend_filter = st.radio("Filtrar por tendencia:", ["Todos", "✅ Nueva actividad", "📉 Perdió actividad", "🔄 Activo en ambos"], horizontal=True)
            if tend_filter != "Todos":
                show_comp = show_comp[show_comp["Tendencia"] == tend_filter]

            st.caption(f"**{len(show_comp)}** negocios")
            st.dataframe(
                show_comp,
                hide_index=True,
                use_container_width=True,
                height=min(500, 55 + len(show_comp) * 35),
                column_config={
                    "Tendencia": st.column_config.TextColumn("Tendencia", width="small"),
                    snap_a:      st.column_config.TextColumn(snap_a, width="small"),
                    snap_b:      st.column_config.TextColumn(snap_b, width="small"),
                }
            )

        elif snap_a == snap_b:
            st.warning("Selecciona dos períodos diferentes para comparar.")


# ── EXPORTAR EXCEL ─────────────────────────────

st.divider()
st.subheader("📥 Exportar a Excel")
st.caption("Excel con 7 pestañas: resumen, actividad, etapas, cambios, nuevos, estancados y matriz completa.")

filename    = f"pipeline_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"

comp_excel = None
snap_keys = list(st.session_state.snapshots.keys())
if len(snap_keys) >= 2:
    snap_a_key = st.session_state.get("snap_a", snap_keys[0])
    snap_b_key = st.session_state.get("snap_b", snap_keys[1] if len(snap_keys) > 1 else snap_keys[0])
    if snap_a_key != snap_b_key and snap_a_key in st.session_state.snapshots and snap_b_key in st.session_state.snapshots:
        sa = st.session_state.snapshots[snap_a_key]
        sb = st.session_state.snapshots[snap_b_key]
        comp_excel = {
            "snap_a_name": snap_a_key,
            "snap_b_name": snap_b_key,
            "dfa":         sa["df_week"],
            "dfb":         sb["df_week"],
            "act_a":       sa["act_df"],
            "act_b":       sb["act_df"],
        }

excel_bytes = generate_excel(week_label, act_df, stage_df, trans_df, new_df, matrix_df, stale_data,
                              comparativa=comp_excel)

c1, c2 = st.columns(2)
with c1:
    st.download_button(
        label="⬇️ Descargar en el navegador",
        data=excel_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with c2:
    if st.button("💾 Guardar en carpeta local", use_container_width=True):
        try:
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            full_path = os.path.join(OUTPUT_DIR, filename)
            with open(full_path, "wb") as f:
                f.write(excel_bytes)
            st.success(f"✅ Guardado en:\n`{full_path}`")
        except Exception as e:
            st.error(f"❌ Error al guardar: {e}")