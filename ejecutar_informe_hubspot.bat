@echo off
REM Cambiar al directorio del entorno base de Anaconda
CALL "C:\Users\Administracion1\anaconda3\Scripts\activate.bat" prueba

REM Cambiar a la carpeta donde está el archivo .py
cd "C:\Users\Administracion1\Zentralcom\Zentralcom S.L - Documentos\Administracion\NAYADE\CODIGOS PYTHON\informe automatico hubspot"

REM Ejecutar el archivo .py con el Python del entorno
python "Informe hubspot negocios.py"
