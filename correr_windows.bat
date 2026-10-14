@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === Reportes de Movilidad Internacional ===
if not exist .venv (
  echo Primera vez: creando entorno de Python e instalando librerias...
  py -3 -m venv .venv || python -m venv .venv
  call .venv\Scripts\activate.bat
  python -m pip install --upgrade pip >nul
  python -m pip install -r requirements.txt
) else (
  call .venv\Scripts\activate.bat
)
python run_pipeline.py %*
echo.
echo Listo. Los resultados estan en la carpeta "salidas".
pause
