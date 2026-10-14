#!/usr/bin/env bash
# Uso: ./correr_mac_linux.sh   [--ciclo 2025-2026] [--anios 5]
set -e
cd "$(dirname "$0")"
echo "=== Reportes de Movilidad Internacional ==="
if [ ! -d .venv ]; then
  echo "Primera vez: creando entorno de Python e instalando librerías..."
  python3 -m venv .venv
  . .venv/bin/activate
  python -m pip install --upgrade pip >/dev/null
  python -m pip install -r requirements.txt
else
  . .venv/bin/activate
fi
python run_pipeline.py "$@"
echo
echo "Listo. Los resultados están en la carpeta 'salidas'."
