#!/bin/bash
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python 3 non trovato. Installa Python 3.10 o successivo da https://www.python.org/downloads/macos/"
  read -r -p "Premi Invio per chiudere..."
  exit 1
fi

if [ ! -d ".venv" ]; then
  echo "Creo l'ambiente locale di Finance Hub..."
  python3 -m venv .venv
fi

source .venv/bin/activate
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt

echo "Installazione completata. Avvio Finance Hub 4.3 aggiornato..."
exec ./AVVIA_FINANCE_HUB.command
