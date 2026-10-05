#!/bin/bash
set -e
cd "$(dirname "$0")"
PORT="${FINANCE_HUB_PORT:-8501}"

if [ ! -x ".venv/bin/python" ]; then
  echo "Ambiente non installato: eseguo l'installazione iniziale."
  exec ./INSTALLA_E_AVVIA.command
fi

if [ -f ".finance_hub.pid" ] && kill -0 "$(cat .finance_hub.pid)" 2>/dev/null; then
  echo "Finance Hub è già attivo."
  command -v open >/dev/null 2>&1 && open "http://localhost:${PORT}"
  exit 0
fi

nohup .venv/bin/python -m streamlit run app.py \
  --server.port "$PORT" \
  --server.address "localhost" \
  --server.headless true \
  > finance_hub.log 2>&1 &
echo $! > .finance_hub.pid

for i in {1..30}; do
  if curl -fsS "http://localhost:${PORT}/_stcore/health" >/dev/null 2>&1; then
    echo "Finance Hub 4.3 aggiornato attivo su http://localhost:${PORT}"
    command -v open >/dev/null 2>&1 && open "http://localhost:${PORT}"
    exit 0
  fi
  sleep 1
done

echo "Avvio non completato. Controlla finance_hub.log."
exit 1
