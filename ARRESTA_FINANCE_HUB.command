#!/bin/bash
cd "$(dirname "$0")"
if [ -f ".finance_hub.pid" ]; then
  PID="$(cat .finance_hub.pid)"
  if kill -0 "$PID" 2>/dev/null; then
    kill "$PID"
    echo "Finance Hub arrestato."
  else
    echo "Finance Hub non era attivo."
  fi
  rm -f .finance_hub.pid
else
  echo "Nessun processo Finance Hub registrato."
fi
