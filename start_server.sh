#!/bin/bash
# Запуск FastAPI-сервера для веб-аудита документов
# Использование: bash start_server.sh

source /home/masked/projects/mos_analiz/.venv/bin/activate

export AUDIT_LOGIN="${AUDIT_LOGIN:-admin}"
export AUDIT_PASSWORD="${AUDIT_PASSWORD:-admin}"

echo "Запуск API-сервера на http://localhost:8080"
echo "Логин: $AUDIT_LOGIN"
echo "Ctrl+C для остановки"
echo "───────────────────────────────────────"

uvicorn api_server:app --host 0.0.0.0 --port 8080 --reload \
  --reload-dir audit_engine \
  --reload-dir doc_configs \
  --reload-include "api_server.py"
