#!/bin/bash
# Перезапуск бэкенда проверки (uvicorn main:app) в screen-сессии `backend`.
# После любой правки кода или правил — обязателен: Python не подхватывает изменения на лету.
# Переменные: API_HOST, API_PORT, PYTHON, LOG_DIR (см. _env.sh и .env).
source "$(dirname "$0")/_env.sh"
[ -x "$PYTHON" ] || { echo "[backend] нет интерпретатора $PYTHON — создайте venv: python3 -m venv .venv && .venv/bin/pip install -r requirements.txt" >&2; exit 1; }
screen -S backend -X quit 2>/dev/null || true
sleep 1
screen -dmS backend bash -c "cd '$ROOT'; if [ -f ./.env ]; then set -a; . ./.env; set +a; fi; exec '$PYTHON' -m uvicorn main:app --host '$API_HOST' --port '$API_PORT' >> '$LOG_DIR/backend.log' 2>&1"
# Ждём готовности до 30 секунд: сервер валидирует окружение и прогревает модули.
# 200 — всё доступно; 503 — сервер поднят, но внешний сервис (модель/OCR/layout) недоступен (degraded).
for i in $(seq 1 30); do
  code=$(curl -s -o /dev/null -w '%{http_code}' "http://127.0.0.1:$API_PORT/api/health" 2>/dev/null || true)
  if [ "$code" = "200" ] || [ "$code" = "503" ]; then break; fi
  sleep 1
done
echo "[backend] http://$API_HOST:$API_PORT/api/health -> ${code:-нет ответа}  (лог: $LOG_DIR/backend.log)"
if [ "$code" = "503" ]; then echo "[backend] состояние degraded — проверьте адреса LLM_BASE_URL/OCR_BASE_URL/LAYOUT_BASE_URL в .env: curl -s http://127.0.0.1:$API_PORT/api/health"; fi
if [ "$code" != "200" ] && [ "$code" != "503" ]; then tail -5 "$LOG_DIR/backend.log"; exit 1; fi
# F29: ответ мог дать чужой процесс на этом порту (старый прод, забытый uvicorn) — сверяем PID слушателя.
check_port_owner "$API_PORT" "backend" "backend" || { tail -5 "$LOG_DIR/backend.log"; exit 1; }
