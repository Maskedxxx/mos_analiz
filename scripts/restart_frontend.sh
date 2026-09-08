#!/bin/bash
# Перезапуск фронтенда: vite preview отдаёт собранный dist и проксирует /api → бэкенд, /gen → генерация.
# Перед запуском пересобирает dist (build_frontend.sh). Screen-сессия `frontend`.
source "$(dirname "$0")/_env.sh"
bash "$ROOT/scripts/build_frontend.sh" || exit 1
screen -S frontend -X quit 2>/dev/null || true
sleep 1
NVM_LINE=""
if [ -n "$NODE_VERSION" ] && [ -s "$HOME/.nvm/nvm.sh" ]; then
  NVM_LINE=". '$HOME/.nvm/nvm.sh' && nvm use '$NODE_VERSION' >/dev/null &&"
fi
screen -dmS frontend bash -c "cd '$ROOT/ui-demo' && $NVM_LINE npx vite preview --port '$UI_PORT' > '$LOG_DIR/frontend.log' 2>&1"
# Ждём готовности до 60 секунд (curl при отказе соединения возвращает пустой код — это нормально).
code=""
for i in $(seq 1 60); do
  code=$(curl -s -o /dev/null -w '%{http_code}' "http://127.0.0.1:$UI_PORT/" 2>/dev/null || true)
  if [ "$code" = "200" ]; then break; fi
  sleep 1
done
echo "[frontend] http://127.0.0.1:$UI_PORT/ -> ${code:-нет ответа}  (лог: $LOG_DIR/frontend.log)"
if [ "$code" != "200" ]; then tail -5 "$LOG_DIR/frontend.log"; exit 1; fi
# F29: 200 мог ответить чужой процесс на этом порту — сверяем PID слушателя с нашей screen-сессией.
check_port_owner "$UI_PORT" "frontend" "frontend" || { tail -5 "$LOG_DIR/frontend.log"; exit 1; }
