#!/bin/bash
# Проверка здоровья: бэкенд (/api/health с состоянием LLM, OCR, layout) и фронтенд (код ответа).
# Выход 0 — всё отвечает; 1 — что-то лежит или бэкенд degraded (HTTP 503: внешний сервис недоступен).
source "$(dirname "$0")/_env.sh"
ok=0
hc=$(curl -s -o /dev/null -w '%{http_code}' --max-time 10 "http://127.0.0.1:$API_PORT/api/health")
h=$(curl -s --max-time 10 "http://127.0.0.1:$API_PORT/api/health")
if [ -n "$h" ]; then
  [ "$hc" = "200" ] || ok=1
  echo "[health] backend :$API_PORT -> HTTP $hc $(echo "$h" | "$PYTHON" -c 'import sys,json
d=json.load(sys.stdin); print(d.get("status"), {k:v.get("status") for k,v in d.get("services",{}).items()})' 2>/dev/null || echo "$h" | head -c 200)"
else
  echo "[health] backend :$API_PORT -> НЕ ОТВЕЧАЕТ"; ok=1
fi
code=$(curl -s -o /dev/null -w '%{http_code}' --max-time 5 "http://127.0.0.1:$UI_PORT/")
echo "[health] frontend :$UI_PORT -> $code"
[ "$code" = "200" ] || ok=1
exit $ok
