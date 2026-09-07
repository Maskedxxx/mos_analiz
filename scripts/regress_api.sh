#!/bin/bash
# Регресс через HTTP API «до/после» правки: логин → аудит файла → результат → отчёт.
# Использование: bash scripts/regress_api.sh <файл> <doc_type>
# Логин/пароль и порт — из .env (AUDIT_LOGIN, AUDIT_PASSWORD, API_PORT). Печатает число нарушений,
# число проверенных правил и размер отчёта — сравнивать с базовым срезом до правки.
set -u
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
if [ -f "$ROOT/.env" ]; then set -a; . "$ROOT/.env"; set +a; fi
F="${1:?файл}"; T="${2:?doc_type}"; P="${API_PORT:-8081}"; J=$(mktemp)
code=$(curl -s -c "$J" -o /dev/null -w '%{http_code}' -X POST "http://127.0.0.1:$P/api/login" \
  -H 'Content-Type: application/json' -d "{\"login\":\"${AUDIT_LOGIN:-}\",\"password\":\"${AUDIT_PASSWORD:-}\"}")
echo "login=$code"; [ "$code" = "200" ] || exit 1
sid=$(curl -s -b "$J" -X POST "http://127.0.0.1:$P/api/audit" -F "file=@$F" -F "doc_type=$T" \
  | python3 -c 'import sys,json;print(json.load(sys.stdin)["session_id"])') || exit 1
echo "session=$sid"
rc=""
for i in $(seq 1 180); do
  rc=$(curl -s -b "$J" -o /tmp/regress_result.json -w '%{http_code}' "http://127.0.0.1:$P/api/audit/$sid/result" || true)
  if [ "$rc" != "202" ]; then break; fi
  sleep 5
done
echo "result_http=$rc"
python3 -c 'import json;d=json.load(open("/tmp/regress_result.json"));print("violations=",len(d.get("violations",[])),"rules_checked=",d.get("rules_checked"),"duration_sec=",round(d.get("duration_sec",0),1))' 2>/dev/null || head -c 300 /tmp/regress_result.json
dl=$(curl -s -b "$J" -o /tmp/regress_report.xlsx -w '%{http_code}' "http://127.0.0.1:$P/api/audit/$sid/download")
echo "download=$dl size=$(stat -c %s /tmp/regress_report.xlsx 2>/dev/null || stat -f %z /tmp/regress_report.xlsx)"
