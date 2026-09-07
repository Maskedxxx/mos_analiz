#!/bin/bash
# Регресс сквозной сверки комплекта через API: файлы → /api/audit/cross → результат → отчёт.
# Использование: bash scripts/regress_cross.sh <файл1> <файл2> <файл3>
# Роли документов определяются по ключевым словам в именах файлов (см. required_docs в config.json типа).
set -u
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
if [ -f "$ROOT/.env" ]; then set -a; . "$ROOT/.env"; set +a; fi
P="${API_PORT:-8081}"; J=$(mktemp)
[ $# -ge 2 ] || { echo "нужно минимум два файла" >&2; exit 1; }
curl -s -c "$J" -o /dev/null -X POST "http://127.0.0.1:$P/api/login" -H 'Content-Type: application/json' \
  -d "{\"login\":\"${AUDIT_LOGIN:-}\",\"password\":\"${AUDIT_PASSWORD:-}\"}"
ARGS=(); for f in "$@"; do ARGS+=(-F "files=@$f"); done
RAW=$(curl -s -b "$J" -X POST "http://127.0.0.1:$P/api/audit/cross" "${ARGS[@]}")
sid=$(echo "$RAW" | python3 -c 'import sys,json;print(json.load(sys.stdin)["session_id"])') || { echo "post=$RAW"; exit 1; }
echo "session=$sid"
rc=""
for i in $(seq 1 120); do
  rc=$(curl -s -b "$J" -o /tmp/regress_cross.json -w '%{http_code}' "http://127.0.0.1:$P/api/audit/$sid/result" || true)
  if [ "$rc" != "202" ]; then break; fi
  sleep 5
done
echo "result_http=$rc"
python3 -c 'import json;d=json.load(open("/tmp/regress_cross.json"));print("violations=",len(d.get("violations",[])),"rules_checked=",d.get("rules_checked"))' 2>/dev/null || head -c 300 /tmp/regress_cross.json
curl -s -b "$J" -o /tmp/regress_cross.xlsx -w 'download=%{http_code}\n' "http://127.0.0.1:$P/api/audit/$sid/download"
