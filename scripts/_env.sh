#!/bin/bash
# Общая часть скриптов запуска: корень проекта, переменные из .env, порты по умолчанию.
# Подключается через `source "$(dirname "$0")/_env.sh"`. Сам по себе ничего не запускает.
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
if [ -f "$ROOT/.env" ]; then
  set -a; . "$ROOT/.env"; set +a
else
  echo "[env] ВНИМАНИЕ: нет $ROOT/.env — скопируйте .env.example в .env и заполните" >&2
fi
API_HOST="${API_HOST:-127.0.0.1}"     # адрес, на котором слушает бэкенд (0.0.0.0 — для доступа снаружи)
API_PORT="${API_PORT:-8081}"          # порт бэкенда проверки
UI_PORT="${UI_PORT:-5174}"            # порт фронтенда (vite preview)
PYTHON="${PYTHON:-$ROOT/.venv/bin/python}"   # интерпретатор виртуального окружения
LOG_DIR="${LOG_DIR:-/tmp}"            # куда писать логи процессов

# --- F29: порт должен слушать именно наш процесс (находка 6.5: чужой процесс отвечал 200, а свой умирал) ---
# PID процесса, слушающего TCP-порт $1 (пусто — никто не слушает).
port_pid() { ss -ltnp "sport = :$1" 2>/dev/null | grep -o 'pid=[0-9]*' | head -1 | cut -d= -f2; }
# PID SCREEN-демона сессии с именем $1.
screen_pid() { screen -ls 2>/dev/null | grep -oE "[0-9]+\.$1[[:space:]]" | head -1 | cut -d. -f1; }
# Код 0, если процесс $1 — потомок процесса $2 (подъём по ppid до init).
pid_is_under() { local p="$1"; while [ -n "$p" ] && [ "$p" != "1" ] && [ "$p" != "0" ]; do [ "$p" = "$2" ] && return 0; p=$(ps -o ppid= -p "$p" 2>/dev/null | tr -d ' '); done; return 1; }
# Порт $1 обязан слушать процесс из screen-сессии $2; иначе сообщение с PID чужого процесса (метка $3) и код 1.
check_port_owner() {
  local lp sp; lp=$(port_pid "$1"); sp=$(screen_pid "$2")
  if [ -z "$lp" ]; then echo "[$3] ОШИБКА: порт $1 никто не слушает" >&2; return 1; fi
  if [ -z "$sp" ] || ! pid_is_under "$lp" "$sp"; then
    echo "[$3] ОШИБКА: порт $1 занят процессом $lp ($(ps -o cmd= -p "$lp" 2>/dev/null | cut -c1-80)) — наш процесс не поднялся, см. лог" >&2
    return 1
  fi
  echo "[$3] порт $1 слушает наш процесс pid=$lp (screen $2)"
}
