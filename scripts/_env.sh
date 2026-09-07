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
