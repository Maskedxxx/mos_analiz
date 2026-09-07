#!/bin/bash
# Полный перезапуск сервиса проверки: бэкенд + фронтенд. Затем — проверка здоровья.
set -e
D="$(dirname "$0")"
bash "$D/restart_backend.sh"
bash "$D/restart_frontend.sh"
bash "$D/healthcheck.sh"
