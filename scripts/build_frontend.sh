#!/bin/bash
# Сборка фронтенда ui-demo (Vite + React). Нужен Node.js 20+ (проверено на 22).
# Если установлен nvm и задана NODE_VERSION — переключается на неё.
set -e
source "$(dirname "$0")/_env.sh"
cd "$ROOT/ui-demo"
if [ -n "$NODE_VERSION" ] && [ -s "$HOME/.nvm/nvm.sh" ]; then
  . "$HOME/.nvm/nvm.sh" && nvm use "$NODE_VERSION" >/dev/null
fi
major=$(node -p 'process.versions.node.split(".")[0]' 2>/dev/null || echo 0)
[ "$major" -ge 20 ] || { echo "[frontend] нужен Node.js 20+, найден: $(node --version 2>/dev/null || echo 'нет node')" >&2; exit 1; }
[ -d node_modules ] || { echo "[frontend] npm ci..."; npm ci; }
echo "[frontend] npm run build (node $(node --version))..."
npm run build
echo "[frontend] build OK → ui-demo/dist"
