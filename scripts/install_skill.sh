#!/bin/bash
# Установка скила `mosmonitoring` для Claude Code на этой машине: ~/.claude/skills/mosmonitoring
# → симлинк на .claude/skills/mosmonitoring в этом репозитории (обновляется вместе с git pull).
# Использование: bash scripts/install_skill.sh   (повторный запуск безопасен)
set -e
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SRC="$ROOT/.claude/skills/mosmonitoring"
DST="$HOME/.claude/skills/mosmonitoring"
[ -f "$SRC/SKILL.md" ] || { echo "нет $SRC/SKILL.md" >&2; exit 1; }
mkdir -p "$HOME/.claude/skills"
if [ -e "$DST" ] && [ ! -L "$DST" ]; then mv "$DST" "$DST.bak.$(date +%Y%m%d%H%M%S)"; echo "старая копия сохранена как $DST.bak.*"; fi
ln -sfn "$SRC" "$DST"
echo "скил установлен: $DST -> $SRC ($(grep -c . "$SRC/SKILL.md") строк SKILL.md, references: $(ls "$SRC/references" | wc -l | tr -d ' '))"
