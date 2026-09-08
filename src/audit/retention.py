# START_MODULE_CONTRACT
# PURPOSE: Ротация каталогов по объёму (F28): пока корень (logs_result/, uploads/) больше лимита в ГБ — удаляются самые старые каталоги целиком.
# INPUTS: корень, лимит в ГБ (0 — не удалять), glob-шаблон каталогов-кандидатов, имена каталогов, которые трогать нельзя (живые сессии).
# OUTPUTS: список удалённых каталогов; фоновый поток, повторяющий очистку раз в сутки.
# KEYWORDS: retention, rotation, disk, LOGS_MAX_GB, UPLOADS_MAX_GB.
# LINKS: src/api/server.py (_retention_job — вызов при старте и в фоне), .env.example (лимиты).
# RATIONALE: Хранение по времени владелец отклонил — лимит по объёму (решение 08.09). Удаляются только
#   целые каталоги session_* / uploads/<id> (вместе с original/ внутри), никогда — файлы внутри живых сессий.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import shutil
import sys
import threading
import time
from pathlib import Path
from typing import Callable, Iterable, List
# END_IMPORTS

GB = 1024 ** 3


def dir_size_bytes(path: Path) -> int:
    """Суммарный размер файлов в каталоге (рекурсивно, без симлинков). Недоступные файлы пропускаются."""
    total = 0
    for p in path.rglob("*"):
        try:
            if p.is_file() and not p.is_symlink():
                total += p.stat().st_size
        except OSError:
            continue
    return total


def prune_oldest(root: Path, max_gb: float, pattern: str, skip: Iterable[str] = ()) -> List[Path]:
    """
    Удаляет самые старые каталоги `root.glob(pattern)` (по mtime), пока размер `root` больше `max_gb`.

    Вход:
        root: корень (logs_result/ или uploads/).
        max_gb: лимит в гигабайтах; 0 или меньше — ничего не удалять.
        pattern: glob каталогов-кандидатов ("*/session_*" для logs_result, "*" для uploads).
        skip: имена каталогов, которые удалять нельзя (сессии, которые сейчас выполняются).
    Выход: список удалённых каталогов (пустой, если лимит не превышен).
    """
    removed: List[Path] = []
    if max_gb <= 0 or not root.exists():
        return removed
    limit = int(max_gb * GB)
    total = dir_size_bytes(root)
    if total <= limit:
        return removed
    skip_names = set(skip)
    candidates = [p for p in root.glob(pattern) if p.is_dir() and p.name not in skip_names]
    candidates.sort(key=lambda p: p.stat().st_mtime)
    # Удаляем по одному самому старому каталогу, пока не уложимся в лимит.
    for p in candidates:
        if total <= limit:
            break
        size = dir_size_bytes(p)
        shutil.rmtree(p, ignore_errors=True)
        total -= size
        removed.append(p)
    return removed


def start_retention_thread(job: Callable[[], None], interval_sec: int = 86400) -> threading.Thread:
    """Запускает daemon-поток: раз в `interval_sec` (сутки) вызывает `job`; ошибки очистки — в stderr, не наружу."""
    def _loop() -> None:
        while True:
            time.sleep(interval_sec)
            try:
                job()
            except Exception as e:
                print(f"[RETENTION] ошибка очистки: {e}", file=sys.stderr, flush=True)
    t = threading.Thread(target=_loop, daemon=True, name="retention")
    t.start()
    return t
