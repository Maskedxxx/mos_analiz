#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Временный скрипт: прогон всех шаблонов через Paddle OCR и сохранение кэша.

Для каждого doc_type с parser=paddle и template.docx:
1. Загружает AuditConfig
2. Парсит template.docx через PaddleParser
3. Сохраняет template_cached.json

Запуск:
    python cache_all_templates.py
"""

import hashlib
import json
import sys
import tempfile
from datetime import datetime
from pathlib import Path

# Корень проекта
ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))

from audit_engine.models import load_audit_config
from audit_engine.paddle_parser import PaddleParser


def compute_hash(file_path: Path) -> str:
    """SHA256 хэш файла."""
    sha256 = hashlib.sha256()
    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            sha256.update(chunk)
    return sha256.hexdigest()


def main():
    configs_dir = ROOT / "doc_configs"
    doc_types = sorted([d.name for d in configs_dir.iterdir() if d.is_dir()])

    print(f"Найдено {len(doc_types)} типов документов\n")

    parsed = 0
    skipped = 0
    errors = 0

    for doc_type in doc_types:
        config_dir = configs_dir / doc_type
        config_path = config_dir / "config.json"

        if not config_path.exists():
            continue

        # Загружаем конфиг
        try:
            config = load_audit_config(config_dir)
        except Exception as e:
            print(f"  [{doc_type}] ❌ Ошибка загрузки конфига: {e}")
            errors += 1
            continue

        # Только paddle-парсер
        if config.parser != "paddle":
            print(f"  [{doc_type}] ⏭️  Парсер={config.parser}, пропускаем")
            skipped += 1
            continue

        # Проверяем наличие шаблона
        template_path = config_dir / "template" / "template.docx"
        if not template_path.exists():
            print(f"  [{doc_type}] ⏭️  Нет template.docx, пропускаем")
            skipped += 1
            continue

        cache_path = config_dir / "template" / "template_cached.json"

        # Проверяем: может кэш уже актуальный?
        file_hash = compute_hash(template_path)
        if cache_path.exists():
            try:
                with open(cache_path, 'r', encoding='utf-8') as f:
                    cached = json.load(f)
                if cached.get("_hash") == file_hash and cached.get("_parser") == "paddle":
                    print(f"  [{doc_type}] ✅ Кэш актуален, пропускаем")
                    skipped += 1
                    continue
            except (json.JSONDecodeError, KeyError):
                pass

        # Парсим через PaddleParser
        print(f"  [{doc_type}] 🔄 Парсинг шаблона...")

        try:
            with tempfile.TemporaryDirectory() as tmp_dir:
                parser = PaddleParser(config=config, log_dir=tmp_dir)
                doc = parser.parse(str(template_path))

            # Сохраняем кэш
            cache_data = dict(doc)
            cache_data["_hash"] = file_hash
            cache_data["_parser"] = "paddle"
            cache_data["_cached_at"] = datetime.now().isoformat()

            with open(cache_path, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)

            # Считаем чанки (без служебных полей)
            chunks = [k for k in doc.keys() if not k.startswith("_")]
            print(f"  [{doc_type}] ✅ Готово! Чанков: {len(chunks)}, кэш: {cache_path.name}")
            parsed += 1

        except Exception as e:
            print(f"  [{doc_type}] ❌ Ошибка парсинга: {e}")
            errors += 1

    print(f"\n{'='*60}")
    print(f"Итого: {parsed} распарсено, {skipped} пропущено, {errors} ошибок")


if __name__ == "__main__":
    main()
