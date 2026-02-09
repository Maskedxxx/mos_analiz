#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Единый CLI для аудита документов.

Использование:
    # Аудит документа
    python run_audit.py --doc-type presentation_eu --target document.pptx

    # Только парсинг (без проверок)
    python run_audit.py --doc-type presentation_eu --target document.pptx --parse-only

    # Одно правило
    python run_audit.py --doc-type presentation_eu --target document.pptx --rule-filter 3

    # Без кэша шаблона
    python run_audit.py --doc-type presentation_eu --target document.pptx --no-cache

    # Список типов документов
    python run_audit.py --list-types
"""

import argparse
import importlib
import json
import sys
from pathlib import Path

from audit_engine import AuditEngine

# Маппинг специальных движков (не Vision pipeline)
SPECIAL_ENGINES = {
    "drivers": "audit_engine.drivers",
    "kpsc": "audit_engine.kpsc",
}


def _detect_engine(doc_type: str) -> str:
    """
    Определяет тип движка по config.json.

    Если в config.json есть поле "engine" — возвращает его значение.
    Иначе возвращает "vision" (стандартный pipeline).
    """
    config_path = Path(__file__).parent / "doc_configs" / doc_type / "config.json"
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("engine", "vision")
    return "vision"


def main():
    """Точка входа CLI."""
    parser = argparse.ArgumentParser(
        description="Единый аудит документов — Vision Pipeline + LLM"
    )
    parser.add_argument(
        "--doc-type",
        default=None,
        help="Тип документа (имя папки в doc_configs/)"
    )
    parser.add_argument(
        "--target",
        default=None,
        help="Путь к целевому документу"
    )
    parser.add_argument(
        "--template",
        default=None,
        help="Путь к шаблону (опционально, по умолчанию — из конфига)"
    )
    parser.add_argument(
        "--model",
        default=None,
        help="Модель OpenAI (переопределяет конфиг)"
    )
    parser.add_argument(
        "--temperature",
        type=float,
        default=None,
        help="Температура генерации"
    )
    parser.add_argument(
        "--rule-filter",
        default=None,
        help="Проверить только указанное правило (номер: 3 для Vision, 1.1 для КПСЦ)"
    )
    parser.add_argument(
        "--parse-only",
        action="store_true",
        help="Только Vision-парсинг, без проверки правил"
    )
    parser.add_argument(
        "--no-cache",
        action="store_true",
        help="Не использовать кэш шаблона"
    )
    parser.add_argument(
        "--print-prompts",
        action="store_true",
        help="Режим отладки: только промпты без вызова LLM"
    )
    parser.add_argument(
        "--session-dir",
        default=None,
        help="Директория для логов сессии"
    )
    parser.add_argument(
        "--chunk-filter",
        default=None,
        help="Парсить только указанный чанк"
    )
    parser.add_argument(
        "--out-xlsx",
        default=None,
        help="Путь для сохранения Excel"
    )
    parser.add_argument(
        "--secondary",
        default=None,
        help="Путь к вторичному файлу (XLSX для multi-file аудитов)"
    )
    parser.add_argument(
        "--list-types",
        action="store_true",
        help="Показать список доступных типов документов"
    )

    args = parser.parse_args()

    # Режим --list-types
    if args.list_types:
        doc_types = AuditEngine.list_doc_types()
        if not doc_types:
            print("Нет доступных типов документов в doc_configs/")
            print("Создайте папку doc_configs/<doc_type>/ с config.json, rules.json и chunks_vision.json")
            sys.exit(0)

        print(f"\n📋 Доступные типы документов ({len(doc_types)}):")
        print(f"{'─'*50}")
        for dt in doc_types:
            title = f" — {dt['doc_title']}" if dt['doc_title'] else ""
            print(f"  {dt['doc_type']}{title}")
        print(f"\nИспользование: python run_audit.py --doc-type <doc_type> --target <file>")
        sys.exit(0)

    # Валидация аргументов
    if not args.doc_type:
        parser.error("--doc-type обязателен (или используйте --list-types)")

    if not args.target:
        parser.error("--target обязателен")

    # Проверяем, является ли doc_type специальным движком
    engine_type = _detect_engine(args.doc_type)

    if engine_type in SPECIAL_ENGINES:
        # Специальный движок (drivers, kpsc) — вызываем напрямую
        module = importlib.import_module(SPECIAL_ENGINES[engine_type])
        result = module.run(args)
    else:
        # Стандартный Vision pipeline
        # Для Vision pipeline rule_filter должен быть int
        rule_filter = int(args.rule_filter) if args.rule_filter is not None else None
        engine = AuditEngine(args.doc_type)
        result = engine.run(
            target_path=args.target,
            template_path=args.template,
            model=args.model,
            temperature=args.temperature,
            rule_filter=rule_filter,
            parse_only=args.parse_only,
            no_cache=args.no_cache,
            print_prompts=args.print_prompts,
            session_dir=args.session_dir,
            chunk_filter=args.chunk_filter,
            out_xlsx=args.out_xlsx,
            secondary_path=args.secondary
        )

    # Код возврата: 0 если нет нарушений, 1 если есть
    sys.exit(1 if result.violations else 0)


if __name__ == "__main__":
    main()
