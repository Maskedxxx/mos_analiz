#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Стартовый скрипт для LLM-аудита документа "Приказ о создании ИЦ потока".

Использование:
  python start.py <документ> [--variant paper|electronic]

Примеры:
  # Бумажный вариант (по умолчанию)
  python start.py prikaz_ic_potoka_ok.docx

  # Электронный вариант
  python start.py prikaz_ic_potoka_el_ok.docx --variant electronic

  # С указанием шаблона
  python start.py prikaz_ic_potoka_ok.docx --template prikaz_ic_potoka_template.docx
"""

import os
import sys
import subprocess
from pathlib import Path

# ============================================================================
# КОНФИГУРАЦИЯ
# ============================================================================

# Директории относительно скрипта
SCRIPT_DIR = Path(__file__).parent
DOCUMENTS_DIR = SCRIPT_DIR / "documents"
TEMPLATES_DIR = SCRIPT_DIR / "templates"

# Шаблоны по умолчанию
DEFAULT_TEMPLATE_PAPER = "prikaz_ic_potoka_template.docx"
DEFAULT_TEMPLATE_ELECTRONIC = "prikaz_ic_potoka_el_template.docx"

# API ключ OpenAI
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")

# Модель OpenAI
MODEL = "gpt-4.1-mini"

# ============================================================================


def find_document(filename: str) -> Path:
    """Ищет документ в папке documents/"""
    # Точное совпадение
    path = DOCUMENTS_DIR / filename
    if path.exists():
        return path

    # Поиск по частичному имени
    matches = list(DOCUMENTS_DIR.glob(f"*{filename}*"))
    matches = [m for m in matches if not m.name.startswith("~$")]

    if len(matches) == 1:
        return matches[0]
    elif len(matches) > 1:
        print(f"⚠️  Найдено несколько файлов:")
        for m in matches:
            print(f"    - {m.name}")
        sys.exit(1)
    else:
        print(f"❌ Документ не найден: {filename}")
        print(f"📁 Доступные документы в {DOCUMENTS_DIR}:")
        for f in DOCUMENTS_DIR.glob("*.docx"):
            if not f.name.startswith("~$"):
                print(f"    - {f.name}")
        sys.exit(1)


def find_template(filename: str) -> Path:
    """Ищет шаблон в папке templates/"""
    path = TEMPLATES_DIR / filename
    if path.exists():
        return path

    print(f"❌ Шаблон не найден: {filename}")
    print(f"📁 Доступные шаблоны в {TEMPLATES_DIR}:")
    for f in TEMPLATES_DIR.glob("*.docx"):
        print(f"    - {f.name}")
    sys.exit(1)


def detect_variant(doc_name: str) -> str:
    """Определяет вариант документа по имени файла."""
    doc_lower = doc_name.lower()
    if "эл" in doc_lower or "_el" in doc_lower:
        return "electronic"
    return "paper"


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="LLM-аудит документа 'Приказ о создании ИЦ потока'"
    )
    parser.add_argument(
        "document",
        nargs="?",
        help="Имя целевого документа из папки documents/"
    )
    parser.add_argument(
        "--variant",
        choices=["paper", "electronic"],
        default=None,
        help="Вариант документа (автоопределяется если не указано)"
    )
    parser.add_argument(
        "--template",
        default=None,
        help="Имя шаблона из папки templates/"
    )
    parser.add_argument(
        "--print-prompts",
        action="store_true",
        help="Режим отладки: вывести промпты без вызова LLM"
    )
    parser.add_argument(
        "--use-vision",
        action="store_true",
        help="Использовать Vision Pipeline (gpt-4.1-mini) вместо python-docx парсера"
    )

    args = parser.parse_args()

    # Показываем справку если документ не указан
    if not args.document:
        print("📋 LLM-аудит документа 'Приказ о создании ИЦ потока'")
        print()
        print("Использование: python start.py <документ> [--variant paper|electronic]")
        print()
        print("📁 Документы (documents/):")
        for f in DOCUMENTS_DIR.glob("*.docx"):
            if not f.name.startswith("~$"):
                variant_hint = " (electronic)" if "el" in f.name.lower() else " (paper)"
                print(f"    - {f.name}{variant_hint}")
        print()
        print("📁 Шаблоны (templates/):")
        for f in TEMPLATES_DIR.glob("*.docx"):
            print(f"    - {f.name}")
        print()
        print("Примеры:")
        print("  python start.py prikaz_ic_potoka_ok.docx")
        print("  python start.py prikaz_ic_potoka_el_ok.docx --variant electronic")
        sys.exit(0)

    # Находим документ
    doc_path = find_document(args.document)

    # Определяем вариант
    variant = args.variant or detect_variant(doc_path.name)

    # Определяем шаблон
    if args.template:
        template_name = args.template
    else:
        template_name = DEFAULT_TEMPLATE_ELECTRONIC if variant == "electronic" else DEFAULT_TEMPLATE_PAPER

    template_path = find_template(template_name)

    # Выводим информацию
    print(f"📄 Документ:  {doc_path.name}")
    print(f"📋 Шаблон:    {template_path.name}")
    print(f"📝 Вариант:   {variant}")
    print()

    # Запускаем аудит
    audit_script = SCRIPT_DIR / "scripts" / "run_prikaz_ic_potoka_audit.py"

    env = os.environ.copy()
    env["OPENAI_API_KEY"] = env.get("OPENAI_API_KEY", OPENAI_API_KEY)

    cmd = [
        sys.executable,
        str(audit_script),
        "--target", str(doc_path),
        "--template", str(template_path),
        "--variant", variant,
        "--model", MODEL
    ]

    if args.print_prompts:
        cmd.append("--print-prompts")

    if args.use_vision:
        cmd.append("--use-vision")

    result = subprocess.run(cmd, env=env)
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
