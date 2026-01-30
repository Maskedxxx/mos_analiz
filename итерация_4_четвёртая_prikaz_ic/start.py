#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Стартовый скрипт для LLM-аудита документа "Приказ о создании ИЦ".

Использование:
  python start.py <документ> [шаблон]

Примеры:
  python start.py prikaz_ic_ok.docx
  python start.py prikaz_ic_errors.docx
  python start.py prikaz_ic_ok.docx prikaz_ic_template.docx
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

# Шаблон по умолчанию
DEFAULT_TEMPLATE = "prikaz_ic_template.docx"

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
        print(f"Найдено несколько файлов:")
        for m in matches:
            print(f"  - {m.name}")
        sys.exit(1)
    else:
        print(f"Документ не найден: {filename}")
        print(f"Доступные документы в {DOCUMENTS_DIR}:")
        for f in DOCUMENTS_DIR.glob("*.docx"):
            if not f.name.startswith("~$"):
                print(f"  - {f.name}")
        sys.exit(1)


def find_template(filename: str) -> Path:
    """Ищет шаблон в папке templates/"""
    path = TEMPLATES_DIR / filename
    if path.exists():
        return path

    print(f"Шаблон не найден: {filename}")
    print(f"Доступные шаблоны в {TEMPLATES_DIR}:")
    for f in TEMPLATES_DIR.glob("*.docx"):
        print(f"  - {f.name}")
    sys.exit(1)


def main():
    if len(sys.argv) < 2:
        print("Использование: python start.py <документ> [шаблон]")
        print()
        print("Документы (documents/):")
        for f in DOCUMENTS_DIR.glob("*.docx"):
            if not f.name.startswith("~$"):
                print(f"  - {f.name}")
        print()
        print("Шаблоны (templates/):")
        for f in TEMPLATES_DIR.glob("*.docx"):
            print(f"  - {f.name}")
        print()
        print(f"Шаблон по умолчанию: {DEFAULT_TEMPLATE}")
        sys.exit(1)

    # Получаем аргументы
    doc_name = sys.argv[1]
    template_name = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_TEMPLATE

    # Находим файлы
    doc_path = find_document(doc_name)
    template_path = find_template(template_name)

    print(f"Документ: {doc_path.name}")
    print(f"Шаблон:   {template_path.name}")
    print()

    # Запускаем аудит
    audit_script = SCRIPT_DIR / "scripts" / "run_prikaz_ic_llm_audit.py"

    env = os.environ.copy()
    env["OPENAI_API_KEY"] = env.get("OPENAI_API_KEY", OPENAI_API_KEY)

    cmd = [
        sys.executable,
        str(audit_script),
        "--target", str(doc_path),
        "--template", str(template_path),
        "--model", MODEL
    ]

    result = subprocess.run(cmd, env=env)
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
