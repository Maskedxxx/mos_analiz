#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Стартовый скрипт для LLM-аудита документов:
- "Приказ о проведении выхода" (DOCX)
- "График обхода ОМ" (XLSX)

Использование:
  python start.py <docx_файл> <xlsx_файл>
  python start.py prikaz_vyhod_ok.docx grafik_obhod_ok.xlsx

Примеры:
  python start.py prikaz_vyhod_ok.docx grafik_obhod_ok.xlsx
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

# Шаблон DOCX по умолчанию
DEFAULT_TEMPLATE = "prikaz_vyhod_template.docx"

# API ключ OpenAI
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")

# Модель OpenAI
MODEL = "gpt-4.1-mini"

# ============================================================================


def find_file(filename: str, directory: Path, file_type: str) -> Path:
    """Ищет файл в указанной папке."""
    # Точное совпадение
    path = directory / filename
    if path.exists():
        return path

    # Поиск по частичному имени
    matches = list(directory.glob(f"*{filename}*"))
    matches = [m for m in matches if not m.name.startswith("~$")]

    if len(matches) == 1:
        return matches[0]
    elif len(matches) > 1:
        print(f"Найдено несколько {file_type} файлов:")
        for m in matches:
            print(f"  - {m.name}")
        sys.exit(1)
    else:
        print(f"{file_type} файл не найден: {filename}")
        print(f"Доступные файлы в {directory}:")
        for f in directory.iterdir():
            if not f.name.startswith("~$") and f.is_file():
                print(f"  - {f.name}")
        sys.exit(1)


def main():
    if len(sys.argv) < 3:
        print("Использование: python start.py <docx_файл> <xlsx_файл> [шаблон]")
        print()
        print("Документы (documents/):")
        for f in DOCUMENTS_DIR.iterdir():
            if not f.name.startswith("~$") and f.is_file():
                print(f"  - {f.name}")
        print()
        print("Шаблоны (templates/):")
        for f in TEMPLATES_DIR.iterdir():
            if not f.name.startswith("~$") and f.is_file():
                print(f"  - {f.name}")
        print()
        print(f"Шаблон DOCX по умолчанию: {DEFAULT_TEMPLATE}")
        sys.exit(1)

    # Получаем аргументы
    docx_name = sys.argv[1]
    xlsx_name = sys.argv[2]
    template_name = sys.argv[3] if len(sys.argv) > 3 else DEFAULT_TEMPLATE

    # Находим файлы
    docx_path = find_file(docx_name, DOCUMENTS_DIR, "DOCX")
    xlsx_path = find_file(xlsx_name, DOCUMENTS_DIR, "XLSX")
    template_path = find_file(template_name, TEMPLATES_DIR, "Шаблон")

    print(f"📄 DOCX:    {docx_path.name}")
    print(f"📊 XLSX:    {xlsx_path.name}")
    print(f"📋 Шаблон:  {template_path.name}")
    print()

    # Запускаем аудит
    audit_script = SCRIPT_DIR / "scripts" / "run_prikaz_vyhod_audit.py"

    env = os.environ.copy()
    if OPENAI_API_KEY:
        env["OPENAI_API_KEY"] = OPENAI_API_KEY

    cmd = [
        sys.executable,
        str(audit_script),
        "--docx", str(docx_path),
        "--xlsx", str(xlsx_path),
        "--template", str(template_path),
        "--model", MODEL
    ]

    result = subprocess.run(cmd, env=env)
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
