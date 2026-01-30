#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсер для документа "Приказ о проведении выхода" (.docx).

Извлекает 5 чанков:
- шапка: юридическая форма + наименование предприятия
- номер_дата: номер приказа + дата
- заголовок_город: заголовок приказа + город подписания
- текст_приказа: пункты 1-4 с "ПРИКАЗЫВАЮ"
- должность_фио_подписанта: подпись внизу документа

Использование:
    from parser_prikaz_vyhod_docx import parse_prikaz_vyhod
    chunks = parse_prikaz_vyhod("path/to/document.docx")
"""

import re
from pathlib import Path
from docx import Document


def parse_prikaz_vyhod(docx_path: str) -> dict:
    """
    Парсит документ "Приказ о проведении выхода".

    Args:
        docx_path: путь к .docx файлу

    Returns:
        dict с чанками: шапка, заголовок_город, текст_приказа, должность_фио_подписанта
    """
    doc = Document(docx_path)

    # Собираем весь текст в список строк
    lines = []
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            lines.append(text)

    # Инициализация чанков
    chunks = {
        "имя_файла": Path(docx_path).name,
        "шапка": "",
        "номер_дата": "",
        "заголовок_город": "",
        "текст_приказа": "",
        "должность_фио_подписанта": ""
    }

    # === Парсинг ===

    # Маркеры для определения границ
    idx_prikazyvayu = -1
    idx_prikaz_word = -1

    for i, line in enumerate(lines):
        # Ищем "ПРИКАЗЫВАЮ:" или "ПРИКАЗЫВАЮ"
        if "ПРИКАЗЫВАЮ" in line.upper():
            idx_prikazyvayu = i
        # Ищем слово "Приказ" (заголовок документа)
        if line.lower() == "приказ":
            idx_prikaz_word = i

    # --- шапка ---
    # Первая строка обычно содержит ООО/ЗАО/АО + наименование
    # Ищем строку с юр.формой
    for i, line in enumerate(lines):
        if re.search(r'\b(ООО|ЗАО|АО|ПАО|ИП)\b', line):
            chunks["шапка"] = line
            break

    # --- номер_дата ---
    # Строка с "№" и датой, обычно перед словом "Приказ"
    for i, line in enumerate(lines):
        if "№" in line or re.search(r'\d{1,2}[./]\d{1,2}[./]\d{2,4}', line) or \
           re.search(r'\d{1,2}\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)', line.lower()):
            # Пропускаем если это текст приказа (п.3 ссылка на приложение)
            if "приложение" in line.lower():
                continue
            chunks["номер_дата"] = line
            break

    # --- заголовок_город ---
    # От слова "Приказ" до "ПРИКАЗЫВАЮ" (не включая текст программы)
    if idx_prikaz_word >= 0:
        header_parts = []
        for i in range(idx_prikaz_word, min(idx_prikaz_word + 4, len(lines))):
            line = lines[i]
            # Пропускаем длинный текст "С целью обеспечения..."
            if "целью обеспечения" in line.lower():
                break
            header_parts.append(line)
        chunks["заголовок_город"] = "\n".join(header_parts)

    # --- текст_приказа ---
    # От "С целью..." до подписанта
    if idx_prikazyvayu >= 0:
        text_parts = []
        # Ищем начало текста (обычно "С целью...")
        start_idx = -1
        for i, line in enumerate(lines):
            if "целью обеспечения" in line.lower() or "целях" in line.lower():
                start_idx = i
                break

        if start_idx >= 0:
            # Собираем до последнего пункта (обычно п.4)
            for i in range(start_idx, len(lines)):
                line = lines[i]
                # Останавливаемся на подписанте
                if re.search(r'(директор|генеральный|руководитель)', line.lower()) and i > idx_prikazyvayu:
                    break
                text_parts.append(line)
            chunks["текст_приказа"] = "\n".join(text_parts)

    # --- должность_фио_подписанта ---
    # Последние строки документа (должность + ФИО)
    # Ищем снизу вверх
    for i in range(len(lines) - 1, -1, -1):
        line = lines[i]
        if re.search(r'(директор|руководитель|председатель)', line.lower()):
            chunks["должность_фио_подписанта"] = line
            break

    return chunks


# === Тест при запуске напрямую ===
if __name__ == "__main__":
    import sys

    # По умолчанию тестируем на файле ООО Пример
    project_dir = Path(__file__).parent.parent
    default_path = project_dir / "documents" / "prikaz_vyhod_ok.docx"

    if len(sys.argv) > 1:
        test_path = sys.argv[1]
    else:
        test_path = str(default_path)

    print(f"📄 Парсинг: {Path(test_path).name}")
    print("=" * 60)

    chunks = parse_prikaz_vyhod(test_path)

    for key, value in chunks.items():
        print(f"\n[{key}]")
        print("-" * 40)
        print(value if value else "(пусто)")
