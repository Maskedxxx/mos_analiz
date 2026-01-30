#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсер документа "Приказ о создании ИЦ потока".

Разбивает документ на 7 чанков по маркерам:
1. шапка — таблица реквизитов (организация, ПРИКАЗ, дата)
2. преамбула — от заголовка "О создании..." до "ПРИКАЗЫВАЮ"
3. текст_приказа — пункты 1-5 с распоряжениями + подписант
4. приложение_1 — форма инфоцентра потока + таблицы показателей
5. приложение_2_регламент — регламент работы (разделы 1-4)
6. лист_ознакомления — таблица подписей
7. подписант — должность + ФИО подписанта (из конца текста_приказа)
"""

import json
import argparse
import re
from pathlib import Path
from difflib import SequenceMatcher
from typing import List, Dict, Any, Tuple
from docx import Document


# Маркеры для разделения документа на чанки
MARKERS = {
    "заголовок_приказа": "О создании информационного центра пилотного потока",
    "приказываю": "ПРИКАЗЫВАЮ",
    "приложение_1": "Приложение № 1",
    "приложение_2": "Приложение № 2",
    "регламент": "РЕГЛАМЕНТ",
    "приложение_1_к_регламенту": "Приложение № 1 к Регламенту",
    "лист_ознакомления": "ЛИСТ",
}

# Порог fuzzy matching
FUZZY_THRESHOLD = 0.80


def fuzzy_match(text: str, pattern: str, threshold: float = FUZZY_THRESHOLD) -> bool:
    """
    Проверяет fuzzy совпадение текста с паттерном.
    """
    text_lower = text.lower().strip()
    pattern_lower = pattern.lower().strip()

    # Точное вхождение
    if pattern_lower in text_lower:
        return True

    # Fuzzy matching
    ratio = SequenceMatcher(None, text_lower, pattern_lower).ratio()
    if ratio >= threshold:
        return True

    # Проверка начала строки
    if len(text_lower) >= len(pattern_lower):
        start_ratio = SequenceMatcher(
            None,
            text_lower[:len(pattern_lower) + 10],
            pattern_lower
        ).ratio()
        if start_ratio >= threshold:
            return True

    return False


def find_marker_index(paragraphs: List[str], marker: str, start_from: int = 0,
                      next_line_contains: str = None) -> int:
    """
    Ищет индекс параграфа, содержащего маркер (с fuzzy matching).

    Args:
        paragraphs: список параграфов
        marker: искомый маркер
        start_from: начать поиск с этого индекса
        next_line_contains: если указано, следующая строка должна содержать эту подстроку

    Возвращает -1 если не найден.
    """
    for i in range(start_from, len(paragraphs)):
        if fuzzy_match(paragraphs[i], marker):
            if next_line_contains:
                if i + 1 < len(paragraphs):
                    next_text = paragraphs[i + 1].lower()
                    if next_line_contains.lower() not in next_text:
                        continue
            return i
    return -1


def extract_all_text(doc: Document) -> Tuple[List[str], List[str]]:
    """
    Извлекает весь текст из документа: параграфы + текст из таблиц.
    Сохраняет нумерацию пунктов из Word.
    Возвращает (список параграфов, список таблиц).
    """
    paragraphs = []

    # Извлекаем параграфы
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            paragraphs.append(text)

    # Извлекаем текст из таблиц
    tables_text = []
    for table in doc.tables:
        table_lines = []
        for row in table.rows:
            row_text = []
            for cell in row.cells:
                cell_text = cell.text.strip()
                if cell_text:
                    row_text.append(cell_text)
            if row_text:
                table_lines.append(' | '.join(row_text))
        if table_lines:
            tables_text.append('\n'.join(table_lines))

    return paragraphs, tables_text


def extract_signatory(text_prikazа: str) -> str:
    """
    Извлекает информацию о подписанте из текста приказа.
    Ищет строку с должностью и ФИО в конце текста.
    """
    lines = text_prikazа.strip().split('\n')

    # Ищем строку с подписантом (обычно последняя непустая строка)
    for line in reversed(lines):
        line = line.strip()
        if not line:
            continue
        # Признаки подписанта: должность + ФИО
        if ('директор' in line.lower() or 'должность' in line.lower() or
            'фамилия' in line.lower() or 'фио' in line.lower() or
            re.search(r'[А-ЯЁ]\.[А-ЯЁ]\.', line) or  # Инициалы
            re.search(r'\s{5,}', line)):  # Много пробелов
            return line

    return ""


def parse_prikaz_ic_potoka(docx_path: str) -> Dict[str, Any]:
    """
    Парсит документ "Приказ о создании ИЦ потока".

    Args:
        docx_path: путь к .docx файлу

    Returns:
        Словарь с распарсенными чанками (raw текст)
    """
    doc = Document(docx_path)
    path = Path(docx_path)

    # Извлекаем параграфы и текст таблиц
    paragraphs, tables_text = extract_all_text(doc)

    # === Находим индексы маркеров ===

    # Заголовок приказа
    idx_заголовок = find_marker_index(paragraphs, MARKERS["заголовок_приказа"])

    # ПРИКАЗЫВАЮ
    idx_приказываю = find_marker_index(
        paragraphs,
        MARKERS["приказываю"],
        start_from=idx_заголовок + 1 if idx_заголовок >= 0 else 0
    )

    # Приложение № 1
    idx_прил1 = find_marker_index(
        paragraphs,
        MARKERS["приложение_1"],
        start_from=idx_приказываю + 1 if idx_приказываю >= 0 else 0,
        next_line_contains="приказ"
    )

    # Приложение № 2
    idx_прил2 = find_marker_index(
        paragraphs,
        MARKERS["приложение_2"],
        start_from=idx_прил1 + 1 if idx_прил1 >= 0 else 0,
        next_line_contains="приказ"
    )

    # РЕГЛАМЕНТ (внутри Приложения 2)
    idx_регламент = find_marker_index(
        paragraphs,
        MARKERS["регламент"],
        start_from=idx_прил2 + 1 if idx_прил2 >= 0 else 0
    )

    # Приложение № 1 к Регламенту (матрица ответственности)
    idx_прил1_рег = find_marker_index(
        paragraphs,
        MARKERS["приложение_1_к_регламенту"],
        start_from=idx_регламент + 1 if idx_регламент >= 0 else 0
    )

    # Лист ознакомления
    idx_лист = find_marker_index(
        paragraphs,
        MARKERS["лист_ознакомления"],
        start_from=idx_прил1_рег + 1 if idx_прил1_рег >= 0 else 0,
        next_line_contains="ознакомлен"
    )

    # === Формируем чанки (raw текст) ===

    # 1. Шапка: таблица 0 (реквизиты)
    шапка = tables_text[0] if tables_text else ""

    # 2. Преамбула: от заголовка до ПРИКАЗЫВАЮ
    if idx_заголовок >= 0 and idx_приказываю > idx_заголовок:
        преамбула = '\n'.join(paragraphs[idx_заголовок:idx_приказываю])
    elif idx_заголовок >= 0:
        преамбула = paragraphs[idx_заголовок] if idx_заголовок < len(paragraphs) else ""
    else:
        преамбула = ""

    # 3. Текст приказа: от ПРИКАЗЫВАЮ до Приложения 1
    if idx_приказываю >= 0 and idx_прил1 > idx_приказываю:
        текст_приказа = '\n'.join(paragraphs[idx_приказываю:idx_прил1])
    elif idx_приказываю >= 0:
        текст_приказа = '\n'.join(paragraphs[idx_приказываю:])
    else:
        текст_приказа = ""

    # 7. Подписант: извлекаем из текста приказа
    подписант = extract_signatory(текст_приказа)

    # 4. Приложение 1: форма инфоцентра + таблицы показателей
    if idx_прил1 >= 0 and idx_прил2 > idx_прил1:
        прил1_text = '\n'.join(paragraphs[idx_прил1:idx_прил2])
        # Добавляем таблицы 1 и 2 (показатели)
        if len(tables_text) > 1:
            прил1_text += '\n\n[ТАБЛИЦА ПОКАЗАТЕЛЕЙ]\n' + tables_text[1]
        if len(tables_text) > 2:
            прил1_text += '\n\n[ТАБЛИЦА ДОКУМЕНТОВ]\n' + tables_text[2]
    else:
        прил1_text = ""

    # 5. Приложение 2 (Регламент): от Приложения 2 до Листа ознакомления
    if idx_прил2 >= 0 and idx_лист > idx_прил2:
        прил2_text = '\n'.join(paragraphs[idx_прил2:idx_лист])
        # Добавляем таблицу 3 (термины) и 4 (матрица)
        if len(tables_text) > 3:
            прил2_text += '\n\n[ТАБЛИЦА ТЕРМИНОВ]\n' + tables_text[3]
        if len(tables_text) > 4:
            прил2_text += '\n\n[МАТРИЦА ОТВЕТСТВЕННОСТИ]\n' + tables_text[4]
    elif idx_прил2 >= 0:
        прил2_text = '\n'.join(paragraphs[idx_прил2:])
    else:
        прил2_text = ""

    # 6. Лист ознакомления
    if idx_лист >= 0:
        лист_text = '\n'.join(paragraphs[idx_лист:])
        # Добавляем таблицу 6 (лист ознакомления)
        if len(tables_text) > 6:
            лист_text += '\n\n[ТАБЛИЦА ОЗНАКОМЛЕНИЯ]\n' + tables_text[6]
    else:
        лист_text = ""

    # Формируем результат
    result = {
        "имя_файла": path.name,
        "путь": str(path.absolute()),
        "шапка": шапка,
        "преамбула": преамбула,
        "текст_приказа": текст_приказа,
        "приложение_1": прил1_text,
        "приложение_2_регламент": прил2_text,
        "лист_ознакомления": лист_text,
        "подписант": подписант,
    }

    return result


def main():
    """Точка входа CLI."""
    parser = argparse.ArgumentParser(
        description="Парсер документа 'Приказ о создании ИЦ потока'"
    )
    parser.add_argument(
        "input",
        help="Путь к входному .docx файлу"
    )
    parser.add_argument(
        "-o", "--output",
        help="Путь для сохранения JSON результата"
    )

    args = parser.parse_args()

    # Парсим документ
    result = parse_prikaz_ic_potoka(args.input)

    # Определяем путь для сохранения
    if args.output:
        output_path = Path(args.output)
    else:
        script_dir = Path(__file__).parent.parent
        results_dir = script_dir / "results"
        results_dir.mkdir(exist_ok=True)
        input_name = Path(args.input).stem
        output_path = results_dir / f"{input_name}.json"

    # Сохраняем результат
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"✅ Результат сохранён: {output_path}")

    # Выводим краткую информацию о чанках
    print("\n📄 Извлечённые чанки:")
    for key, value in result.items():
        if key in ("имя_файла", "путь"):
            continue
        length = len(value) if value else 0
        preview = value[:80].replace('\n', ' ') if value else "(пусто)"
        print(f"  - {key}: {length} символов | {preview}...")

    return result


if __name__ == "__main__":
    main()
