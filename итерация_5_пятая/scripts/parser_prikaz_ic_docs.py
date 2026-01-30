#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсер документа "Приказ о создании ИЦ" (информационного центра).

Разбивает документ на 6 чанков по маркерам.
Каждый чанк — raw текст без дополнительной обработки.

Чанки:
1. шапка — от начала до "О создании информационного центра предприятия"
2. текст_приказа — от "О создании..." до "Приложение № 1 к приказу"
3. приложение_1_к_приказу — до "Приложение № 2 к приказу"
4. приложение_2_к_приказу — до "Приложение № 1 к Регламенту" (содержит РЕГЛАМЕНТ)
5. приложение_1_к_регламенту — до "ЛИСТ ОЗНАКОМЛЕНИЯ"
6. лист_ознакомления — до конца документа
"""

import json
import argparse
import re
from pathlib import Path
from difflib import SequenceMatcher
from typing import List, Dict, Any
from docx import Document

# Маркеры для разделения документа на чанки
MARKERS = {
    "заголовок_приказа": "О создании информационного центра предприятия",
    "приложение_1_к_приказу": "Приложение № 1",
    "приложение_2_к_приказу": "Приложение № 2",
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


def get_paragraph_numbering(para) -> tuple:
    """
    Получает информацию о нумерации параграфа.

    Returns:
        (is_numbered, num_id, ilvl) — есть ли нумерация, ID списка, уровень вложенности
    """
    ppr = para._p.pPr if para._p is not None else None
    if ppr is None or ppr.numPr is None:
        return False, None, None

    num_id = ppr.numPr.numId.val if ppr.numPr.numId is not None else None
    ilvl = ppr.numPr.ilvl.val if ppr.numPr.ilvl is not None else 0

    return True, num_id, ilvl


def extract_all_text(doc: Document) -> List[str]:
    """
    Извлекает весь текст из документа: параграфы + текст из таблиц.
    Сохраняет нумерацию пунктов из Word.
    Возвращает список строк в порядке появления.
    """
    paragraphs = []

    # Счётчики для нумерации: {(num_id, ilvl): counter}
    numbering_counters = {}
    last_num_id = None

    # Извлекаем параграфы с сохранением нумерации
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        is_numbered, num_id, ilvl = get_paragraph_numbering(para)

        if is_numbered and num_id is not None:
            # Сбрасываем счётчик если начался новый список
            if num_id != last_num_id:
                numbering_counters = {}
                last_num_id = num_id

            # Увеличиваем счётчик для этого уровня
            key = (num_id, ilvl)
            if key not in numbering_counters:
                numbering_counters[key] = 0
            numbering_counters[key] += 1

            # Сбрасываем счётчики более глубоких уровней
            for k in list(numbering_counters.keys()):
                if k[0] == num_id and k[1] > ilvl:
                    del numbering_counters[k]

            # Формируем номер пункта
            counter = numbering_counters[key]
            if ilvl == 0:
                prefix = f"{counter}. "
            else:
                # Для вложенных уровней используем точечную нотацию
                prefix = f"{counter}. "

            text = prefix + text

        paragraphs.append(text)

    # Извлекаем текст из таблиц и вставляем примерно в нужные места
    # (упрощённо: таблица 0 — шапка, остальные — по порядку)
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


def parse_prikaz_ic(docx_path: str) -> Dict[str, Any]:
    """
    Парсит документ "Приказ о создании ИЦ".

    Args:
        docx_path: путь к .docx файлу

    Returns:
        Словарь с распарсенными чанками (raw текст)
    """
    doc = Document(docx_path)
    path = Path(docx_path)

    # Извлекаем параграфы и текст таблиц
    paragraphs, tables_text = extract_all_text(doc)

    # Находим индексы маркеров
    idx_заголовок = find_marker_index(paragraphs, MARKERS["заголовок_приказа"])

    idx_прил1 = find_marker_index(
        paragraphs,
        MARKERS["приложение_1_к_приказу"],
        start_from=idx_заголовок + 1 if idx_заголовок >= 0 else 0,
        next_line_contains="приказ"
    )

    idx_прил2 = find_marker_index(
        paragraphs,
        MARKERS["приложение_2_к_приказу"],
        start_from=idx_прил1 + 1 if idx_прил1 >= 0 else 0,
        next_line_contains="приказ"
    )

    idx_прил1_рег = find_marker_index(
        paragraphs,
        MARKERS["приложение_1_к_регламенту"],
        start_from=idx_прил2 + 1 if idx_прил2 >= 0 else 0
    )

    idx_лист = find_marker_index(
        paragraphs,
        MARKERS["лист_ознакомления"],
        start_from=idx_прил1_рег + 1 if idx_прил1_рег >= 0 else 0,
        next_line_contains="ознакомлен"
    )

    # === Формируем чанки (raw текст) ===

    # 1. Шапка: таблица 0 + параграфы до заголовка
    шапка_parts = []
    if tables_text:
        шапка_parts.append(tables_text[0])  # Таблица реквизитов
    if idx_заголовок > 0:
        шапка_parts.extend(paragraphs[:idx_заголовок])
    шапка = '\n'.join(шапка_parts)

    # 2. Текст приказа
    if idx_заголовок >= 0 and idx_прил1 > idx_заголовок:
        текст_приказа = '\n'.join(paragraphs[idx_заголовок:idx_прил1])
    elif idx_заголовок >= 0:
        текст_приказа = '\n'.join(paragraphs[idx_заголовок:])
    else:
        текст_приказа = ""

    # 3. Приложение 1 к приказу
    if idx_прил1 >= 0 and idx_прил2 > idx_прил1:
        прил1_text = '\n'.join(paragraphs[idx_прил1:idx_прил2])
        # Добавляем таблицы 1 и 2 (структура ИЦ и показатели)
        if len(tables_text) > 1:
            прил1_text += '\n\n' + tables_text[1]
        if len(tables_text) > 2:
            прил1_text += '\n\n' + tables_text[2]
    else:
        прил1_text = ""

    # 4. Приложение 2 к приказу (регламент)
    if idx_прил2 >= 0 and idx_прил1_рег > idx_прил2:
        прил2_text = '\n'.join(paragraphs[idx_прил2:idx_прил1_рег])
        # Добавляем таблицу 3 (термины)
        if len(tables_text) > 3:
            прил2_text += '\n\n' + tables_text[3]
    elif idx_прил2 >= 0:
        прил2_text = '\n'.join(paragraphs[idx_прил2:])
    else:
        прил2_text = ""

    # 5. Приложение 1 к регламенту (матрица ответственности)
    if idx_прил1_рег >= 0 and idx_лист > idx_прил1_рег:
        прил1_рег_text = '\n'.join(paragraphs[idx_прил1_рег:idx_лист])
        # Добавляем таблицу 4 (матрица)
        if len(tables_text) > 4:
            прил1_рег_text += '\n\n' + tables_text[4]
    elif idx_прил1_рег >= 0:
        прил1_рег_text = '\n'.join(paragraphs[idx_прил1_рег:])
    else:
        прил1_рег_text = ""

    # 6. Лист ознакомления
    if idx_лист >= 0:
        лист_text = '\n'.join(paragraphs[idx_лист:])
        # Добавляем таблицу 6 (лист ознакомления)
        if len(tables_text) > 6:
            лист_text += '\n\n' + tables_text[6]
    else:
        лист_text = ""

    # Формируем результат
    result = {
        "имя_файла": path.name,
        "путь": str(path.absolute()),
        "шапка": шапка,
        "текст_приказа": текст_приказа,
        "приложение_1_к_приказу": прил1_text,
        "приложение_2_к_приказу": прил2_text,
        "приложение_1_к_регламенту": прил1_рег_text,
        "лист_ознакомления": лист_text,
    }

    return result


def main():
    """Точка входа CLI."""
    parser = argparse.ArgumentParser(
        description="Парсер документа 'Приказ о создании ИЦ'"
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
    result = parse_prikaz_ic(args.input)

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

    return result


if __name__ == "__main__":
    main()
