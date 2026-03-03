#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Препроцессоры для Положения о ПО (polozhenie_po).

extract_longest_bullet_list — извлекает самый длинный маркированный список (•)
из чанка пункт_1_5 (tolerant к пустым строкам от Paddle OCR).

normalize_ознакомление — определяет заполненность листа ознакомления.
OCR не читает рукописный текст → мусорные символы в ячейках.
Препроцессор помечает строки с любым содержимым как "заполнено".
"""

import re

from .registry import register_preprocessor


@register_preprocessor("polozhenie_po", "пункт_1_5")
def extract_longest_bullet_list(text: str) -> str:
    """
    Извлекает самый длинный маркированный список (•) из чанка.

    Алгоритм:
    1. Находит все строки с маркером •
    2. Группирует близкие буллеты (пустые строки между ними — норма для Paddle OCR)
    3. Выбирает группу с наибольшим количеством пунктов
    4. Возвращает заголовок секции + этот список

    Args:
        text: полный текст чанка пункт_1_5
    Returns:
        str: заголовок + самый длинный bullet-список
    """
    lines = text.split('\n')

    # Собираем индексы всех bullet-строк
    bullet_indices = [i for i, line in enumerate(lines) if line.strip().startswith('•')]

    if not bullet_indices:
        return text  # Нет bullet-списков — вернуть как есть

    # Группируем буллеты: разрыв > 3 строк = новая группа
    # (Paddle OCR ставит 1-2 пустые строки между буллетами)
    MAX_GAP = 3
    groups = []  # [(bullet_indices_list)]
    current_group = [bullet_indices[0]]

    for idx in bullet_indices[1:]:
        if idx - current_group[-1] <= MAX_GAP:
            current_group.append(idx)
        else:
            groups.append(current_group)
            current_group = [idx]
    groups.append(current_group)

    # Находим самую длинную группу
    best_group = max(groups, key=len)
    best_start = best_group[0]
    best_end = best_group[-1]

    # Ищем заголовок секции перед списком
    header_line = ""
    for i in range(best_start - 1, max(best_start - 5, -1), -1):
        if i < 0:
            break
        stripped = lines[i].strip()
        if stripped and not stripped.startswith('•') and not stripped.startswith('['):
            header_line = stripped
            break

    # Собираем результат: заголовок + строки от первого до последнего буллета
    result_lines = []
    if header_line:
        result_lines.append(header_line)
    result_lines.extend(lines[best_start:best_end + 1])

    return '\n'.join(result_lines)


@register_preprocessor("polozhenie_po", "лист_ознакомления")
def normalize_ознакомление(text: str) -> str:
    """
    Нормализует лист ознакомления для корректной проверки заполненности.

    Проблема: Paddle OCR не читает рукописный текст — ФИО и даты
    превращаются в мусор типа "('r>4' 1A,nQ(((%4". LLM видит мусор
    и считает строки пустыми.

    Решение: если в ячейке ФИО есть любые символы (длина > 2) — помечаем
    строку как "заполнено (рукописный текст, OCR не распознал)".

    Args:
        text: HTML-таблица листа ознакомления
    Returns:
        str: текст с пометками о заполненности
    """
    # Ищем строки таблицы: <tr>...</tr>
    rows = re.findall(r'<tr>(.*?)</tr>', text, re.DOTALL)
    if len(rows) <= 1:
        return text  # Только заголовок или нет таблицы

    filled_count = 0
    # Пропускаем заголовок (первую строку)
    for row in rows[1:]:
        cells = re.findall(r'<td>(.*?)</td>', row, re.DOTALL)
        if len(cells) >= 2:
            # Ячейка ФИО — обычно вторая колонка
            fio_cell = cells[1].strip()
            # Убираем пробелы, переносы, тире — остаётся ли что-то?
            cleaned = re.sub(r'[\s\-_|/\\.,;:!?\'"()\[\]{}]+', '', fio_cell)
            if len(cleaned) >= 1:
                filled_count += 1

    # Добавляем сводку в конец текста
    if filled_count > 0:
        summary = f"\n\n[СВОДКА: заполнено строк: {filled_count} (рукописный текст, OCR распознал частично)]"
        return text + summary

    return text
