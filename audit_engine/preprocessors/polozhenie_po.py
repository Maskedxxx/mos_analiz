#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Препроцессоры для Положения о ПО (polozhenie_po).

extract_longest_bullet_list — извлекает только самый длинный маркированный
список (•) из чанка пункт_1_5. Удаляет короткие/фантомные списки,
которые путают LLM при сравнении.

Проблема: Vision-парсер часто извлекает ДВА списка из чанка:
  - Короткий (п.1.5, 4-7 пунктов) — неполный
  - Длинный (п.1.7, 10 пунктов) — полный, нужный для сравнения
LLM игнорирует инструкцию "бери длинный" и сравнивает короткий.
Препроцессор решает это на уровне данных.
"""

from .registry import register_preprocessor


@register_preprocessor("polozhenie_po", "пункт_1_5")
def extract_longest_bullet_list(text: str) -> str:
    """
    Извлекает самый длинный маркированный список (•) из чанка.

    Алгоритм:
    1. Находит все группы последовательных строк с маркером •
    2. Выбирает группу с наибольшим количеством пунктов
    3. Возвращает заголовок секции + этот список
    """
    lines = text.split('\n')

    # Собираем группы bullet-строк с их индексами
    groups = []  # [(start_idx, end_idx, bullet_count)]
    current_start = None
    bullet_count = 0

    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith('•'):
            if current_start is None:
                current_start = i
            bullet_count += 1
        else:
            if current_start is not None:
                groups.append((current_start, i - 1, bullet_count))
                current_start = None
                bullet_count = 0

    # Последняя группа
    if current_start is not None:
        groups.append((current_start, len(lines) - 1, bullet_count))

    if not groups:
        return text  # Нет bullet-списков — вернуть как есть

    # Находим самую длинную группу
    best = max(groups, key=lambda g: g[2])
    best_start, best_end, _ = best

    # Ищем заголовок секции перед списком (строка с "руководствуется:")
    header_line = ""
    for i in range(best_start - 1, max(best_start - 5, -1), -1):
        if i < 0:
            break
        stripped = lines[i].strip()
        if stripped and not stripped.startswith('•'):
            header_line = stripped
            break

    # Собираем результат: заголовок + длинный список
    result_lines = []
    if header_line:
        result_lines.append(header_line)
    result_lines.extend(lines[best_start:best_end + 1])

    return '\n'.join(result_lines)
