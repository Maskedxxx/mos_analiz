#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Нечёткий поиск листов в Excel-файлах КПСЦ.

Решает проблему: парсеры хардкодили точные имена листов, но у реальных компаний
листы называются по-разному (trailing spaces, Latin C вместо кириллической С,
суффиксы типа "ВПП", "ТС", "(2)" и т.д.).

Используется всеми 9 парсерами kpsc через find_sheet().
"""

import re
from typing import Optional, List

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet


# Маппинг Latin → Cyrillic для нормализации (часто путают C и С, P и Р и т.д.)
_LATIN_TO_CYRILLIC = str.maketrans({
    'A': 'А', 'B': 'В', 'C': 'С', 'E': 'Е', 'H': 'Н',
    'K': 'К', 'M': 'М', 'O': 'О', 'P': 'Р', 'T': 'Т',
    'X': 'Х',
    'a': 'а', 'c': 'с', 'e': 'е', 'o': 'о', 'p': 'р',
    'x': 'х',
})


def _normalize(name: str) -> str:
    """Нормализация имени листа: strip, lower, Latin→Cyrillic, убираем спец-символы."""
    name = name.strip().lower()
    name = name.translate(_LATIN_TO_CYRILLIC)
    return name


def find_sheet(
    wb: Workbook,
    keywords: List[str],
    *,
    exclude_keywords: Optional[List[str]] = None,
    prefer_keywords: Optional[List[str]] = None,
) -> Optional[Worksheet]:
    """
    Нечёткий поиск листа в workbook по ключевым словам.

    Алгоритм:
    1. Нормализуем имена листов (strip, lower, Latin→Cyrillic)
    2. Ищем листы, содержащие ВСЕ keywords
    3. Исключаем листы с exclude_keywords
    4. Из оставшихся предпочитаем с prefer_keywords
    5. Из финальных кандидатов выбираем лист с максимумом данных

    Args:
        wb: openpyxl Workbook
        keywords: обязательные подстроки (нормализованные, lowercase)
        exclude_keywords: исключающие подстроки (если есть — лист отбрасывается)
        prefer_keywords: предпочтительные подстроки (приоритет при нескольких кандидатах)

    Returns:
        Worksheet или None, если не найден
    """
    if exclude_keywords is None:
        exclude_keywords = []
    if prefer_keywords is None:
        prefer_keywords = []

    # Нормализуем keywords
    kw_norm = [_normalize(kw) for kw in keywords]
    excl_norm = [_normalize(ek) for ek in exclude_keywords]
    pref_norm = [_normalize(pk) for pk in prefer_keywords]

    # Шаг 1: Собираем кандидатов — листы, содержащие ВСЕ keywords
    candidates = []
    for sheet_name in wb.sheetnames:
        name_norm = _normalize(sheet_name)

        # Проверяем, что ВСЕ ключевые слова есть в имени
        if not all(kw in name_norm for kw in kw_norm):
            continue

        # Проверяем exclude
        if any(ek in name_norm for ek in excl_norm):
            continue

        candidates.append(sheet_name)

    if not candidates:
        return None

    if len(candidates) == 1:
        return wb[candidates[0]]

    # Шаг 2: Предпочитаем с prefer_keywords
    if pref_norm:
        preferred = []
        for sn in candidates:
            name_norm = _normalize(sn)
            if any(pk in name_norm for pk in pref_norm):
                preferred.append(sn)
        if preferred:
            candidates = preferred

    if len(candidates) == 1:
        return wb[candidates[0]]

    # Шаг 3: Из оставшихся выбираем лист с максимумом заполненных ячеек
    best_sheet = None
    best_count = -1
    for sn in candidates:
        ws = wb[sn]
        # Считаем непустые ячейки в первых 20 строках (быстрая эвристика)
        count = 0
        for row in ws.iter_rows(min_row=1, max_row=min(20, ws.max_row or 1)):
            for cell in row:
                if cell.value not in (None, ""):
                    count += 1
        if count > best_count:
            best_count = count
            best_sheet = sn

    return wb[best_sheet] if best_sheet else None


def find_sheet_or_raise(
    wb: Workbook,
    keywords: List[str],
    parser_name: str,
    **kwargs,
) -> Worksheet:
    """
    find_sheet() с выбросом исключения если лист не найден.

    Args:
        wb: openpyxl Workbook
        keywords: ключевые слова для поиска
        parser_name: имя парсера (для сообщения об ошибке)
        **kwargs: дополнительные аргументы для find_sheet()

    Returns:
        Worksheet

    Raises:
        ValueError: если лист не найден
    """
    ws = find_sheet(wb, keywords, **kwargs)
    if ws is None:
        raise ValueError(
            f"[{parser_name}] Лист не найден по keywords={keywords} "
            f"среди {wb.sheetnames}"
        )
    return ws
