#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Non-LLM проверки для Чек-листа выбора ЭУ (cheklist_eu).

Правило #5: проверка заполнения оценок критериев
Правило #6: проверка итоговой оценки (сумма баллов)

Paddle-парсер выдаёт HTML-таблицы (<tr>/<td>), не markdown.
Структура строки критерия:
  <tr><td>N</td><td>текст критерия</td><td>оценка</td><td>...</td></tr>
"""

import re
from typing import Any, Dict, List, Optional, Tuple

from .registry import register


def _extract_rows(html: str) -> List[List[str]]:
    """
    Извлекает строки таблицы из HTML.

    Args:
        html: HTML-текст с <tr>/<td> тегами
    Returns:
        Список строк, каждая — список значений ячеек (stripped).
    """
    rows = []
    for tr_match in re.finditer(r'<tr[^>]*>(.*?)</tr>', html, re.DOTALL):
        tr_content = tr_match.group(1)
        cells = [
            cell.strip()
            for cell in re.findall(r'<td[^>]*>(.*?)</td>', tr_content, re.DOTALL)
        ]
        rows.append(cells)
    return rows


def _parse_criteria(table_text: str) -> Tuple[Dict[int, Optional[int]], Optional[int]]:
    """
    Извлекает оценки 7 критериев и итоговую оценку из HTML-таблицы.

    Структура строки критерия (6 ячеек):
      [номер(1-7), текст, оценка(0/1/2/пусто), описание_0, описание_1, описание_2]
    Строка итоговой оценки:
      [пусто, "Итоговая оценка", значение, ...]

    Args:
        table_text: HTML чанка таблица_критериев
    Returns:
        (criteria, total):
          criteria: {1: 2, 2: None, ...} — номер критерия → оценка (None = пустая)
          total: итоговая оценка (int) или None если не найдена
    """
    criteria = {}
    total = None

    for row in _extract_rows(table_text):
        if len(row) < 3:
            continue

        first_cell = row[0].strip()
        second_cell = row[1].strip()
        third_cell = row[2].strip()

        # Строка критерия: первая ячейка = число 1-7
        if first_cell in ('1', '2', '3', '4', '5', '6', '7'):
            num = int(first_cell)
            if third_cell in ('0', '1', '2'):
                criteria[num] = int(third_cell)
            else:
                criteria[num] = None

        # Строка "Итоговая оценка"
        if 'итоговая' in second_cell.lower() and 'оценка' in second_cell.lower():
            # Оценка может быть в 3-й ячейке или дальше
            for cell in row[2:]:
                cleaned = re.sub(r'[^0-9]', '', cell)
                if cleaned and 0 <= int(cleaned) <= 14:
                    total = int(cleaned)
                    break

    return criteria, total


@register("cheklist_eu", 5)
def check_criteria_scores(target_doc: Dict[str, Any], config: Any) -> List[Dict[str, Any]]:
    """
    Правило #5: проверка заполнения оценок критериев.

    Проверяет что все 7 критериев имеют оценку (0, 1 или 2).
    Парсит HTML-таблицу из чанка 'таблица_критериев'.

    Args:
        target_doc: распарсенный документ {чанк: текст}
        config: конфиг аудита
    Returns:
        список нарушений (пустой = всё ок)
    """
    table_text = target_doc.get("таблица_критериев", "")

    if not table_text:
        return [{
            "rule_index": 5,
            "rule_title": "Проверка заполнения таблицы критериев",
            "Целевой документ": "отсутствует",
            "Различие": "Не найден чанк 'таблица_критериев'"
        }]

    criteria, _ = _parse_criteria(table_text)
    violations = []

    for i in range(1, 8):
        if i not in criteria:
            violations.append({
                "rule_index": 5,
                "rule_title": "Проверка заполнения таблицы критериев",
                "Целевой документ": f"Критерий {i} не найден в таблице",
                "Различие": f"Критерий {i} должен присутствовать в таблице с оценкой 0, 1 или 2"
            })
        elif criteria[i] is None:
            violations.append({
                "rule_index": 5,
                "rule_title": "Проверка заполнения таблицы критериев",
                "Целевой документ": f"Критерий {i}, оценка пустая",
                "Различие": f"Для критерия {i} должна быть заполнена оценка: 0, 1 или 2"
            })

    return violations


@register("cheklist_eu", 6)
def check_total_score(target_doc: Dict[str, Any], config: Any) -> List[Dict[str, Any]]:
    """
    Правило #6: проверка итоговой оценки.

    Проверяет что указанная итоговая оценка равна сумме баллов по 7 критериям.
    Если оценки не заполнены — пропускаем (Rule 5 уже поймала).

    Args:
        target_doc: распарсенный документ {чанк: текст}
        config: конфиг аудита
    Returns:
        список нарушений (пустой = всё ок)
    """
    table_text = target_doc.get("таблица_критериев", "")

    if not table_text:
        return [{
            "rule_index": 6,
            "rule_title": "Проверка итоговой оценки",
            "Целевой документ": "отсутствует",
            "Различие": "Не найден чанк 'таблица_критериев'"
        }]

    criteria, stated_total = _parse_criteria(table_text)

    # Собираем заполненные оценки
    scores = {k: v for k, v in criteria.items() if v is not None}

    # Если не все 7 заполнены — не можем считать сумму, Rule 5 уже ругнулась
    if len(scores) != 7:
        return []

    calculated_sum = sum(scores.values())

    if stated_total is None:
        return [{
            "rule_index": 6,
            "rule_title": "Проверка итоговой оценки",
            "Целевой документ": "Строка «Итоговая оценка» не найдена или пустая",
            "Различие": f"Должна быть указана итоговая оценка. Рассчитанная сумма: {calculated_sum}"
        }]

    if stated_total != calculated_sum:
        return [{
            "rule_index": 6,
            "rule_title": "Проверка итоговой оценки",
            "Целевой документ": f"Итоговая оценка: {stated_total}",
            "Различие": f"Указанная сумма ({stated_total}) не совпадает с рассчитанной ({calculated_sum})"
        }]

    return []
