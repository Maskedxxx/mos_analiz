#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Non-LLM проверки для Чек-листа выбора ЭУ (cheklist_eu).

Правило #5: проверка заполнения оценок критериев
Правило #6: проверка итоговой оценки (сумма баллов)
"""

import re
from typing import Any, Dict, List

from .registry import register


@register("cheklist_eu", 5)
def check_criteria_scores(target_doc: Dict[str, Any], config: Any) -> List[Dict[str, Any]]:
    """
    Правило #5: проверка заполнения оценок критериев.

    Проверяет что все 7 критериев имеют оценку (0, 1 или 2).
    Парсит markdown-таблицу из чанка 'таблица_критериев'.
    """
    table_text = target_doc.get("таблица_критериев", "")

    if not table_text:
        return [{
            "rule_index": 5,
            "rule_title": "Проверка заполнения таблицы критериев",
            "Целевой документ": "отсутствует",
            "Различие": "Не найден чанк 'таблица_критериев'"
        }]

    violations = []
    found_criteria = {}  # {номер_критерия: оценка или None}

    # Паттерн: | номер(1-7) | текст | оценка(0/1/2) |
    pattern = r'^\|\s*([1-7])\s*\|[^|]+\|\s*([012]?)\s*\|'

    for line in table_text.split('\n'):
        match = re.match(pattern, line.strip())
        if match:
            criterion_num = int(match.group(1))
            score_str = match.group(2).strip()

            if score_str in ('0', '1', '2'):
                found_criteria[criterion_num] = int(score_str)
            else:
                found_criteria[criterion_num] = None

    # Проверяем все 7 критериев
    for i in range(1, 8):
        if i not in found_criteria:
            violations.append({
                "rule_index": 5,
                "rule_title": "Проверка заполнения таблицы критериев",
                "Целевой документ": f"строка {i}, колонка 'Оценка (значение )' отсутствует",
                "Различие": f"Для критерия {i} должна быть цифра 0, 1 или 2 в колонке 'Оценка (значение )'"
            })
        elif found_criteria[i] is None:
            violations.append({
                "rule_index": 5,
                "rule_title": "Проверка заполнения таблицы критериев",
                "Целевой документ": f"строка {i}, оценка пустая",
                "Различие": f"Для критерия {i} должна быть цифра 0, 1 или 2 в колонке 'Оценка (значение )'"
            })

    return violations


@register("cheklist_eu", 6)
def check_total_score(target_doc: Dict[str, Any], config: Any) -> List[Dict[str, Any]]:
    """
    Правило #6: проверка итоговой оценки.

    Проверяет что указанная итоговая оценка равна сумме баллов по 7 критериям.
    """
    table_text = target_doc.get("таблица_критериев", "")

    if not table_text:
        return [{
            "rule_index": 6,
            "rule_title": "Проверка итоговой оценки",
            "Целевой документ": "отсутствует",
            "Различие": "Не найден чанк 'таблица_критериев'"
        }]

    # Извлекаем оценки критериев 1-7
    scores = {}
    pattern = r'^\|\s*([1-7])\s*\|[^|]+\|\s*([012])\s*\|'
    for line in table_text.split('\n'):
        match = re.match(pattern, line.strip())
        if match:
            criterion_num = int(match.group(1))
            score = int(match.group(2))
            scores[criterion_num] = score

    # Проверяем что все 7 критериев найдены
    if len(scores) != 7:
        missing = [i for i in range(1, 8) if i not in scores]
        return [{
            "rule_index": 6,
            "rule_title": "Проверка итоговой оценки",
            "Целевой документ": f"найдено {len(scores)} критериев из 7",
            "Различие": f"Не удалось извлечь оценки для критериев: {missing}"
        }]

    # Считаем сумму
    calculated_sum = sum(scores.values())

    # Ищем строку "Итоговая оценка"
    total_pattern = r'\|\s*\|\s*Итоговая оценка\s*\|\s*(\d+)\s*\|'
    total_match = re.search(total_pattern, table_text, re.IGNORECASE)

    if not total_match:
        return [{
            "rule_index": 6,
            "rule_title": "Проверка итоговой оценки",
            "Целевой документ": "строка 'Итоговая оценка' не найдена",
            "Различие": f"Должна быть строка с итоговой оценкой. Рассчитанная сумма: {calculated_sum}"
        }]

    stated_total = int(total_match.group(1))

    if stated_total != calculated_sum:
        return [{
            "rule_index": 6,
            "rule_title": "Проверка итоговой оценки",
            "Целевой документ": f"Итоговая оценка: {stated_total}",
            "Различие": f"Указанная сумма ({stated_total}) не совпадает с рассчитанной ({calculated_sum})"
        }]

    return []
