#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Non-LLM проверка обязательных элементов шапки.

Проверяет наличие ключевых слов в чанке «шапка»:
- «Приложение» — слово в заголовке
- «№» + цифры — номер приложения
- «приказ» / «приказу» — ссылка на приказ

Заменяет LLM-проверку, которая галлюцинирует на простых задачах поиска слов.
"""

import re
from typing import Any, Dict, List

from .registry import register


def check_shapka_elements_universal(
    target_doc: Dict[str, Any],
    config: Any,
    rule_index: int,
    rule_title: str
) -> List[Dict[str, Any]]:
    """
    Проверяет наличие обязательных элементов в шапке документа.

    Обязательные элементы:
    1. Слово «Приложение» (регистронезависимо)
    2. Номер приложения — «№» + цифры/буквы
    3. Ссылка на приказ — «приказ» или «приказу» (регистронезависимо)
    """
    shapka = target_doc.get("шапка", "")
    if not shapka:
        return [{
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Чанк «шапка» не найден в документе"
        }]

    shapka_lower = shapka.lower()
    missing = []

    # 1. Слово «Приложение»
    if "приложение" not in shapka_lower:
        missing.append("слово 'Приложение'")

    # 2. Номер приложения: «№» + хотя бы одна цифра/буква
    if not re.search(r'№\s*\S+', shapka):
        missing.append("номер приложения (№...)")

    # 3. Ссылка на приказ: «приказ» или «приказу»
    if not re.search(r'приказ[а-яё]*', shapka_lower):
        missing.append("ссылка на приказ (слово 'приказ'/'приказу')")

    if missing:
        return [{
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": shapka.strip()[:200],
            "Различие": f"Отсутствуют элементы: {', '.join(missing)}"
        }]

    return []


@register("polozhenie_comp_ppu", 6)
def check_shapka_polozhenie_comp_ppu(target_doc, config):
    """Правило #6: проверка обязательных элементов шапки для Положения о конкурсах ППУ."""
    return check_shapka_elements_universal(
        target_doc, config, 6,
        "Дополнительная сверка шапки с шаблоном."
    )
