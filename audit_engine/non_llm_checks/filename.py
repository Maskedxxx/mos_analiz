#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Универсальная проверка имени файла (non-LLM).

Сравнивает имя целевого документа с паттерном из config.json.
Паттерн задаётся в поле filename_pattern конфига.
"""

import re
from typing import Any, Dict, List

from .registry import register


def check_filename_universal(
    target_doc: Dict[str, Any],
    config: Any,
    rule_index: int = 1,
    rule_title: str = "Проверка названия файла"
) -> List[Dict[str, Any]]:
    """
    Универсальная проверка имени файла.

    Два режима:
    1. filename_keywords (приоритет) — все ключевые слова должны быть в имени файла
    2. filename_pattern (обратная совместимость) — подстрока должна быть в имени
    """
    actual = target_doc.get("имя_файла", "")
    # Убираем расширение
    actual_clean = re.sub(r'\.(docx?|pptx?|pdf)$', '', actual, flags=re.IGNORECASE)
    name_lower = actual_clean.lower()

    # Режим 1: список ключевых слов (приоритет)
    keywords = getattr(config, 'filename_keywords', None)
    if keywords:
        if all(kw.lower() in name_lower for kw in keywords):
            return []
        # Определяем какие слова отсутствуют
        missing = [kw for kw in keywords if kw.lower() not in name_lower]
        return [{
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": actual_clean,
            "Различие": f"Ожидалось имя файла, содержащее все ключевые слова: {keywords}. "
                        f"Отсутствуют: {missing}"
        }]

    # Режим 2: подстрока (обратная совместимость)
    expected_pattern = getattr(config, 'filename_pattern', "")
    if not expected_pattern:
        return []

    if expected_pattern.lower() in name_lower:
        return []

    return [{
        "rule_index": rule_index,
        "rule_title": rule_title,
        "Целевой документ": actual_clean,
        "Различие": f"Ожидалось имя файла, содержащее: «{expected_pattern}»"
    }]


# Регистрируем для всех известных типов документов
# При миграции итераций нужно добавить регистрацию для каждого doc_type
@register("presentation_eu", 1)
def check_filename_presentation(target_doc, config):
    """Правило #1: проверка имени файла для Презентации ЭУ."""
    return check_filename_universal(target_doc, config, 1, "Проверка названия файла")


@register("cheklist_eu", 1)
def check_filename_cheklist(target_doc, config):
    """Правило #1: проверка имени файла для Чек-листа ЭУ."""
    return check_filename_universal(target_doc, config, 1, "Проверка названия файла")


@register("prikaz_ic", 1)
def check_filename_prikaz(target_doc, config):
    """Правило #1: проверка имени файла для Приказа о создании ИЦ."""
    return check_filename_universal(target_doc, config, 1, "Проверка названия файла")


@register("polozhenie_po", 1)
def check_filename_polozhenie_po(target_doc, config):
    """Правило #1: проверка имени файла для Положения о ПО."""
    return check_filename_universal(target_doc, config, 1, "Проверка названия файла")


@register("prikaz_formirovanie_po", 1)
def check_filename_prikaz_formirovanie_po(target_doc, config):
    """Правило #1: проверка имени файла для Приказа о формировании ПО."""
    return check_filename_universal(target_doc, config, 1, "Проверка названия файла")


@register("prikaz_ic_el", 1)
def check_filename_prikaz_ic_el(target_doc, config):
    """Правило #1: проверка имени файла для Приказа о создании ИЦ (эл. вид)."""
    return check_filename_universal(target_doc, config, 1, "Проверка имени файла документа.")


@register("prikaz_ic_potoka", 2)
def check_filename_prikaz_ic_potoka(target_doc, config):
    """Правило #2: проверка имени файла для Приказа о создании ИЦ потока."""
    return check_filename_universal(target_doc, config, 2, "Проверка имени файла документа.")


@register("prikaz_ic_potoka_el", 2)
def check_filename_prikaz_ic_potoka_el(target_doc, config):
    """Правило #2: проверка имени файла для Приказа о создании ИЦ потока (эл. вид)."""
    return check_filename_universal(target_doc, config, 2, "Проверка имени файла документа.")


@register("prikaz_vyhod", 2)
def check_filename_prikaz_vyhod(target_doc, config):
    """Правило #2: проверка имени файла для Приказа о проведении выхода."""
    return check_filename_universal(target_doc, config, 2, "Проверка названия файла документа")


@register("prikaz_comp_ppu", 2)
def check_filename_prikaz_comp_ppu(target_doc, config):
    """Правило #2: проверка имени файла для Приказа о конкурсах ППУ."""
    return check_filename_universal(target_doc, config, 2, "Проверка имени файла документа.")


@register("prikaz_ppu", 2)
def check_filename_prikaz_ppu(target_doc, config):
    """Правило #2: проверка имени файла для Приказа о ППУ."""
    return check_filename_universal(target_doc, config, 2, "Проверка имени файла документа.")


@register("polozhenie_comp_ppu", 2)
def check_filename_polozhenie_comp_ppu(target_doc, config):
    """Правило #2: проверка имени файла для Положения о конкурсах ППУ."""
    return check_filename_universal(target_doc, config, 2, "Проверка имени файла документа.")


@register("polozhenie_ppu", 2)
def check_filename_polozhenie_ppu(target_doc, config):
    """Правило #2: проверка имени файла для Положения о ППУ."""
    return check_filename_universal(target_doc, config, 2, "Проверка имени файла документа.")


@register("akt_nachala", 1)
def check_filename_akt_nachala(target_doc, config):
    """Non-LLM: проверка имени файла для Акта начала мероприятий"""
    return check_filename_universal(target_doc, config, 1, "Проверка названия файла")


@register("prikaz_pa", 1)
def check_filename_prikaz_pa(target_doc, config):
    """Правило #1: проверка имени файла для Приказа о внедрении ПА."""
    return check_filename_universal(target_doc, config, 1, "Проверка имени файла документа.")


@register("prikaz_tirazh", 1)
def check_filename_prikaz_tirazh(target_doc, config):
    """Правило #1: проверка имени файла для Приказа о тиражировании."""
    return check_filename_universal(target_doc, config, 1, "Проверка имени файла документа.")
