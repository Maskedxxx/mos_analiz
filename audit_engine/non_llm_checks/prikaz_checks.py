#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Детерминированные проверки для приказов (prikaz_comp_ppu, prikaz_ppu и др.).

Правило 4: наличие слова ПРИКАЗ и номера приказа.
Правило 5: наличие города и даты приказа.

Переведены с LLM на non-LLM из-за систематических ошибок:
- LLM путает scope (проверяет номер вместо даты)
- LLM отвергает garbled OCR номера как «некорректные»
- LLM не распознаёт словесный формат даты
"""

import re
from typing import Any, Dict, List

from .registry import register


def _get_scopes_text(target_doc: Dict[str, Any], scopes: List[str]) -> str:
    """Собирает текст из нескольких scope в одну строку."""
    parts = []
    for scope in scopes:
        val = target_doc.get(scope, "")
        if val:
            parts.append(str(val))
    return "\n".join(parts)


def check_prikaz_and_number(
    target_doc: Dict[str, Any],
    config: Any,
    rule_index: int = 4,
    rule_title: str = "Проверка наличия слова ПРИКАЗ и номера приказа."
) -> List[Dict[str, Any]]:
    """
    Проверяет наличие слова ПРИКАЗ и номера приказа.

    Допустимые формы: ПРИКАЗ, Приказ, П Р И К А З
    Номер: символ № + любой непустой текст (кроме подчёркиваний)
    """
    text = _get_scopes_text(target_doc, ["заголовок_город", "номер_дата"])
    violations = []

    # Проверка слова ПРИКАЗ
    has_prikaz = bool(re.search(
        r'П\s*Р\s*И\s*К\s*А\s*З|Приказ',
        text,
        re.IGNORECASE
    ))

    if not has_prikaz:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует слово «ПРИКАЗ»"
        })

    # Проверка номера: ищем № с непустым текстом после него
    # Находим все вхождения №
    number_match = re.search(r'№\s*(.+)', text)
    if number_match:
        # Текст после №
        after_sign = number_match.group(1).strip()
        # Убираем подчёркивания и пробелы — если осталось что-то, номер заполнен
        cleaned = re.sub(r'[_\s]', '', after_sign)
        if not cleaned:
            violations.append({
                "rule_index": rule_index,
                "rule_title": rule_title,
                "Целевой документ": f"№{after_sign}",
                "Различие": "Номер приказа не заполнен (пустой или подчёркивания)"
            })
    else:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует символ № с номером приказа"
        })

    return violations


def check_city_and_date(
    target_doc: Dict[str, Any],
    config: Any,
    rule_index: int = 5,
    rule_title: str = "Проверка наличия города и даты приказа."
) -> List[Dict[str, Any]]:
    """
    Проверяет наличие города и полной даты приказа.

    Город: «г. Москва», полный адрес с городом, «Москва»
    Дата: числовой (ДД.ММ.ГГГГ) или словесный (01 сентября 2025 г.)
    """
    text = _get_scopes_text(target_doc, ["заголовок_город", "номер_дата"])
    violations = []

    # Проверка города
    # Паттерны: "г. Москва", "г.Москва", "Москва г,", адрес с городом
    has_city = bool(re.search(
        r'г\.\s*[А-ЯЁ][а-яё]+|'           # г. Москва, г.Москва
        r'[А-ЯЁ][а-яё]+\s+г[.,]|'          # Москва г, / Москва г.
        r'город\s+[А-ЯЁ]',                  # город Москва
        text
    ))

    if not has_city:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует город подписания (например «г. Москва»)"
        })

    # Проверка даты
    # Месяцы для словесного формата
    months = (
        r'январ[яь]|феврал[яь]|март[а]?|апрел[яь]|ма[йя]|'
        r'июн[яь]|июл[яь]|август[а]?|сентябр[яь]|октябр[яь]|'
        r'ноябр[яь]|декабр[яь]'
    )

    has_date = bool(re.search(
        r'\d{2}[.\s]+\d{2}[.\s]+\d{4}|'                     # ДД.ММ.ГГГГ
        r'[«"\'"]?\d{1,2}[»"\'"]?\s*(?:' + months + r')',    # ДД месяц (словесный)
        text,
        re.IGNORECASE
    ))

    # Проверяем что дата не является плейсхолдером (__.__.202_)
    if has_date:
        placeholder = re.search(r'_+\._+\._+', text)
        if placeholder:
            has_date = False

    if not has_date:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует или неполная",
            "Различие": "Отсутствует полная дата приказа"
        })

    return violations


# === Регистрация для prikaz_comp_ppu ===

@register("prikaz_comp_ppu", 4)
def check_prikaz_number_comp_ppu(target_doc, config):
    """Правило #4: ПРИКАЗ + номер для Приказа о конкурсах ППУ."""
    return check_prikaz_and_number(target_doc, config, 4,
                                   "Проверка наличия слова ПРИКАЗ и номера приказа.")


@register("prikaz_comp_ppu", 5)
def check_city_date_comp_ppu(target_doc, config):
    """Правило #5: город + дата для Приказа о конкурсах ППУ."""
    return check_city_and_date(target_doc, config, 5,
                                "Проверка наличия города и даты приказа.")
