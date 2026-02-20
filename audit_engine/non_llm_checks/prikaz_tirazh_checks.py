#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Детерминированные проверки для prikaz_tirazh (0.5 Приказ о тиражировании).

Правило 2: наличие номера приказа (№ + цифры/буквы) и полной даты (ДД.ММ.ГГГГ).
Правило 7: наличие полной даты в пункте 11 «Отменить действие приказа от...».

Обе проверки работают на scope "текст_приказа" (OCR-текст страниц 1-2).
"""

import re
from typing import Any, Dict, List

from .registry import register


def _check_number_and_date(
    text: str,
    rule_index: int,
    rule_title: str
) -> List[Dict[str, Any]]:
    """
    Проверяет наличие номера приказа и полной даты в тексте.

    Args:
        text: OCR-текст приказа (страницы 1-2)
        rule_index: индекс правила
        rule_title: заголовок правила

    Returns:
        List[Dict] — список нарушений (пустой = норма)
    """
    violations = []

    # Шапка приказа — текст ДО «ПРИКАЗЫВАЮ» (номер и дата всегда в шапке)
    prikaz_split = re.split(r'ПРИКАЗЫВАЮ', text, maxsplit=1, flags=re.IGNORECASE)
    header_text = prikaz_split[0] if prikaz_split else text[:500]

    # --- Проверка номера: символ № с непустым текстом в шапке ---
    number_match = re.search(r'№\s*(.+)', header_text)
    if number_match:
        after_sign = number_match.group(1).strip()
        # Берём текст до конца строки или до "от"
        after_sign = re.split(r'\n|от\s', after_sign)[0].strip()
        # Убираем подчёркивания и пробелы — если осталось что-то, номер заполнен
        cleaned = re.sub(r'[_\s]', '', after_sign)
        if not cleaned:
            violations.append({
                "rule_index": rule_index,
                "rule_title": rule_title,
                "Целевой документ": f"№ {after_sign}",
                "Различие": "Номер приказа не заполнен (пустой или подчёркивания после №)"
            })
    else:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует",
            "Различие": "Отсутствует символ № с номером приказа"
        })

    # --- Проверка даты: ДД.ММ.ГГГГ (не плейсхолдер) в шапке ---
    # Паттерн полной даты: 2 цифры . 2 цифры . 4 цифры
    date_match = re.search(r'(\d{2})[.\s]+(\d{2})[.\s]+(\d{4})', header_text)

    # Также проверяем словесный формат: "01 сентября 2025 г."
    months = (
        r'январ[яь]|феврал[яь]|март[а]?|апрел[яь]|ма[йя]|'
        r'июн[яь]|июл[яь]|август[а]?|сентябр[яь]|октябр[яь]|'
        r'ноябр[яь]|декабр[яь]'
    )
    verbal_date = re.search(
        r'\d{1,2}\s*(?:' + months + r')\s*\d{4}',
        header_text, re.IGNORECASE
    )

    has_date = bool(date_match) or bool(verbal_date)

    # Проверяем что это не плейсхолдер: __.__.202_
    if has_date and re.search(r'_+[.\s]*_+[.\s]*_+', header_text):
        # Плейсхолдер рядом — проверяем что реальная дата есть ПОМИМО плейсхолдера
        if not date_match and not verbal_date:
            has_date = False

    if not has_date:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "отсутствует или неполная",
            "Различие": "Отсутствует полная дата приказа (ожидается ДД.ММ.ГГГГ)"
        })

    return violations


def _check_date_in_punkt_11(
    text: str,
    rule_index: int,
    rule_title: str
) -> List[Dict[str, Any]]:
    """
    Проверяет наличие полной даты в пункте 11 «Отменить действие приказа от...».

    Args:
        text: OCR-текст приказа (страницы 1-2)
        rule_index: индекс правила
        rule_title: заголовок правила

    Returns:
        List[Dict] — список нарушений (пустой = норма)
    """
    violations = []

    # Ищем пункт 11 с текстом "Отменить"
    # Варианты: "11. Отменить", "11.Отменить", "11 Отменить"
    punkt_match = re.search(
        r'11\s*[.)\s]\s*(Отменить[^\n]*(?:\n(?!\d+\s*[.)])[^\n]*)*)',
        text, re.IGNORECASE
    )

    if not punkt_match:
        # Альтернатива: ищем просто "Отменить действие приказа"
        punkt_match = re.search(
            r'(Отменить\s+действие\s+приказа[^\n]*)',
            text, re.IGNORECASE
        )

    if not punkt_match:
        # Пункт не найден — не можем проверить
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": "пункт 11 не найден",
            "Различие": "Не удалось найти пункт 11 «Отменить действие приказа» в тексте"
        })
        return violations

    punkt_text = punkt_match.group(1)

    # Ищем дату после "от" в пункте
    date_after_ot = re.search(r'от\s+(\d{2})[.\s]+(\d{2})[.\s]+(\d{4})', punkt_text)

    # Также проверяем просто дату ДД.ММ.ГГГГ в тексте пункта
    date_anywhere = re.search(r'(\d{2})[.\s]+(\d{2})[.\s]+(\d{4})', punkt_text)

    has_date = bool(date_after_ot) or bool(date_anywhere)

    # Проверяем плейсхолдер: __.___202_, «от _» и т.д.
    has_placeholder = bool(re.search(r'_+[.\s]*_+[.\s]*_+|_+\s*202_', punkt_text))

    if has_placeholder and not date_after_ot:
        has_date = False

    if not has_date:
        violations.append({
            "rule_index": rule_index,
            "rule_title": rule_title,
            "Целевой документ": punkt_text.strip()[:150],
            "Различие": "Дата отменяемого приказа не заполнена или неполная (ожидается ДД.ММ.ГГГГ)"
        })

    return violations


# === Регистрация проверок для prikaz_tirazh ===

@register("prikaz_tirazh", 2)
def check_number_date_tirazh(target_doc: Dict[str, Any], config: Any) -> List[Dict[str, Any]]:
    """Правило #2: номер приказа + дата для Приказа о тиражировании."""
    text = target_doc.get("текст_приказа", "")
    return _check_number_and_date(
        text, 2, "Проверка номера приказа и даты."
    )


@register("prikaz_tirazh", 7)
def check_date_punkt_11_tirazh(target_doc: Dict[str, Any], config: Any) -> List[Dict[str, Any]]:
    """Правило #7: дата отменяемого приказа в п.11 для Приказа о тиражировании."""
    text = target_doc.get("текст_приказа", "")
    return _check_date_in_punkt_11(
        text, 7, "Проверка даты отменяемого приказа в пункте 11."
    )
