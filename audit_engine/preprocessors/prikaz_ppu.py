#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Препроцессор для Приказа о ППУ (prikaz_ppu).

normalize_title — нормализует заголовок:
  удаляет номер приказа и город (проверяются отдельными правилами),
  оставляя только текст заголовка для сравнения с шаблоном.

normalize_text_for_rule3 — заменяет вариативные плейсхолдеры
  (должность + ФИО, даты) на унифицированные метки.

ВАЖНО: использует [ \\t]* вместо \\s* — чтобы не съедать \\n.
"""

import re

from .registry import register_preprocessor


@register_preprocessor("prikaz_ppu", "заголовок_город")
def normalize_title(text: str) -> str:
    """
    Нормализует чанк заголовок_город для сравнения с шаблоном (правило #1).

    Vision извлекает элементы в разном порядке. Убираем вариативные части,
    оставляя ТОЛЬКО текст заголовка приказа.
    """
    lines = text.strip().split('\n')
    result = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        # Пропускаем строку с «ПРИКАЗ №...» — проверяется правилом #4
        if re.match(r'^П\s*Р\s*И\s*К\s*А\s*З', stripped) or stripped.startswith('ПРИКАЗ'):
            continue
        # Пропускаем строку с городом «г. ...» — проверяется правилом #5
        if re.match(r'^г\.\s*', stripped):
            continue
        # Пропускаем строку с номером приказа (голый «№ ...»)
        if re.match(r'^№\s*', stripped):
            continue
        # Пропускаем строку с датой (ДД.ММ.ГГГГ или плейсхолдер)
        if re.match(r'^[\d"_<]', stripped) and not re.match(r'^\d+\.', stripped):
            continue
        result.append(stripped)
    return '\n'.join(result)


@register_preprocessor("prikaz_ppu", "текст_приказа")
def normalize_text_for_rule3(text: str) -> str:
    """
    Нормализует текст приказа о ППУ для сравнения с шаблоном (правило #3).

    Удаляет вариативные части (даты, ФИО, подписант),
    оставляя только текстовую структуру.
    """
    # 1. «возложить на <ФИО + должность>.» → «возложить на [ДОЛЖНОСТЬ_ФИО].»
    #    ВАЖНО: жадный .+ (не .+?) чтобы матчить до ПОСЛЕДНЕЙ точки в строке,
    #    иначе ленивый квантор останавливается на первой точке инициалов (А.)
    text = re.sub(
        r'(возложить на\s*).+\.',
        r'\1[ДОЛЖНОСТЬ_ФИО].',
        text
    )

    # 2. Удаляем дату после "в срок до"
    # ВАЖНО: [ \t]* а не \s* — сохраняем переносы строк
    text = re.sub(
        r'(в срок до)[ \t]*[^\n]+',
        r'\1',
        text
    )

    # 3. Удаляем конкретные даты в формате ДД.ММ.ГГГГ
    text = re.sub(
        r'\d{2}\.\d{2}\.\d{4}(?:[ \t]*г\.?)?',
        '[ДАТА]',
        text
    )

    # 4. Удаляем строку подписанта целиком (она вариативна)
    lines = text.split('\n')
    new_lines = []
    for line in lines:
        line_stripped = line.strip()

        # Пропускаем пустые
        if not line_stripped:
            new_lines.append(line)
            continue

        # Плейсхолдеры подписанта из шаблона
        if line_stripped in ('[ПОДПИСАНТ]', 'Генеральный директор', 'И.О. Фамилия'):
            continue

        # ФИО вида «И.О. Фамилия» — строка подписанта
        if re.match(r'^[А-ЯЁ]\.[А-ЯЁ]\.\s+[А-ЯЁ][а-яё]+$', line_stripped):
            continue

        # Строка с признаками подписанта (должность + много пробелов + ФИО)
        if (re.search(r'\s{5,}', line_stripped) and
            ('директор' in line_stripped.lower() or
             'фамилия' in line_stripped.lower() or
             re.search(r'[А-Я]\.[А-Я]\.', line_stripped))):
            continue

        # Плейсхолдеры шаблона
        if line_stripped.startswith('<') and line_stripped.endswith('>'):
            continue

        new_lines.append(line)

    text = '\n'.join(new_lines)
    return text
