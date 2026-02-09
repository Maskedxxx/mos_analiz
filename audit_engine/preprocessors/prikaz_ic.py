#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Препроцессор для Приказа о создании ИЦ (prikaz_ic).

normalize_text_for_rule3 — удаляет плейсхолдеры из текста
для сравнения ТЕКСТОВОЙ СТРУКТУРЫ (правило #3).
Заполненность плейсхолдеров проверяет правило #4.

ВАЖНО: использует [ \\t]* вместо \\s* — чтобы не съедать \\n и не ломать п.5.
"""

import re

from .registry import register_preprocessor


@register_preprocessor("prikaz_ic", "текст_приказа")
@register_preprocessor("prikaz_ic_el", "текст_приказа")
def normalize_text_for_rule3(text: str) -> str:
    """
    Нормализует текст приказа для сравнения с шаблоном.

    Удаляет вариативные части (даты, ФИО, подписант),
    оставляя только текстовую структуру.
    """
    # 1. П.3: удаляем дату после "с"
    # "показателям с 27.08.2025" → "показателям"
    text = re.sub(
        r'(приступить к заполнению разделов по своим показателям)[ \t]+с[ \t]*[^\n]*',
        r'\1',
        text
    )
    text = re.sub(
        r'(приступить к заполнению разделов по своим показателям)[ \t]+с[ \t]*\[ДАТА\]',
        r'\1',
        text
    )

    # 2. П.4: унифицируем должность+ФИО до "организовать"
    # "4. Директору Иванову И.И. организовать" → "4. [ДОЛЖНОСТЬ] организовать"
    text = re.sub(
        r'4\.\s*.+?\s+организовать',
        r'4. [ДОЛЖНОСТЬ] организовать',
        text
    )

    # 3. Удаляем дату после "в срок до"
    # ВАЖНО: [ \t]* а не \s* — сохраняем переносы строк
    text = re.sub(
        r'(в срок до)[ \t]*[^\n]+',
        r'\1',
        text
    )
    text = re.sub(
        r'(в срок до)[ \t]*\[ДАТА\]',
        r'\1',
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

        # Строка с признаками подписанта (должность + много пробелов + ФИО)
        if (re.search(r'\s{5,}', line_stripped) and
            ('директор' in line_stripped.lower() or
             'фамилия' in line_stripped.lower() or
             re.search(r'[А-Я]\.[А-Я]\.', line_stripped))):
            continue

        new_lines.append(line)

    text = '\n'.join(new_lines)
    return text
