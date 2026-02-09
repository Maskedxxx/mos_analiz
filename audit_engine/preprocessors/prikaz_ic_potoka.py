#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Препроцессор для Приказа о создании ИЦ потока (итерация 7).

normalize_text_for_rule3 — заменяет переменные данные на унифицированные токены:
- Даты → [ДАТА]
- Должность+ФИО перед "организовать" → [ДОЛЖНОСТЬ_ФИО]
- Строка подписанта → [ПОДПИСАНТ]

Применяется к scope "текст_приказа" при template-сравнении (правило #3).
Используется для обоих вариантов: бумажного и электронного.
"""

import re

from .registry import register_preprocessor


def normalize_text_for_rule3(text: str) -> str:
    """
    Нормализует текст приказа для сравнения с шаблоном.

    Заменяет переменные части на токены, оставляя структуру для сравнения.
    """
    # 0. Нормализуем пробелы
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*', '\n', text)

    # 1. Даты: конкретные и плейсхолдеры → [ДАТА]
    text = re.sub(r'\d{1,2}\.\d{1,2}\.\d{4}', '[ДАТА]', text)
    text = re.sub(r'_{2,}\.\s*_{2,}\.\s*202_?', '[ДАТА]', text)

    # 2. Должность+ФИО перед "организовать" → [ДОЛЖНОСТЬ_ФИО]
    # Шаблон: "(Указать наименование должности) организовать..."
    text = re.sub(
        r'\(Указать наименование должности\)',
        '[ДОЛЖНОСТЬ_ФИО]',
        text
    )
    # Целевой: "Начальнику службы ... Ивановой А.А. организовать..."
    text = re.sub(
        r'[А-ЯЁа-яё\s]+[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]\.\s+организовать',
        '[ДОЛЖНОСТЬ_ФИО] организовать',
        text
    )

    # 3. Разделяем слипшиеся [ДАТА][ДОЛЖНОСТЬ_ФИО]
    text = re.sub(
        r'\[ДАТА\]\s*\[ДОЛЖНОСТЬ_ФИО\]',
        r'[ДАТА]\n[ДОЛЖНОСТЬ_ФИО]',
        text
    )

    # 4. Подписант → [ПОДПИСАНТ]
    lines = text.split('\n')
    for i in range(len(lines) - 1, -1, -1):
        line = lines[i].strip()
        if line and ('директор' in line.lower() or 'должность' in line.lower() or
                     'фамилия' in line.lower() or 'фио' in line.lower() or
                     re.search(r'[А-ЯЁ]\.[А-ЯЁ]\.', line)):
            lines[i] = '[ПОДПИСАНТ]'
            break
    text = '\n'.join(lines)

    return text


# Регистрация для обоих вариантов — бумажный и электронный
@register_preprocessor("prikaz_ic_potoka", "текст_приказа")
def preprocess_potoka_tekst(text: str) -> str:
    """Нормализация текста приказа ИЦ потока (бумажный)."""
    return normalize_text_for_rule3(text)


@register_preprocessor("prikaz_ic_potoka_el", "текст_приказа")
def preprocess_potoka_el_tekst(text: str) -> str:
    """Нормализация текста приказа ИЦ потока (электронный)."""
    return normalize_text_for_rule3(text)
