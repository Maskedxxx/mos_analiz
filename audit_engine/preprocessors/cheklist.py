#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Препроцессоры для Чек-листа выбора ЭУ (cheklist_eu).

- normalize_text_for_comparison — маскирование плейсхолдеров (даты, ФИО, организации)
- normalize_table_for_comparison — нормализация таблицы критериев
"""

import re

from .registry import register_preprocessor


def normalize_text_for_comparison(text: str) -> str:
    """
    Нормализует текст перед сравнением, заменяя плейсхолдеры на унифицированные метки.
    Применяется и к целевому документу и к шаблону.
    """
    # 0. Убираем аннотации менеджера (текст типа " - заголовок", " - дата приказа")
    text = re.sub(r'\s+-\s+(заголовок|дата|номер|форма|место|подпись).*$', '', text, flags=re.MULTILINE | re.IGNORECASE)

    # 1. Нормализуем пробелы
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n\s*\n', '\n\n', text)

    # 2. Нормализуем даты
    text = re.sub(r'\d{1,2}\.\d{1,2}\.\d{4}\s*г?\.?', '[ДАТА]', text)
    text = re.sub(r'«?\d{1,2}»?\s*[а-яё]+\s*\d{4}\s*г?\.?', '[ДАТА]', text, flags=re.IGNORECASE)
    text = re.sub(r'_{2,}\.\s*_{2,}\.\s*\d{4}|_{2,}\.\s*_{2,}\.\s*202_?', '[ДАТА]', text)

    # 3. Нормализуем ФИО
    text = re.sub(r'[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.[А-ЯЁ]\.', '[ФИО]', text)
    text = re.sub(r'[А-ЯЁ]\.[А-ЯЁ]\.\s*[А-ЯЁ][а-яё]+', '[ФИО]', text)
    text = re.sub(r'И\.О\.\s*Фамилия', '[ФИО]', text)
    text = re.sub(r'\(Фамилия И\.О\.\)', '', text)

    # 4. Нормализуем организации
    text = re.sub(r'(ООО|ЗАО|АО|ПАО)\s*[«"][\w\s]+[»"]', '[ОРГАНИЗАЦИЯ]', text)
    text = re.sub(r'(ООО|ЗАО|АО|ПАО)\s*[«"]_+[»"]', '[ОРГАНИЗАЦИЯ]', text)

    # 5. Нормализуем номера приказов
    text = re.sub(r'№\s*\d+[а-яА-Я]*', '№ [НОМЕР]', text)
    text = re.sub(r'№\s*_+', '№ [НОМЕР]', text)

    # 6. Убираем оставшиеся длинные подчёркивания
    text = re.sub(r'_+', '[ПЛЕЙСХОЛДЕР]', text)

    # 7. Убираем дублирующиеся токены
    text = re.sub(r'\[ДАТА\]\s*\[ДАТА\]', '[ДАТА]', text)
    text = re.sub(r'\[ФИО\]\s*\[ФИО\]', '[ФИО]', text)

    return text.strip()


def normalize_table_for_comparison(table_text: str) -> str:
    """
    Нормализует таблицу критериев для сравнения TARGET и TEMPLATE.

    1. Заменяет числовые оценки (0, 1, 2) на [ОЦЕНКА]
    2. Добавляет номера критериев если отсутствуют
    3. Убирает структурные различия
    """
    if not table_text:
        return table_text

    # Карта критериев по ключевым словам → номер
    criteria_keywords = {
        'Влияние результатов работы эталонного участка': '1',
        'Руководство предприятия выделяет этот участок': '2',
        'На участке выявлены резервы повышения производительности': '3',
        'Применение обязательных инструментов БП': '4',
        'На участке есть проблемы, которые возможно исключить': '5',
        'На какие потоки предприятия влияет': '6',
        'Оцените потенциал тиражирования': '7',
    }

    lines = table_text.split('\n')
    normalized_lines = []

    for line in lines:
        # Пропускаем служебные строки
        if line.strip().startswith('[') or 'Итоговая оценка' in line:
            normalized_lines.append(line)
            continue

        # Для строк таблицы с "|"
        if '|' in line:
            parts = line.split('|')

            # Строки с 5+ ячейками (№, критерий, оценка, описание_0, описание_1, ...)
            # Оставляем только первые 3 значимых столбца: №, критерий, оценка
            # parts[0] = '' (до первого |), parts[1] = №, parts[2] = критерий, parts[3] = оценка
            if len(parts) > 5:
                line = '|'.join(parts[:4]) + '|'

            # Добавляем номер критерия если отсутствует
            for keyword, num in criteria_keywords.items():
                if keyword in line:
                    pattern = r'^\|\s*\|\s*(' + re.escape(keyword[:20]) + ')'
                    if re.search(pattern, line):
                        line = re.sub(pattern, r'| ' + num + r' | \1', line)
                    break

            # Заменяем числовые оценки на [ОЦЕНКА]
            line = re.sub(r'\|\s*([012])\s*\|', r'| [ОЦЕНКА] |', line)

            # Разделительные строки (|---|---|...) — тоже обрезаем
            if re.match(r'^\|[\s\-|]+$', line):
                dashes = line.split('|')
                if len(dashes) > 5:
                    line = '|'.join(dashes[:4]) + '|'

        normalized_lines.append(line)

    return '\n'.join(normalized_lines)


# Регистрируем препроцессоры для чек-листа
# Для compare="template" — нормализация текста и таблиц

@register_preprocessor("cheklist_eu", "шапка")
def preprocess_cheklist_shapa(text: str) -> str:
    """Нормализация шапки чек-листа."""
    return normalize_text_for_comparison(text)


@register_preprocessor("cheklist_eu", "таблица_критериев")
def preprocess_cheklist_table(text: str) -> str:
    """Нормализация таблицы критериев."""
    text = normalize_text_for_comparison(text)
    text = normalize_table_for_comparison(text)
    return text


@register_preprocessor("cheklist_eu", "подвал")
def preprocess_cheklist_podval(text: str) -> str:
    """Нормализация подвала чек-листа."""
    return normalize_text_for_comparison(text)
