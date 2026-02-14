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

    # 5. Нормализация пробелов в начале строк (OCR-артефакты)
    # "  2. Регламент..." → "2. Регламент..."
    lines = text.split('\n')
    text = '\n'.join(line.lstrip() for line in lines)

    return text


@register_preprocessor("prikaz_ic", "шапка")
@register_preprocessor("prikaz_ic_el", "шапка")
def normalize_header(text: str) -> str:
    """
    Нормализует шапку приказа — заменяет юридический адрес на краткий формат города.

    OCR иногда извлекает полный юрадрес: "109316, Москва г, Внутригородская..."
    LLM цепляется за формат адреса вместо проверки наличия города.
    Нормализуем: извлекаем город, заменяем всю строку адреса на "г. Город".
    """
    # Известные города
    known_cities = ['Москва', 'Санкт-Петербург', 'Новосибирск', 'Екатеринбург',
                    'Казань', 'Нижний Новгород', 'Челябинск', 'Самара', 'Омск',
                    'Ростов-на-Дону', 'Уфа', 'Красноярск', 'Пермь', 'Воронеж',
                    'Волгоград', 'Краснодар', 'Тюмень', 'Тольятти', 'Барнаул']

    lines = text.split('\n')
    new_lines = []
    for line in lines:
        stripped = line.strip()
        # Ищем строку с полным юридическим адресом (индекс + город + улица)
        if re.match(r'^\d{5,6}\s*,', stripped):
            # Извлекаем город из адреса
            for city in known_cities:
                if city.lower() in stripped.lower():
                    new_lines.append(f'г. {city}')
                    break
            else:
                new_lines.append(line)
        # "Москва г" или "г Москва" без контекста адреса — нормализуем
        elif re.match(r'^(г\.?\s+)?(' + '|'.join(known_cities) + r')\s+г\.?$', stripped):
            city_match = re.search(r'(' + '|'.join(known_cities) + r')', stripped)
            if city_match:
                new_lines.append(f'г. {city_match.group(1)}')
            else:
                new_lines.append(line)
        else:
            new_lines.append(line)
    return '\n'.join(new_lines)


@register_preprocessor("prikaz_ic", "приложение_2_к_приказу")
@register_preprocessor("prikaz_ic_el", "приложение_2_к_приказу")
def trim_appendix2_to_relevant_sections(text: str) -> str:
    """
    Сокращает чанк приложение_2_к_приказу — убирает основное тело Регламента,
    оставляя заголовок (номер приказа, дата) и секцию «Приложение №1 к Регламенту».

    Без этого LLM путает нумерацию основного тела (1.1, 2.1, 4.1) с нумерацией
    Приложения к Регламенту (1.1.1, 2.1, 4.1), где находятся проверяемые поля.
    """
    lines = text.split('\n')
    header_lines = []
    appendix_lines = []
    in_appendix = False
    header_collected = False

    for line in lines:
        stripped = line.strip()
        # Собираем заголовок (первые строки до РЕГЛАМЕНТ или 1.)
        if not header_collected:
            if re.match(r'^(РЕГЛАМЕНТ|1\.\s)', stripped):
                header_collected = True
            else:
                header_lines.append(line)
                continue

        # Ищем начало Приложения к Регламенту
        if not in_appendix:
            if re.match(r'^Приложение\s*№?\s*1\s*(к\s+Регламенту|к\s+регламенту)', stripped, re.IGNORECASE):
                in_appendix = True
                appendix_lines.append(line)
        else:
            appendix_lines.append(line)

    # Если нашли Приложение — возвращаем заголовок + Приложение
    if appendix_lines:
        return '\n'.join(header_lines + [''] + appendix_lines)

    # Если Приложение не найдено — возвращаем как есть
    return text
