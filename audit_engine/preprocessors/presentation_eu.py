#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Препроцессор для Презентации ЭУ (presentation_eu).

Для compare="template" (правило #3 — проверка структуры):
сжимает каждый чанк до заголовка + краткой сводки.
LLM видит «слайд есть, тип контента такой-то», не получая полный текст.
"""

from .registry import register_preprocessor


def _compress_chunk(text: str, max_lines: int = 3) -> str:
    """
    Сжимает текст чанка до первых значимых строк.

    Логика:
    1. Убираем пустые строки в начале
    2. Берём первые max_lines непустых строк (заголовки/начало контента)
    3. Добавляем сводку о размере оригинала
    """
    if not text or not text.strip():
        return "[Слайд отсутствует или пуст]"

    lines = text.strip().split('\n')
    # Фильтруем пустые строки, берём первые значимые
    meaningful = [l.strip() for l in lines if l.strip()]

    if not meaningful:
        return "[Слайд отсутствует или пуст]"

    # Берём первые max_lines строк
    header_lines = meaningful[:max_lines]
    total_lines = len(meaningful)

    result = "[Слайд присутствует]\n"
    result += "\n".join(header_lines)

    if total_lines > max_lines:
        result += f"\n(... ещё {total_lines - max_lines} строк содержимого)"

    return result


# Регистрируем препроцессоры для всех чанков, используемых в правиле #3
# scope правила #3: титульный, план_мероприятий, инструменты_5с,
#                    результаты_проблемы, стандарты, последний

@register_preprocessor("presentation_eu", "титульный")
def compress_title(text: str) -> str:
    """Сжатие титульного слайда для структурной проверки."""
    return _compress_chunk(text)


@register_preprocessor("presentation_eu", "план_мероприятий")
def compress_plan(text: str) -> str:
    """Сжатие слайда плана мероприятий."""
    return _compress_chunk(text)


@register_preprocessor("presentation_eu", "инструменты_5с")
def compress_instruments(text: str) -> str:
    """Сжатие слайдов инструментов 5С."""
    return _compress_chunk(text, max_lines=5)


@register_preprocessor("presentation_eu", "результаты_проблемы")
def compress_results(text: str) -> str:
    """Сжатие слайдов результатов."""
    return _compress_chunk(text, max_lines=4)


@register_preprocessor("presentation_eu", "стандарты")
def compress_standards(text: str) -> str:
    """Сжатие слайдов стандартов."""
    return _compress_chunk(text)


@register_preprocessor("presentation_eu", "последний")
def compress_last(text: str) -> str:
    """Сжатие последнего слайда."""
    return _compress_chunk(text)
