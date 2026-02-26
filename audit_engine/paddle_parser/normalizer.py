#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Нормализация текста Paddle OCR pipeline.

Две функции:
- build_page_markdown() — собирает markdown из распознанных регионов одной страницы
- normalize_paddle_text() — нормализация для ChunkAssembler (без html_table_to_markdown)
"""

import logging
import re
from typing import List, Dict, Any

from . import config
from .table_parser import parse_paddle_table_html, is_table_format

logger = logging.getLogger(__name__)


def build_page_markdown(
    recognized_regions: List[Dict[str, Any]],
    page_num: int = 0,
) -> str:
    """
    Собирает markdown из распознанных регионов одной страницы.

    Фильтрует регионы без промптов и пустые. Таблицы конвертируются
    из формата PaddleOCR-VL (<fcel>/<nl>) в HTML. Каждый блок
    оборачивается в markdown по CLASS_MD_WRAPPER.

    Args:
        recognized_regions: список словарей:
            [{"class_name": str, "text": str, "bbox": list, "score": float}, ...]
            Уже отсортированный по reading order.
        page_num: номер страницы (для логирования).

    Returns:
        str: markdown-текст страницы.
    """
    parts = []

    for region in recognized_regions:
        cls = region["class_name"]
        text = region.get("text", "").strip()

        # Пропускаем классы без промптов
        if config.CLASS_PROMPTS.get(cls) is None:
            logger.debug(f"Пропуск: {cls} (нет промпта)")
            continue

        if not text:
            logger.debug(f"Пропуск: {cls} (пустой текст)")
            continue

        # Таблицы: парсим из формата PaddleOCR-VL в HTML
        if cls == "Table" and is_table_format(text):
            if config.TABLE_OUTPUT_FORMAT == "html":
                text = parse_paddle_table_html(text)
            else:
                from .table_parser import parse_paddle_table
                text = parse_paddle_table(text)

        # Оборачиваем в markdown
        wrapper = config.CLASS_MD_WRAPPER.get(cls, "{text}")
        md_block = wrapper.format(text=text)
        parts.append(md_block)

    result = "\n\n".join(parts)
    logger.info(f"Страница {page_num}: {len(parts)} блоков, {len(result)} символов markdown")
    return result


def normalize_paddle_text(text: str) -> str:
    """
    Нормализация OCR-текста от Paddle pipeline для ChunkAssembler.

    Отличия от normalize_ocr_text (HunyuanOCR):
    - НЕ вызывает unwrap_latex_artifacts() — PaddleOCR-VL не генерирует LaTeX-артефакты
    - НЕ вызывает html_table_to_markdown() — таблицы уже в HTML из build_page_markdown,
      и мы СОХРАНЯЕМ HTML с colspan/rowspan для LLM-аудита

    Операции:
    1. Множественные пустые строки -> одна
    2. Пробелы в конце строк

    Args:
        text: текст страницы (может содержать HTML-таблицы)

    Returns:
        str: нормализованный текст
    """
    if not text:
        return ""

    # Множественные пустые строки -> одна
    text = re.sub(r'\n{3,}', '\n\n', text)

    # Пробелы в конце строк
    text = re.sub(r' +\n', '\n', text)

    return text.strip()
