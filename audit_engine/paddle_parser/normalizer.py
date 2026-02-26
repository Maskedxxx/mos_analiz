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


def merge_text_and_tables(
    full_page_text: str,
    table_entries: list,
    page_size: tuple,
) -> str:
    """
    Вставляет HTML-таблицы в текст страницы по y-позиции на странице.

    Таблицы были замаскированы белым при full-page OCR, поэтому их нет
    в full_page_text. Здесь мы вставляем их обратно в правильное место.

    Args:
        full_page_text: текст страницы от VLM (без таблиц).
        table_entries: [{"bbox": [x1,y1,x2,y2], "html": str}],
                       отсортированные по y (сверху вниз).
        page_size: (width, height) страницы в пикселях.

    Returns:
        str: текст с вставленными HTML-таблицами, нормализованный.
    """
    if not table_entries:
        return normalize_paddle_text(full_page_text)

    _, page_h = page_size
    lines = full_page_text.split('\n')
    total_lines = len(lines)

    if total_lines == 0:
        # Страница без текста — только таблицы
        parts = [t["html"] for t in table_entries]
        return "\n\n".join(parts)

    # Для каждой таблицы: позиция вставки = (y_center / page_height) * total_lines
    insertions = []
    for entry in table_entries:
        y1 = entry["bbox"][1]
        y2 = entry["bbox"][3]
        y_center = (y1 + y2) / 2
        relative_pos = y_center / page_h
        insert_at = int(relative_pos * total_lines)
        insert_at = max(0, min(insert_at, total_lines))
        insertions.append((insert_at, entry["html"]))

    # Вставляем с конца, чтобы не сбивать индексы
    insertions.sort(key=lambda x: x[0], reverse=True)
    for line_idx, html in insertions:
        lines.insert(line_idx, f"\n{html}\n")

    result = "\n".join(lines)
    return normalize_paddle_text(result)


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
