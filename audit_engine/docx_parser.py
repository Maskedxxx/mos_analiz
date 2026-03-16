#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DOCX-парсер: извлечение текста и таблиц из DOCX через python-docx.

Альтернатива Paddle OCR для формата DOCX — читает текст и таблицы
напрямую из XML, без конвертации в изображения.

Возвращает Dict[str, Any] в формате legacy-парсера (как PaddleParser/VisionParser).

Ограничение: DOCX не имеет фиксированных страниц (пагинация зависит от рендерера).
Поле "pages" в chunks_vision.json интерпретируется как приблизительные доли документа:
  - pages: [1] → первая часть документа
  - pages: [-1] → последняя часть документа
  - pages: "all" или полный список → весь документ
"""

import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx import Document

logger = logging.getLogger(__name__)


def parse_docx(
    file_path: str,
    chunks_vision_path: str,
    chunk_filter: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Парсинг DOCX-файла через python-docx.

    Args:
        file_path: путь к DOCX-файлу
        chunks_vision_path: путь к chunks_vision.json (берём name + pages)
        chunk_filter: имя чанка (если указан — собираем только его)

    Returns:
        Dict[str, Any] — {"имя_файла": ..., "путь": ..., "чанк1": текст, ...}
    """
    file_path = Path(file_path)

    # --- Загрузка конфигурации чанков ---
    with open(chunks_vision_path, "r", encoding="utf-8") as f:
        chunks_config = json.load(f)
    chunks = chunks_config.get("chunks", [])

    # --- Открытие документа ---
    doc = Document(str(file_path))
    logger.info(f"DOCX: {file_path.name}, абзацев: {len(doc.paragraphs)}, таблиц: {len(doc.tables)}")

    # --- Извлечение всех элементов в порядке документа ---
    elements = _extract_elements_in_order(doc)
    total_elements = len(elements)
    logger.info(f"  Извлечено элементов: {total_elements}")

    # --- Определяем общее число "страниц" для маппинга ---
    # Берём максимальный номер страницы из всех чанков
    max_page = 1
    for chunk in chunks:
        pages = chunk.get("pages", [])
        if pages == "all":
            continue
        for p in pages:
            abs_p = abs(p)
            if abs_p > max_page:
                max_page = abs_p
    total_pages = max(max_page, 1)

    # --- Разбиваем элементы на "страницы" (равные доли) ---
    pages_content = _split_into_pages(elements, total_pages)

    # --- Сборка чанков ---
    result: Dict[str, Any] = {
        "имя_файла": file_path.name,
        "путь": str(file_path.resolve()),
    }

    for chunk in chunks:
        name = chunk["name"]
        if chunk_filter and name != chunk_filter:
            continue

        pages = chunk.get("pages", [])

        # "all" или полный список → весь документ
        if pages == "all":
            result[name] = "\n\n".join(elements)
            continue

        # Резолвим номера страниц
        resolved = []
        for p in pages:
            if p < 0:
                resolved.append(total_pages + p + 1)
            else:
                resolved.append(p)

        # Собираем текст из указанных "страниц"
        parts = []
        for page_num in resolved:
            if page_num < 1 or page_num > total_pages:
                logger.warning(f"  Чанк '{name}': страница {page_num} вне диапазона (1-{total_pages})")
                continue
            page_text = pages_content.get(page_num, "")
            if page_text.strip():
                parts.append(page_text)

        result[name] = "\n\n".join(parts)

    return result


def _extract_elements_in_order(doc: Document) -> List[str]:
    """
    Извлекает все элементы документа (абзацы и таблицы) в порядке появления.

    DOCX хранит абзацы и таблицы как siblings в document.body.
    Обходим XML напрямую, чтобы сохранить правильный порядок.

    Args:
        doc: объект Document (python-docx)

    Returns:
        List[str] — элементы документа в порядке появления
    """
    from docx.oxml.ns import qn

    elements = []
    body = doc.element.body

    # Маппинг XML-элементов таблиц → python-docx Table objects
    table_elements = {}
    for table in doc.tables:
        table_elements[table._element] = table

    for child in body:
        tag = child.tag

        if tag == qn("w:p"):
            # Параграф
            text = child.text or ""
            # Собираем текст из всех run-элементов
            runs_text = []
            for r in child.findall(qn("w:r")):
                for t in r.findall(qn("w:t")):
                    if t.text:
                        runs_text.append(t.text)
            full_text = "".join(runs_text).strip()
            if full_text:
                elements.append(full_text)

        elif tag == qn("w:tbl"):
            # Таблица
            table = table_elements.get(child)
            if table:
                html = _table_to_html(table)
                if html.strip():
                    elements.append(html)

    return elements


def _split_into_pages(elements: List[str], total_pages: int) -> Dict[int, str]:
    """
    Разбивает список элементов на N приблизительно равных частей ("страниц").

    Args:
        elements: список текстовых элементов
        total_pages: количество "страниц"

    Returns:
        Dict[int, str] — номер страницы (1-based) → текст
    """
    if total_pages <= 1:
        return {1: "\n\n".join(elements)}

    # Делим по количеству элементов
    chunk_size = max(1, len(elements) // total_pages)
    pages = {}
    for i in range(total_pages):
        start = i * chunk_size
        if i == total_pages - 1:
            # Последняя страница забирает остаток
            end = len(elements)
        else:
            end = start + chunk_size
        pages[i + 1] = "\n\n".join(elements[start:end])

    return pages


def _table_to_html(table) -> str:
    """
    Конвертирует таблицу python-docx в HTML <table>.

    Args:
        table: объект docx.table.Table

    Returns:
        str — HTML-таблица
    """
    rows_html = []
    for row in table.rows:
        cells_html = []
        for cell in row.cells:
            cell_text = cell.text.strip().replace("\n", " ")
            cells_html.append(f"  <td>{cell_text}</td>")
        rows_html.append("<tr>\n" + "\n".join(cells_html) + "\n</tr>")

    if not rows_html:
        return ""

    return "<table>\n" + "\n".join(rows_html) + "\n</table>"
