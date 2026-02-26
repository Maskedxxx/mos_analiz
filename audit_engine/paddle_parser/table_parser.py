#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсер табличного формата PaddleOCR-VL → HTML / markdown.

PaddleOCR-VL возвращает таблицы в формате:
- <fcel> — ячейка с текстом (filled cell)
- <ecel> — пустая ячейка (empty cell)
- <lcel> — продолжение ячейки влево (colspan, left-continuing)
- <ucel> — продолжение ячейки вверх (rowspan, upper-continuing)
- <nl>   — конец строки (newline)

Конверсия в HTML сохраняет colspan/rowspan.
Конверсия в markdown теряет merge-информацию (пустые ячейки).
"""

import logging
import re
from dataclasses import dataclass, field
from typing import List, Optional

logger = logging.getLogger(__name__)


@dataclass
class Cell:
    """Одна ячейка таблицы."""
    text: str = ""
    cell_type: str = "empty"   # filled / empty / lcel / ucel
    colspan: int = 1
    rowspan: int = 1
    covered: bool = False      # True если ячейка перекрыта colspan/rowspan соседа


def _tokenize_row(row_str: str) -> List[Cell]:
    """
    Парсит строку PaddleOCR-VL в список Cell с сохранением типа тега.

    Args:
        row_str: строка вида "<fcel>Текст<ecel><lcel><fcel>Текст2".

    Returns:
        Список Cell.
    """
    cells = []
    tokens = re.split(r"(<fcel>|<ecel>|<lcel>|<ucel>)", row_str)

    current_type = None
    current_text = ""

    for token in tokens:
        if token in ("<fcel>", "<ecel>", "<lcel>", "<ucel>"):
            if current_type is not None:
                cells.append(Cell(
                    text=current_text.strip(),
                    cell_type=current_type,
                ))
            type_map = {
                "<fcel>": "filled",
                "<ecel>": "empty",
                "<lcel>": "lcel",
                "<ucel>": "ucel",
            }
            current_type = type_map[token]
            current_text = ""
        else:
            current_text += token

    if current_type is not None:
        cells.append(Cell(
            text=current_text.strip(),
            cell_type=current_type,
        ))

    return cells


def _build_grid(raw: str) -> Optional[List[List[Cell]]]:
    """
    Парсит сырой PaddleOCR-VL текст в сетку Cell с colspan/rowspan.

    Обрабатывает <lcel> (увеличивает colspan предыдущей fcel) и
    <ucel> (увеличивает rowspan ячейки сверху).

    Args:
        raw: сырой текст с тегами.

    Returns:
        Сетка Cell или None если формат не распознан.
    """
    if "<fcel>" not in raw and "<ecel>" not in raw:
        return None

    rows_raw = re.split(r"<nl>", raw)
    rows_raw = [r.strip() for r in rows_raw if r.strip()]
    if not rows_raw:
        return None

    # Парсим каждую строку в список Cell
    grid: List[List[Cell]] = []
    for row_str in rows_raw:
        cells = _tokenize_row(row_str)
        if cells:
            grid.append(cells)

    if not grid:
        return None

    # Обработка <lcel>: увеличиваем colspan предыдущей filled/empty ячейки
    for row in grid:
        i = 0
        while i < len(row):
            if row[i].cell_type == "lcel":
                # Ищем предыдущую не-lcel ячейку
                for j in range(i - 1, -1, -1):
                    if row[j].cell_type != "lcel":
                        row[j].colspan += 1
                        row[i].covered = True
                        break
            i += 1

    # Определяем логическую ширину таблицы (с учётом colspan)
    max_logical_cols = 0
    for row in grid:
        width = sum(c.colspan for c in row if not c.covered)
        max_logical_cols = max(max_logical_cols, width)

    # Строим карту позиций для обработки <ucel>
    # position_map[row][col] → ссылка на ячейку-источник (для rowspan)
    # Сначала раскладываем ячейки по логическим колонкам
    position_map: List[List[Optional[Cell]]] = []
    for row_idx, row in enumerate(grid):
        pos_row: List[Optional[Cell]] = [None] * max_logical_cols
        col = 0
        for cell in row:
            if cell.covered:
                continue
            # Пропускаем колонки, занятые rowspan'ом сверху
            while col < max_logical_cols and pos_row[col] is not None:
                col += 1
            if col >= max_logical_cols:
                break

            if cell.cell_type == "ucel":
                # Ищем ячейку-источник сверху в этой колонке
                cell.covered = True
                for prev_row_idx in range(row_idx - 1, -1, -1):
                    if prev_row_idx < len(position_map):
                        source = position_map[prev_row_idx][col]
                        if source is not None and not source.covered:
                            source.rowspan += 1
                            break
                        elif source is not None and source.covered:
                            # Ищем оригинальную ячейку
                            break
                col += 1
            else:
                pos_row[col] = cell
                # Заполняем colspan-позиции
                for cs in range(1, cell.colspan):
                    if col + cs < max_logical_cols:
                        pos_row[col + cs] = cell
                col += cell.colspan

        # Копируем rowspan из предыдущих строк
        if position_map:
            prev = position_map[-1]
            for c in range(max_logical_cols):
                if pos_row[c] is None and prev[c] is not None:
                    src = prev[c]
                    if not src.covered and src.rowspan > 1:
                        # Ещё действует rowspan
                        remaining = src.rowspan - (row_idx - _find_origin_row(position_map, src))
                        if remaining > 0:
                            pos_row[c] = src

        position_map.append(pos_row)

    n_rows = len(grid)
    n_cols = max_logical_cols
    logger.info(f"Таблица: {n_rows} строк × {n_cols} колонок")

    return grid


def _find_origin_row(position_map: List[List[Optional[Cell]]], cell: Cell) -> int:
    """Находит номер строки, где ячейка впервые появляется."""
    for row_idx, row in enumerate(position_map):
        if cell in row:
            return row_idx
    return 0


def parse_paddle_table_html(raw: str) -> str:
    """
    Конвертирует табличный формат PaddleOCR-VL в HTML-таблицу.

    Сохраняет colspan и rowspan.

    Args:
        raw: сырой текст с тегами <fcel>, <ecel>, <lcel>, <ucel>, <nl>.

    Returns:
        HTML-таблица. Если формат не распознан, возвращает сырой текст.
    """
    grid = _build_grid(raw)
    if grid is None:
        return raw

    lines = ["<table>"]
    for row in grid:
        lines.append("<tr>")
        for cell in row:
            if cell.covered:
                continue
            attrs = ""
            if cell.colspan > 1:
                attrs += f' colspan="{cell.colspan}"'
            if cell.rowspan > 1:
                attrs += f' rowspan="{cell.rowspan}"'
            text = cell.text if cell.text else ""
            lines.append(f"  <td{attrs}>{text}</td>")
        lines.append("</tr>")
    lines.append("</table>")

    return "\n".join(lines)


def parse_paddle_table(raw: str) -> str:
    """
    Конвертирует табличный формат PaddleOCR-VL в markdown-таблицу.

    Markdown не поддерживает colspan/rowspan — merged cells становятся пустыми.

    Args:
        raw: сырой текст с тегами <fcel>, <ecel>, <nl>.

    Returns:
        Markdown-таблица. Если формат не распознан, возвращает сырой текст.
    """
    grid = _build_grid(raw)
    if grid is None:
        return raw

    # Собираем простые строки текста (без colspan/rowspan)
    max_cols = 0
    text_rows: List[List[str]] = []
    for row in grid:
        texts = []
        for cell in row:
            if cell.covered:
                texts.append("")
            else:
                texts.append(cell.text)
                # Добавляем пустые ячейки для colspan
                for _ in range(cell.colspan - 1):
                    texts.append("")
        max_cols = max(max_cols, len(texts))
        text_rows.append(texts)

    # Выравниваем
    for row in text_rows:
        while len(row) < max_cols:
            row.append("")

    # Собираем markdown
    lines = []
    header = "| " + " | ".join(text_rows[0]) + " |"
    lines.append(header)
    sep = "|" + "|".join(["---"] * max_cols) + "|"
    lines.append(sep)
    for row in text_rows[1:]:
        line = "| " + " | ".join(row) + " |"
        lines.append(line)

    return "\n".join(lines)


def is_table_format(text: str) -> bool:
    """
    Определяет, содержит ли текст табличный формат PaddleOCR-VL.

    Args:
        text: текст от VLM.

    Returns:
        True если текст в табличном формате.
    """
    return "<fcel>" in text or ("<ecel>" in text and "<nl>" in text) or "<lcel>" in text
