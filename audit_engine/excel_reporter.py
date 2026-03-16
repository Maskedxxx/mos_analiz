#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Экспорт результатов аудита в Excel.

Формирует отчёт с колонками:
- №  — номер правила
- Проверка — название правила
- Целевой документ — что фактически в документе
- Различие — что должно быть или в чём проблема
"""

from pathlib import Path
from typing import Any, Dict, List

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# Названия столбцов для отображения
_DISPLAY_HEADERS = ["№", "Проверка", "Целевой документ", "Различие"]
# Ключи в данных
_DATA_KEYS = ["rule_index", "rule_title", "Целевой документ", "Различие"]

# Стили
_FONT_SIZE = 14
_HEADER_FONT = Font(name="Calibri", size=_FONT_SIZE, bold=True, color="FFFFFF")
_HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
_CELL_FONT = Font(name="Calibri", size=_FONT_SIZE)
_WRAP_ALIGNMENT = Alignment(wrap_text=True, vertical="top")
_CENTER_ALIGNMENT = Alignment(horizontal="center", vertical="top")
_THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

# Минимальная ширина столбцов (в символах)
_MIN_WIDTHS = [6, 35, 40, 40]
# Максимальная ширина
_MAX_WIDTHS = [6, 50, 60, 60]


def save_to_excel(violations: List[Dict[str, Any]], output_path: str) -> None:
    """
    Сохраняет результаты аудита в форматированный Excel.

    Args:
        violations: список нарушений (list of dicts)
        output_path: путь для сохранения .xlsx
    """
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Результаты проверки"

    # --- Заголовки ---
    for col_idx, header in enumerate(_DISPLAY_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _CENTER_ALIGNMENT
        cell.border = _THIN_BORDER
    ws.row_dimensions[1].height = 30

    # --- Данные ---
    if violations:
        # Сортируем по rule_index
        sorted_violations = sorted(violations, key=lambda v: v.get("rule_index", 0))
        for row_idx, v in enumerate(sorted_violations, start=2):
            for col_idx, key in enumerate(_DATA_KEYS, start=1):
                value = v.get(key, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = _CELL_FONT
                cell.border = _THIN_BORDER
                # Первый столбец (№) — по центру, остальные — перенос текста
                if col_idx == 1:
                    cell.alignment = _CENTER_ALIGNMENT
                else:
                    cell.alignment = _WRAP_ALIGNMENT

    # --- Ширина столбцов ---
    for col_idx in range(1, len(_DISPLAY_HEADERS) + 1):
        # Находим максимальную длину контента в столбце
        max_len = len(str(ws.cell(row=1, column=col_idx).value))
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value:
                    # Для переноса строк берём самую длинную строку
                    lines = str(cell.value).split("\n")
                    longest = max(len(line) for line in lines)
                    max_len = max(max_len, longest)

        # Ограничиваем ширину
        min_w = _MIN_WIDTHS[col_idx - 1]
        max_w = _MAX_WIDTHS[col_idx - 1]
        # +2 символа запас для padding
        width = min(max(max_len + 2, min_w), max_w)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # --- Закрепить шапку ---
    ws.freeze_panes = "A2"

    wb.save(output_path)
