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
_DISPLAY_HEADERS = ["№", "Проверка", "Статус", "Целевой документ", "Различие"]

# Стили для статуса
_OK_FONT = Font(name="Calibri", size=14, bold=True, color="1F7A1F")
_FAIL_FONT = Font(name="Calibri", size=14, bold=True, color="CC0000")

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
_MIN_WIDTHS = [6, 35, 10, 40, 40]
# Максимальная ширина
_MAX_WIDTHS = [6, 50, 10, 60, 60]

# Заливка для строк ОК / FAIL
_OK_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
_FAIL_FILL = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")


def save_to_excel(
    violations: List[Dict[str, Any]],
    output_path: str,
    all_rules: List[Any] = None,
) -> None:
    """
    Сохраняет результаты аудита в форматированный Excel.

    Показывает ВСЕ правила: ОК если нарушений нет, FAIL с деталями если есть.

    Args:
        violations: список нарушений (list of dicts)
        output_path: путь для сохранения .xlsx
        all_rules: список всех RuleSpec (если передан — показываем все правила)
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

    # --- Собираем данные: все правила с ОК/FAIL ---
    # Группируем нарушения по rule_index
    violations_by_rule = {}
    for v in violations:
        idx = v.get("rule_index", 0)
        violations_by_rule.setdefault(idx, []).append(v)

    rows_data = []

    if all_rules:
        # Показываем ВСЕ правила
        for rule in sorted(all_rules, key=lambda r: r.index):
            rule_violations = violations_by_rule.get(rule.index, [])
            if rule_violations:
                # FAIL — строка для каждого нарушения
                for v in rule_violations:
                    rows_data.append({
                        "index": rule.index,
                        "title": rule.title,
                        "status": "FAIL",
                        "target": v.get("Целевой документ", ""),
                        "diff": v.get("Различие", ""),
                    })
            else:
                # ОК
                rows_data.append({
                    "index": rule.index,
                    "title": rule.title,
                    "status": "ОК",
                    "target": "",
                    "diff": "",
                })
    else:
        # Fallback: только нарушения (старое поведение)
        for v in sorted(violations, key=lambda v: v.get("rule_index", 0)):
            rows_data.append({
                "index": v.get("rule_index", ""),
                "title": v.get("rule_title", ""),
                "status": "FAIL",
                "target": v.get("Целевой документ", ""),
                "diff": v.get("Различие", ""),
            })

    # --- Записываем строки ---
    for row_idx, row in enumerate(rows_data, start=2):
        is_ok = row["status"] == "ОК"
        fill = _OK_FILL if is_ok else _FAIL_FILL

        values = [row["index"], row["title"], row["status"], row["target"], row["diff"]]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = _CELL_FONT
            cell.border = _THIN_BORDER
            cell.fill = fill

            if col_idx == 1:  # №
                cell.alignment = _CENTER_ALIGNMENT
            elif col_idx == 3:  # Статус
                cell.alignment = _CENTER_ALIGNMENT
                cell.font = _OK_FONT if is_ok else _FAIL_FONT
            else:
                cell.alignment = _WRAP_ALIGNMENT

    # --- Ширина столбцов ---
    for col_idx in range(1, len(_DISPLAY_HEADERS) + 1):
        max_len = len(str(ws.cell(row=1, column=col_idx).value))
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value:
                    lines = str(cell.value).split("\n")
                    longest = max(len(line) for line in lines)
                    max_len = max(max_len, longest)

        min_w = _MIN_WIDTHS[col_idx - 1]
        max_w = _MAX_WIDTHS[col_idx - 1]
        width = min(max(max_len + 2, min_w), max_w)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # --- Закрепить шапку ---
    ws.freeze_panes = "A2"

    wb.save(output_path)
