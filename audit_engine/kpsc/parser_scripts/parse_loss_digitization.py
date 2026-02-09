#!/usr/bin/env python3
"""
Парсер листа "Оцифровка потерь КПСЦ".
Определяет границы единой таблицы (заголовок начинается в строке с колонками
\"Описание проблемы\" и \"Вид потери\"), захватывает прямоугольник от строки
заголовков до последней непустой строки в этом диапазоне.
Учитывает merged‑ячейки, возвращает координаты, значения и метаданные.
"""

import argparse
import json
from pathlib import Path
from typing import Optional, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries


HEADER_KEYS = ("описание проблемы", "вид потери")


def find_header_row(ws) -> Optional[int]:
    """Находим строку заголовков по ключевым фразам."""
    for r in range(1, ws.max_row + 1):
        lower_vals = [str(c.value).lower() for c in ws[r] if isinstance(c.value, str)]
        if lower_vals and all(any(key in v for v in lower_vals) for key in HEADER_KEYS):
            return r
    return None


def compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    """Границы по непустым ячейкам строки заголовков (от первой до последней)."""
    left = None
    right = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v not in (None, ""):
            if left is None:
                left = c
            right = c
    if left is None:
        left = 1
        right = 1
    return left, right


def find_bottom_row(ws, header_row: int, left_col: int, right_col: int) -> int:
    """
    Ищем конец таблицы: первая строка после заголовка, где данные отсутствуют
    во всех колонках кроме порядкового номера (left_col). Строки с одним
    номером считаем пустыми.
    """
    for r in range(header_row + 1, ws.max_row + 1):
        has_data = any(
            ws.cell(row=r, column=c).value not in (None, "")
            for c in range(left_col + 1, right_col + 1)
        )
        if not has_data:
            return r - 1
    return ws.max_row


def build_merged_lookup(ws):
    lookup = {}
    for merge in ws.merged_cells.ranges:
        coord = merge.coord
        for r in range(merge.min_row, merge.max_row + 1):
            for c in range(merge.min_col, merge.max_col + 1):
                lookup[(r, c)] = coord
    return lookup


def extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            coord = f"{get_column_letter(c)}{r}"
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = (r == min_row_m and c == min_col_m)
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = ws.cell(row=r, column=c).value
            row_cells.append(
                {
                    "coord": coord,
                    "row": r,
                    "col": c,
                    "value": value,
                    "merge_range": merge_range,
                    "merge_anchor": anchor if merge_range else False,
                }
            )
        rows.append({"row": r, "cells": row_cells})
    return rows


def build_payload(xlsx_path: Path, sheet: str):
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet]

    header_row = find_header_row(ws)
    if not header_row:
        raise SystemExit("Не найден заголовок таблицы (колонки 'Описание проблемы' и 'Вид потери').")

    left_col, right_col = compute_col_bounds(ws, header_row)
    bottom_row = find_bottom_row(ws, header_row, left_col, right_col)
    rows = extract_table(ws, header_row, bottom_row, left_col, right_col)

    return {
        "meta": {"workbook": str(xlsx_path), "sheet": sheet, "header_row": header_row},
        "bounds": {
            "top_row": header_row,
            "bottom_row": bottom_row,
            "left_col": left_col,
            "right_col": right_col,
            "left_letter": get_column_letter(left_col),
            "right_letter": get_column_letter(right_col),
            "height": bottom_row - header_row + 1,
            "width": right_col - left_col + 1,
        },
        "rows": rows,
    }


def parse(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит лист 'Оцифровка потерь КПСЦ' и сохраняет в ocifrovka_poteri_v2.json."""
    payload = build_payload(xlsx_path, "Оцифровка потерь КПСЦ")
    output_path = output_dir / "ocifrovka_poteri_v2.json"
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    return payload


def main():
    ap = argparse.ArgumentParser(description="Парсер листа 'Оцифровка потерь КПСЦ'")
    ap.add_argument("-i", "--input", required=True, help="Путь к XLSX файлу")
    ap.add_argument("-s", "--sheet", default="Оцифровка потерь КПСЦ", help="Имя листа")
    ap.add_argument("-o", "--output", help="JSON файл вывода (stdout если не указан)")
    args = ap.parse_args()

    payload = build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding="utf-8")
        print(f"Wrote {args.output}")
    else:
        print(data)


if __name__ == "__main__":
    main()
