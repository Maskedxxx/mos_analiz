#!/usr/bin/env python3
"""
Парсер верхнего блока листа "КПСЦ" (A1:AE13).
Извлекает ключевые поля + сырые ячейки, merged-диапазоны и комментарии.
"""
import argparse
import json
from pathlib import Path
from typing import Dict, Any, List
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

REGION_MAX_ROW = 13
REGION_MAX_COL = 31  # AE


def collect_cells(ws):
    cells = []
    comments = []
    for r in range(1, REGION_MAX_ROW + 1):
        for c in range(1, REGION_MAX_COL + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if val not in (None, ""):
                cells.append({
                    "coord": f"{get_column_letter(c)}{r}",
                    "row": r,
                    "col": c,
                    "value": val,
                })
            if cell.comment:
                comments.append({
                    "coord": f"{get_column_letter(c)}{r}",
                    "row": r,
                    "col": c,
                    "author": cell.comment.author,
                    "text": cell.comment.text,
                })
    return cells, comments


def collect_merged(ws):
    merged = []
    for m in ws.merged_cells.ranges:
        if m.min_row <= REGION_MAX_ROW and m.min_col <= REGION_MAX_COL:
            merged.append({
                "coord": m.coord,
                "min_row": m.min_row,
                "max_row": m.max_row,
                "min_col": m.min_col,
                "max_col": m.max_col,
            })
    return merged


def extract_fields(cells: List[Dict[str, Any]]):
    def val(coord):
        for c in cells:
            if c["coord"] == coord:
                return c["value"]
        return None

    return {
        "title": val("B1"),
        "flow_name": val("C4"),
        "responsible": val("C5"),
        "date_developed": val("C6"),
        "date_implementation": val("C7"),
        "compiled_by": val("C8"),
        "unit_note": val("B12"),
        "first_operation": val("D12"),
        "takt_time": val("B13"),
    }


def build_payload(xlsx: Path, sheet: str):
    wb = load_workbook(xlsx, data_only=True)
    ws = wb[sheet]
    cells, _comments = collect_cells(ws)
    fields = extract_fields(cells)

    return {
        "meta": {
            "workbook": str(xlsx),
            "sheet": sheet,
            "region": "A1:AE13",
        },
        "fields": fields,
    }


def parse(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит верхний блок КПСЦ и сохраняет результат в output_dir/kpsc_header_v2.json."""
    payload = build_payload(xlsx_path, "КПСЦ")
    output_path = output_dir / "kpsc_header_v2.json"
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    return payload


def main():
    ap = argparse.ArgumentParser(description="Парсер верхнего блока КПСЦ (A1:AE13)")
    ap.add_argument("-i", "--input", required=True)
    ap.add_argument("-s", "--sheet", default="КПСЦ")
    ap.add_argument("-o", "--output")
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
