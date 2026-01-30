#!/usr/bin/env python3
"""
Парсер графика между таблицами 7 и 8 на листе "КПСЦ".
Извлекает chart1.xml (bar chart), серию, значения, опорные ячейки и bbox на листе.
"""
import argparse
import json
import posixpath
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Tuple, Optional

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wb": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}


def read_sheet_rel_paths(xlsx: Path, sheet_name: str) -> Tuple[str, str]:
    """Return sheet_target, drawing_xml_path for given sheet."""
    with zipfile.ZipFile(xlsx) as z:
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        wb_rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in wb_rels}
        sheet_target = None
        for sheet in wb_xml.find("wb:sheets", NS):
            if sheet.attrib["name"] == sheet_name:
                rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
                sheet_target = rel_map[rid]
                break
        if not sheet_target:
            raise SystemExit(f"Sheet '{sheet_name}' not found")
        rel_path = f"xl/worksheets/_rels/{posixpath.basename(sheet_target)}.rels"
        rels = ET.fromstring(z.read(rel_path))
        drawing_path = None
        for rel in rels:
            if rel.attrib.get("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing":
                drawing_path = posixpath.normpath(posixpath.join("xl/worksheets", rel.attrib["Target"]))
                break
        if not drawing_path:
            raise SystemExit("No drawing for sheet")
    return sheet_target, drawing_path


def find_chart_anchor_and_rid(xlsx: Path, drawing_path: str) -> Tuple[dict, str]:
    """Return bbox (from/to) and chart rId for the first graphicFrame with chart (chart1)."""
    with zipfile.ZipFile(xlsx) as z:
        root = ET.fromstring(z.read(drawing_path))
        d_rels_path = posixpath.join(posixpath.dirname(drawing_path), "_rels", posixpath.basename(drawing_path) + ".rels")
        d_rels = ET.fromstring(z.read(d_rels_path)) if d_rels_path in z.namelist() else None
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in d_rels} if d_rels is not None else {}

    for anc in root.findall("./", NS):
        content = anc.find("xdr:graphicFrame", NS)
        if content is None:
            continue
        chart_el = content.find(".//c:chart", {"c": NS["c"], "r": NS["r"]})
        if chart_el is None:
            continue
        rid = chart_el.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        p_from = anc.find("xdr:from", NS)
        p_to = anc.find("xdr:to", NS)
        def _pos(n):
            return int(n.find("xdr:col", NS).text), int(n.find("xdr:row", NS).text)
        fcol, frow = _pos(p_from)
        tcol, trow = _pos(p_to)
        bbox = {
            "from": {"col": fcol + 1, "row": frow + 1, "cell": f"{get_column_letter(fcol+1)}{frow+1}"},
            "to": {"col": tcol + 1, "row": trow + 1, "cell": f"{get_column_letter(tcol+1)}{trow+1}"},
        }
        target = rel_map.get(rid)
        return bbox, target
    raise SystemExit("Chart graphicFrame not found")


def expand_formula_refs(formula: str) -> List[Tuple[int, int]]:
    """Expand formula like '(Sheet!$E$57,Sheet!$L$57:$Q$57,...)' to list of (row,col)."""
    if not formula:
        return []
    # strip sheet name and parentheses
    # split list parts, keep sheet name optional per part
    formula = formula.strip()
    if formula.startswith("(") and formula.endswith(")"):
        formula = formula[1:-1]
    parts = [p.strip() for p in formula.split(",") if p.strip()]
    coords = []
    for part in parts:
        if "!" in part:
            sheet_part, part = part.split("!", 1)
        if ":" in part:
            min_col, min_row, max_col, max_row = range_boundaries(part)
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    coords.append((r, c))
        else:
            # single cell like $E$57
            letters = "".join(ch for ch in part if ch.isalpha())
            numbers = "".join(ch for ch in part if ch.isdigit())
            if not numbers:
                continue
            col = sum((ord(ch) - 64) * (26 ** i) for i, ch in enumerate(reversed(letters)))
            row = int(numbers)
            coords.append((row, col))
    return coords


def read_values(ws, refs: List[Tuple[int, int]]) -> List[dict]:
    out = []
    # determine header row above: find nearest row above min_row that contains "Показатель" in column B
    min_row = min(r for r, _ in refs)
    header_row = None
    for r in range(min_row - 1, 0, -1):
        if ws.cell(row=r, column=2).value == "Показатель":
            header_row = r
            break
    for r, c in refs:
        coord = f"{get_column_letter(c)}{r}"
        header = ws.cell(row=header_row, column=c).value if header_row else None
        out.append({"coord": coord, "row": r, "col": c, "value": ws.cell(row=r, column=c).value, "header": header})
    return out


def parse_chart(xlsx: Path, sheet_name: str):
    sheet_target, drawing_path = read_sheet_rel_paths(xlsx, sheet_name)
    bbox, chart_target = find_chart_anchor_and_rid(xlsx, drawing_path)
    chart_path = posixpath.normpath(posixpath.join("xl/drawings", chart_target)) if chart_target.startswith("../") else chart_target

    wb = load_workbook(xlsx, data_only=True)
    ws = wb[sheet_name]

    with zipfile.ZipFile(xlsx) as z:
        chart_xml = ET.fromstring(z.read(chart_path))

    plot_area = chart_xml.find("c:chart/c:plotArea", NS)
    series_out = []
    line_series = None
    bars = []
    for chart_type, tag in [("bar", "barChart"), ("line", "lineChart")]:
        chart = plot_area.find(f"c:{tag}", NS)
        if chart is None:
            continue
        for ser in chart.findall("c:ser", NS):
            name_f = ser.find("c:tx/c:strRef/c:f", NS)
            name_formula = name_f.text if name_f is not None else None
            name_cells = expand_formula_refs(name_formula) if name_formula else []
            name_values = [ws.cell(row=r, column=c).value for r, c in name_cells]

            val_f = ser.find("c:val/c:numRef/c:f", NS)
            val_formula = val_f.text if val_f is not None else None
            val_refs = expand_formula_refs(val_formula)
            values = read_values(ws, val_refs)

            series_obj = {
                "chart_type": chart_type,
                "name_formula": name_formula,
                "name_values": [v for v in name_values if v not in (None, "")],
                "value_formula": val_formula,
                "points": values,
            }
            series_out.append(series_obj)
            if chart_type == "line" and line_series is None:
                line_series = series_obj
            elif chart_type == "bar":
                bars.append(series_obj)

    intersections = []
    if line_series:
        line_val = next((pt["value"] for pt in line_series["points"] if pt["value"] is not None), None)
        if line_val is not None:
            for bar in bars:
                hits = [
                    {"coord": pt["coord"], "header": pt["header"], "value": pt["value"]}
                    for pt in bar["points"]
                    if pt["value"] is not None and pt["value"] >= line_val
                ]
                intersections.append(
                    {
                        "bar_series_name": bar["name_values"],
                        "line_series_name": line_series["name_values"],
                        "line_value": line_val,
                        "points": hits,
                    }
                )

    return {
        "meta": {
            "workbook": str(xlsx),
            "sheet": sheet_name,
            "drawing_xml": drawing_path,
            "chart_xml": chart_path,
        },
        "anchor": bbox,
        "series": series_out,
        "intersections": intersections,
    }


def main():
    ap = argparse.ArgumentParser(description="Парсер графика (chart1) между таблицами 7 и 8")
    ap.add_argument("-i", "--input", required=True)
    ap.add_argument("-s", "--sheet", default="КПСЦ")
    ap.add_argument("-o", "--output")
    args = ap.parse_args()

    payload = parse_chart(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding="utf-8")
        print(f"Wrote {args.output}")
    else:
        print(data)


if __name__ == "__main__":
    main()
