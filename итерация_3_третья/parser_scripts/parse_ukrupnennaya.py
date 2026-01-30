#!/usr/bin/env python3
"""
Парсер укрупненной схемы из XLSX:
  - тянет фигуры/коннекторы из drawing*.xml (openpyxl их теряет)
  - тянет ячейки, заголовок, комментарии
  - сопоставляет метрики под узлами (аннотации в строке под схемой)
  - строит простой граф (узлы + ребра) и сохраняет единую структуру в JSON

Использование:
  python scripts/parse_ukrupnennaya.py -i path/to/file.xlsx [-s Лист] [-o out.json]
"""

import argparse
import json
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wb": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}


def read_workbook_parts(xlsx_path: Path):
    """Return map: sheet_name -> worksheet xml target path."""
    with zipfile.ZipFile(xlsx_path) as z:
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        wb_rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in wb_rels}
        sheets = {}
        for sheet in wb_xml.find("wb:sheets", NS):
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            sheets[sheet.attrib["name"]] = rel_map[rid]
        return sheets


def worksheet_drawing_path(xlsx_path: Path, sheet_target: str):
    """Return drawing XML path for worksheet or None."""
    rel_path = f"xl/worksheets/_rels/{posixpath.basename(sheet_target)}.rels"
    with zipfile.ZipFile(xlsx_path) as z:
        if rel_path not in z.namelist():
            return None
        rels = ET.fromstring(z.read(rel_path))
        for rel in rels:
            if rel.attrib.get("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing":
                target = rel.attrib["Target"]  # usually ../drawings/drawing1.xml
                return posixpath.normpath(posixpath.join("xl/worksheets", target))
    return None


def _pos(node):
    return {
        "col": int(node.find("xdr:col", NS).text),
        "colOff": int(node.find("xdr:colOff", NS).text),
        "row": int(node.find("xdr:row", NS).text),
        "rowOff": int(node.find("xdr:rowOff", NS).text),
    }


def _to_cell(col_zero_based: int, row_zero_based: int) -> str:
    return f"{get_column_letter(col_zero_based + 1)}{row_zero_based + 1}"


def parse_drawing(xlsx_path: Path, drawing_path: str):
    """Return shapes and connectors from drawing xml."""
    shapes, connectors = [], []
    if drawing_path is None:
        return shapes, connectors

    with zipfile.ZipFile(xlsx_path) as z:
        if drawing_path not in z.namelist():
            return shapes, connectors
        root = ET.fromstring(z.read(drawing_path))

    for idx, anc in enumerate(root.findall("./", NS), start=1):
        anchor_type = anc.tag.split("}")[-1]
        p_from = anc.find("xdr:from", NS)
        p_to = anc.find("xdr:to", NS)
        bbox = None
        if p_from is not None and p_to is not None:
            fpos, tpos = _pos(p_from), _pos(p_to)
            bbox = {
                "from": {
                    "row": fpos["row"] + 1,
                    "col": fpos["col"] + 1,
                    "rowOff": fpos["rowOff"],
                    "colOff": fpos["colOff"],
                    "cell": _to_cell(fpos["col"], fpos["row"]),
                },
                "to": {
                    "row": tpos["row"] + 1,
                    "col": tpos["col"] + 1,
                    "rowOff": tpos["rowOff"],
                    "colOff": tpos["colOff"],
                    "cell": _to_cell(tpos["col"], tpos["row"]),
                },
            }

        content_el, content_kind = None, None
        for child in anc:
            name = child.tag.split("}")[-1]
            if name in ("sp", "cxnSp", "pic", "graphicFrame"):
                content_el, content_kind = child, name
                break
        if content_el is None:
            continue

        text_runs = [t.text for t in content_el.findall(".//a:t", NS) if t.text]
        text = " ".join(text_runs).strip()
        geom_el = content_el.find(".//a:prstGeom", NS)
        geom = geom_el.attrib.get("prst") if geom_el is not None else None

        def center(box):
            fx, tx = box["from"], box["to"]
            cx = (fx["col"] - 1 + fx["colOff"] / 1e9 + tx["col"] - 1 + tx["colOff"] / 1e9) / 2
            cy = (fx["row"] - 1 + fx["rowOff"] / 1e9 + tx["row"] - 1 + tx["rowOff"] / 1e9) / 2
            return {"x": cx, "y": cy}

        record = {
            "id": idx,
            "anchor_type": anchor_type,
            "content_type": content_kind,
            "text": text,
            "geometry": geom,
            "bbox": bbox,
        }
        if bbox:
            record["center"] = center(bbox)

        if content_kind == "cxnSp":
            connectors.append(record)
        else:
            shapes.append(record)
    return shapes, connectors


def parse_cells_and_comments(xlsx_path: Path, sheet_name: str):
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet_name]
    cells = []
    comments = []
    merged_ranges = []
    for m in ws.merged_cells.ranges:
        merged_ranges.append(
            {
                "min_col": m.min_col,
                "max_col": m.max_col,
                "min_row": m.min_row,
                "max_row": m.max_row,
                "coord": m.coord,
            }
        )
    for row in ws.iter_rows():
        for c in row:
            if c.value not in (None, ""):
                cells.append({"cell": c.coordinate, "value": c.value})
            if c.comment:
                comments.append(
                    {
                        "cell": c.coordinate,
                        "author": c.comment.author,
                        "text": c.comment.text,
                    }
                )
    return cells, comments, merged_ranges


def detect_title(cells):
    for c in cells:
        if c["cell"].startswith("B1"):
            return c["value"]
    # fallback: first long string in first three rows
    for c in cells:
        row = int("".join(ch for ch in c["cell"] if ch.isdigit()))
        if row <= 3 and isinstance(c["value"], str) and len(c["value"]) > 10:
            return c["value"]
    return None


def detect_metrics_row(cells, shapes):
    """Heuristic: choose row with max filled cells below top of shapes."""
    if not shapes:
        return None
    min_shape_row = min(s["bbox"]["from"]["row"] for s in shapes if s.get("bbox"))
    counts = defaultdict(int)
    for c in cells:
        row = int("".join(ch for ch in c["cell"] if ch.isdigit()))
        if row >= min_shape_row:
            counts[row] += 1
    if not counts:
        return None
    return max(counts.items(), key=lambda kv: kv[1])[0]


def attach_metrics_to_nodes(shapes, metrics_row_cells, merged_ranges, metrics_row):
    """
    Привязка метрик к узлам:
      - учитываем merged-ячейки, чтобы понять горизонтальный охват метрики;
      - если bbox узла пересекает диапазон метрики — связываем, при нескольких берем ближайшую к центру;
      - если пересечений нет, fallback на ближайшую по колонке метрику;
      - иначе linked_value = None.
    """
    metrics = []
    for c in metrics_row_cells:
        coord = c["cell"]
        col_letter = "".join(ch for ch in coord if ch.isalpha())
        col_idx = column_index_from_string(col_letter)
        span_min = span_max = col_idx
        for rng in merged_ranges:
            if rng["min_row"] <= metrics_row <= rng["max_row"] and rng["min_col"] <= col_idx <= rng["max_col"]:
                span_min, span_max = rng["min_col"], rng["max_col"]
                break
        metrics.append(
            {
                "col_start": span_min,
                "col_end": span_max,
                "cell": coord,
                "value": c["value"],
            }
        )

    for s in shapes:
        bbox = s.get("bbox")
        if not bbox:
            s["linked_value"] = None
            continue
        span_min = min(bbox["from"]["col"], bbox["to"]["col"])
        span_max = max(bbox["from"]["col"], bbox["to"]["col"])
        center_col = s.get("center", {}).get("x", 0) + 1

        overlaps = [
            m
            for m in metrics
            if not (span_max < m["col_start"] or span_min > m["col_end"])
        ]

        def mid(m):
            return (m["col_start"] + m["col_end"]) / 2

        if overlaps:
            nearest = min(overlaps, key=lambda m: abs(center_col - mid(m)))
            s["linked_value"] = {
                "col_start": nearest["col_start"],
                "col_end": nearest["col_end"],
                "cell": nearest["cell"],
                "value": nearest["value"],
            }
        else:
            s["linked_value"] = None


def build_edges(shapes_sorted, connectors):
    """Build edges using connectors if counts match, else fall back to L->R chain."""
    edges = []
    if connectors and len(connectors) == len(shapes_sorted) - 1:
        connectors_sorted = sorted(connectors, key=lambda c: c.get("center", {}).get("x", 0))
        for i, cn in enumerate(connectors_sorted):
            if i + 1 >= len(shapes_sorted):
                break
            edges.append({"from_id": shapes_sorted[i]["id"], "to_id": shapes_sorted[i + 1]["id"], "connector_id": cn["id"]})
        assumption = "linear left-to-right mapping by connector order (count = nodes-1)"
    else:
        # nearest left/right fallback
        for cn in connectors:
            cx = cn.get("center", {}).get("x")
            if cx is None:
                continue
            left = [s for s in shapes_sorted if s["center"]["x"] < cx]
            right = [s for s in shapes_sorted if s["center"]["x"] > cx]
            if not left or not right:
                continue
            src = max(left, key=lambda s: s["center"]["x"])
            tgt = min(right, key=lambda s: s["center"]["x"])
            edges.append({"from_id": src["id"], "to_id": tgt["id"], "connector_id": cn["id"]})
        if not connectors:
            # no connectors at all -> simple chain
            for i in range(len(shapes_sorted) - 1):
                edges.append({"from_id": shapes_sorted[i]["id"], "to_id": shapes_sorted[i + 1]["id"], "connector_id": None})
            assumption = "no connectors; chained by left-to-right"
        else:
            assumption = "connectors mapped by nearest left/right centers"
    return edges, assumption


def build_unified_structure(xlsx_path: Path, sheet_name: str):
    sheets_map = read_workbook_parts(xlsx_path)
    if sheet_name not in sheets_map:
        raise SystemExit(f"Sheet '{sheet_name}' not found. Available: {list(sheets_map)}")

    drawing_path = worksheet_drawing_path(xlsx_path, sheets_map[sheet_name])
    shapes, connectors = parse_drawing(xlsx_path, drawing_path)
    cells, comments, merged_ranges = parse_cells_and_comments(xlsx_path, sheet_name)

    title = detect_title(cells)

    # notes: heuristically все строки ниже метрик со свободным текстом
    metrics_row = detect_metrics_row(cells, shapes)
    metrics_row_cells = [
        c for c in cells if metrics_row and c["cell"][len("".join(filter(str.isalpha, c["cell"]))):] == str(metrics_row)
    ]

    # упорядочим узлы/коннекторы слева-направо для стабильности вывода
    shapes_sorted = sorted(shapes, key=lambda s: (s.get("center", {}).get("x", 0), s.get("center", {}).get("y", 0)))
    connectors_sorted = sorted(connectors, key=lambda c: c.get("center", {}).get("x", 0))

    attach_metrics_to_nodes(shapes_sorted, metrics_row_cells, merged_ranges, metrics_row)

    notes = [c for c in cells if metrics_row and int("".join(ch for ch in c["cell"] if ch.isdigit())) > metrics_row]

    edges, assumption = build_edges(shapes_sorted, connectors_sorted)

    unified = {
        "meta": {
            "workbook": str(xlsx_path),
            "sheet": sheet_name,
            "drawing_xml": drawing_path,
        },
        "title": title,
        "notes": notes,
        "comments": comments,
        "metrics_row": metrics_row,
        "graph": {
            "nodes": shapes_sorted,
            "edges": edges,
            "assumption": assumption,
        },
        "connectors": connectors_sorted,
        "cells": cells,
    }
    return unified


def main():
    parser = argparse.ArgumentParser(description="Парсер укрупненной схемы КПСЦ/спагетти.")
    parser.add_argument("-i", "--input", required=True, help="Путь к XLSX")
    parser.add_argument("-s", "--sheet", default="Укрупненная", help="Имя листа (по умолчанию 'Укрупненная')")
    parser.add_argument("-o", "--output", default=None, help="Путь для JSON (stdout если не задан)")
    args = parser.parse_args()

    xlsx_path = Path(args.input)
    if not xlsx_path.exists():
        raise SystemExit(f"File not found: {xlsx_path}")

    unified = build_unified_structure(xlsx_path, args.sheet)

    payload = json.dumps(unified, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(payload, encoding="utf-8")
        print(f"Wrote {args.output}")
    else:
        sys.stdout.write(payload)


if __name__ == "__main__":
    main()
