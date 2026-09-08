# START_MODULE_CONTRACT
# PURPOSE: Генератор Excel-отчёта по результатам аудита. Показывает ВСЕ правила (PASS+FAIL), разделяет «Базовая»/«Методическая» проверки цветом и колонкой «Тип проверки».
# INPUTS: violations (list of dict), output_path (str), all_rules (legacy RuleSpec) или multi_rules (rules_multi.json + rules_methodology.json), warnings (list of str — предупреждения парсера).
# OUTPUTS: .xlsx файл с шапкой, цветными ячейками PASS=зелёная/FAIL=розовая (для базовых) или голубая/розовая (для методических).
# KEYWORDS: excel, openpyxl, report, multi-rule, methodology.
# LINKS: src/audit/engine.py (AuditEngine.run пишет финальный отчёт через `save_to_excel`), src/doc_type_validators/plan_grafik.py (тоже использует).
# RATIONALE: Excel-отчёт — отдельная отчётная утилита; не зависит ни от engine, ни от runner-ов. Все стили/размеры/цвета — модуль-локальные приватные константы.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
from pathlib import Path
from typing import Any, Dict, List, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
# END_IMPORTS


# START_STYLE_CONSTANTS
# PURPOSE: Стили оформления Excel-таблицы (приватные, не экспортируются).
_DISPLAY_HEADERS = ["№", "Проверка", "Тип проверки", "Статус", "Источник (МР/МУ)", "Целевой документ", "Различие", "Обоснование"]
_FONT_SIZE = 14
_HEADER_FONT = Font(name="Calibri", size=_FONT_SIZE, bold=True, color="FFFFFF")
_HEADER_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
_CELL_FONT = Font(name="Calibri", size=_FONT_SIZE)
_WRAP_ALIGNMENT = Alignment(wrap_text=True, vertical="top")
_CENTER_ALIGNMENT = Alignment(horizontal="center", vertical="top")
_THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
_MIN_WIDTHS = [6, 35, 18, 10, 30, 40, 40, 50]
_MAX_WIDTHS = [6, 50, 18, 10, 45, 60, 60, 70]
_OK_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
_FAIL_FILL = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
_OK_FONT = Font(name="Calibri", size=_FONT_SIZE, bold=True, color="1F7A1F")
_FAIL_FONT = Font(name="Calibri", size=_FONT_SIZE, bold=True, color="CC0000")
_METH_OK_FILL = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")
_METH_FAIL_FILL = PatternFill(start_color="FADBD8", end_color="FADBD8", fill_type="solid")
# END_STYLE_CONSTANTS


def _layer_label(layer: str) -> str:
    """Человекочитаемый лейбл для колонки «Тип проверки»."""
    if layer == "methodology":
        return "Методическая"
    return "Базовая"


def _format_source(source_ref: str, nature: str) -> str:
    """Собирает строку источника для столбца «Источник (МР/МУ)» в Excel (порт из прода)."""
    if not source_ref and not nature:
        return ""
    nature_label = ""
    if nature:
        n = nature.strip().lower()
        if n.startswith("обяз"):
            nature_label = " (обязательно)"
        elif n.startswith("реком"):
            nature_label = " (рекомендательно)"
        else:
            nature_label = f" ({nature})"
    return f"{source_ref}{nature_label}".strip()


# START_SAVE_TO_EXCEL
def save_to_excel(
    violations: List[Dict[str, Any]],
    output_path: str,
    multi_rules: Optional[List[Dict[str, Any]]] = None,
    warnings: Optional[List[str]] = None,
) -> None:
    """
    Назначение:
        Сохраняет результаты аудита в форматированный Excel.
        Показывает ВСЕ правила: ОК если нарушений нет, FAIL с деталями если есть.
        Колонка «Тип проверки» разделяет базовые и методические правила.

    Вход:
        violations: список нарушений (list of dicts, каждый с полем "layer").
        output_path: путь для сохранения .xlsx.
        multi_rules: список dict-правил из rules_multi.json + rules_methodology.json.
            Если задан — итерируем по нему (показываем PASS+FAIL); иначе — только по
            `violations` (без PASS-строк).
        warnings: предупреждения парсера (например, не распознанные страницы PDF). Если
            список непустой — добавляется второй лист «Предупреждения», по строке на каждое.

    Выход:
        None (пишет файл по `output_path`).

    Логика:
        1. Группируем violations по (rule_index, layer).
        2. Если задан `multi_rules` — итерируем по нему (PASS+FAIL); иначе — только violations.
        3. Каждое правило → одна строка PASS («ОК») или N строк FAIL (по числу нарушений).
        4. Цвет ячеек: base PASS=зелёная, base FAIL=розовая, methodology PASS=голубая, methodology FAIL=светло-розовая.
        5. Авто-ширина колонок в пределах [_MIN_WIDTHS..._MAX_WIDTHS].
        6. Закреплена шапка (`freeze_panes='A2'`).
        7. Если есть `warnings` — второй лист «Предупреждения».
    """
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Результаты проверки"

    # Шапка
    for col_idx, header in enumerate(_DISPLAY_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _CENTER_ALIGNMENT
        cell.border = _THIN_BORDER
    ws.row_dimensions[1].height = 30

    # Группируем violations по (rule_index, layer)
    violations_by_rule: Dict[Any, List[Dict[str, Any]]] = {}
    for v in violations:
        key = (v.get("rule_index", 0), v.get("layer", "base"))
        violations_by_rule.setdefault(key, []).append(v)

    # Собираем строки данных
    rows_data: List[Dict[str, Any]] = []
    if multi_rules:
        for rule in sorted(multi_rules, key=lambda r: (0 if r.get("layer", "base") == "base" else 1, r.get("index", 0))):
            idx = rule.get("index", 0)
            layer = rule.get("layer", "base")
            key = (idx, layer)
            source = _format_source(rule.get("source_ref", ""), rule.get("nature", ""))
            rule_violations = violations_by_rule.get(key, [])
            if rule_violations:
                for v in rule_violations:
                    rows_data.append({
                        "index": idx, "title": rule.get("title", ""), "layer": layer,
                        "status": "FAIL", "source": source,
                        "target": v.get("Целевой документ", ""), "diff": v.get("Различие", ""),
                        "reasoning": v.get("Обоснование", ""),
                    })
            else:
                rows_data.append({
                    "index": idx, "title": rule.get("title", ""), "layer": layer,
                    "status": "ОК", "source": source, "target": "", "diff": "", "reasoning": "",
                })
    else:
        for v in sorted(violations, key=lambda v: v.get("rule_index", 0)):
            rows_data.append({
                "index": v.get("rule_index", ""), "title": v.get("правило", ""),
                "layer": v.get("layer", "base"), "status": "FAIL",
                "source": _format_source(v.get("Источник", ""), v.get("Природа", "")),
                "target": v.get("Целевой документ", ""), "diff": v.get("Различие", ""),
                "reasoning": v.get("Обоснование", ""),
            })

    # Раскраска ячеек
    for row_idx, row in enumerate(rows_data, start=2):
        is_ok = row["status"] == "ОК"
        is_meth = row.get("layer") == "methodology"
        if is_meth:
            fill = _METH_OK_FILL if is_ok else _METH_FAIL_FILL
        else:
            fill = _OK_FILL if is_ok else _FAIL_FILL
        values = [row["index"], row["title"], _layer_label(row.get("layer", "base")),
                  row["status"], row.get("source", ""), row["target"], row["diff"],
                  row.get("reasoning", "")]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = _CELL_FONT
            cell.border = _THIN_BORDER
            cell.fill = fill
            if col_idx == 1:
                cell.alignment = _CENTER_ALIGNMENT
            elif col_idx in (3, 4):
                cell.alignment = _CENTER_ALIGNMENT
                if col_idx == 4:
                    cell.font = _OK_FONT if is_ok else _FAIL_FONT
            else:
                cell.alignment = _WRAP_ALIGNMENT

    # Авто-ширина колонок
    for col_idx in range(1, len(_DISPLAY_HEADERS) + 1):
        max_len = len(str(ws.cell(row=1, column=col_idx).value))
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value:
                    lines = str(cell.value).split("\n")
                    longest = max((len(line) for line in lines))
                    max_len = max(max_len, longest)
        min_w = _MIN_WIDTHS[col_idx - 1]
        max_w = _MAX_WIDTHS[col_idx - 1]
        width = min(max(max_len + 2, min_w), max_w)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.freeze_panes = "A2"

    # Лист «Предупреждения» — только если парсер их вернул (например, не распознанные страницы PDF).
    if warnings:
        ws_warn = wb.create_sheet("Предупреждения")
        head = ws_warn.cell(row=1, column=1, value="Предупреждение")
        head.font = _HEADER_FONT
        head.fill = _HEADER_FILL
        head.border = _THIN_BORDER
        for row_idx, text in enumerate(warnings, start=2):
            cell = ws_warn.cell(row=row_idx, column=1, value=text)
            cell.font = _CELL_FONT
            cell.alignment = _WRAP_ALIGNMENT
            cell.border = _THIN_BORDER
        ws_warn.column_dimensions["A"].width = 120

    wb.save(output_path)
# END_SAVE_TO_EXCEL
