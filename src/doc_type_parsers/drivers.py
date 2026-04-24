# START_MODULE_CONTRACT
# PURPOSE: Парсер doc_type «Расчёт драйверов» (xlsx). Читает xlsx напрямую через zipfile + xml.etree (без openpyxl) для скорости и кастомной логики: собирает секции-опросники с ответами респондентов, итогами и средними.
# INPUTS: Путь к xlsx файлу, опциональное имя листа (по умолчанию — первый видимый).
# OUTPUTS: Dict `{meta, sections}` — структурированный представление опросника с секциями и вопросами.
# KEYWORDS: drivers, xlsx, zipfile, xml-parsing, questionnaire, structured-parsing.
# LINKS: main.py::drivers__run_pipeline (единственный внешний потребитель `parse_excel_to_json`).
# RATIONALE: Драйверы — это специфичный xlsx-опросник с фиксированной структурой «секция → вопросы → ответы респондентов». Парсер читает байты zip напрямую, что быстрее openpyxl для больших книг.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional
# END_IMPORTS


# START_XML_NAMESPACES
# PURPOSE: XML namespaces для разбора структуры xlsx (SpreadsheetML + Relationships).
MAIN_NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
RELS_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
# Регулярка разбирает cell-reference вида `A1`, `AB123` → колонка + номер строки.
CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")
# END_XML_NAMESPACES


# START_DATA_MODELS
# PURPOSE: Frozen dataclass'ы для строк, диапазонов слияний, схемы колонок и секций.
# INPUTS: —
# OUTPUTS: Типизированные структуры, используемые внутри парсера.
# KEYWORDS: dataclass, row, merge-range, column-layout, section.
@dataclass(frozen=True)
class Row:
    """
    Одна строка xlsx.

    Поля:
        index: 1-based номер строки.
        cells: `{column_letter → value}` — непустые ячейки.
    """

    index: int
    cells: Dict[str, str]

    def get(self, col: str, default: str = "") -> str:
        """Значение ячейки по букве колонки или `default`, если пусто."""
        return self.cells.get(col, default)

    @property
    def text(self) -> str:
        """Все непустые значения строки, склеенные пробелами в порядке колонок."""
        parts = [value.strip() for _col, value in sorted(self.cells.items()) if value and value.strip()]
        return " ".join(parts)


@dataclass(frozen=True)
class MergeRange:
    """Диапазон слияния ячеек (`A1:D3`).

    Поля:
        start_col, end_col: буквы колонок (inclusive).
        start_row, end_row: номера строк (inclusive).
    """

    start_col: str
    start_row: int
    end_col: str
    end_row: int

    @property
    def start_col_index(self) -> int:
        """1-based индекс начальной колонки."""
        return column_to_index(self.start_col)

    @property
    def end_col_index(self) -> int:
        """1-based индекс конечной колонки."""
        return column_to_index(self.end_col)


@dataclass(frozen=True)
class ColumnLayout:
    """
    Схема колонок для одной секции опросника (кто респондент, где итог и среднее).

    Поля:
        respondents: `{column_letter → имя респондента}` — основные колонки.
        total_column, total_label: колонка и имя для «Всего/Итог/Total».
        average_column, average_label: колонка и имя для «Среднее/Average».
    """

    respondents: Dict[str, str]
    total_column: Optional[str]
    total_label: Optional[str]
    average_column: Optional[str]
    average_label: Optional[str]


@dataclass
class Section:
    """
    Одна секция опросника (тематический блок вопросов).

    Поля:
        title: Заголовок секции (если найден).
        header_rows: Строки, составляющие шапку (для определения колонок-респондентов).
        questions: Список вопросов — `{row, number, question, problem_comment, scores, notes}`.
        summary: Итоговый комментарий секции (если есть).
        column_layout: Определяется лениво при первом вопросе. Хранит схему колонок.
    """

    title: Optional[str]
    header_rows: List[Row]
    questions: List[Dict[str, object]]
    summary: Optional[str]
    column_layout: Optional[ColumnLayout] = None
# END_DATA_MODELS


# START_COLUMN_HELPERS
# PURPOSE: Простые утилиты работы с xlsx‑адресами (буква ↔ индекс, разбор cell reference).
def column_to_index(col: str) -> int:
    """
    Буква колонки → 1-based индекс.

    Вход: `A` → 1, `Z` → 26, `AA` → 27.
    Выход: int.
    Бросает `ValueError`, если в строке есть символы вне `A-Z`.
    """
    col = col.upper()
    result = 0
    for ch in col:
        if not "A" <= ch <= "Z":
            raise ValueError(f"Invalid column letter: {col}")
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def split_cell_reference(ref: str) -> tuple[str, int]:
    """
    Разбор cell reference.

    Вход: `A1`, `AB123`.
    Выход: `("A", 1)`, `("AB", 123)`.
    Бросает `ValueError` при несовпадении с `CELL_REF_RE`.
    """
    match = CELL_REF_RE.match(ref)
    if not match:
        raise ValueError(f"Unexpected cell reference: {ref}")
    col, row = match.groups()
    return (col, int(row))
# END_COLUMN_HELPERS


# START_XLSX_READERS
# PURPOSE: Чтение сырых структур xlsx: sharedStrings, путь листа, строки с merge-ranges.
# INPUTS: Открытый `zipfile.ZipFile` xlsx-файла.
# OUTPUTS: Списки строк, merge-ranges, резолвленные XML-пути.
# KEYWORDS: xlsx, zipfile, shared-strings, sheet-path, rows.
def load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    """
    Загружает таблицу shared strings из `xl/sharedStrings.xml`.

    Вход: открытый zip.
    Выход: Список строк по индексам. Пустой список, если файла нет.
    """
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    strings: List[str] = []
    for si in root.findall("s:si", MAIN_NS):
        texts: List[str] = []
        for node in si.findall(".//s:t", MAIN_NS):
            texts.append(node.text or "")
        strings.append("".join(texts))
    return strings


def resolve_sheet_path(zf: zipfile.ZipFile, preferred_name: Optional[str] = None) -> str:
    """
    Резолвит внутрипуть листа в xlsx через `workbook.xml` + relationships.

    Вход:
        zf: открытый zip.
        preferred_name: если задано и такой лист существует — берём его; иначе первый
            не-hidden; иначе первый в списке.

    Выход: Путь вида `xl/worksheets/sheet1.xml`.
    """
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    sheets = workbook.find("s:sheets", MAIN_NS)
    if sheets is None:
        raise RuntimeError("Workbook does not contain <sheets>")
    candidates: List[tuple[str, str, Optional[str]]] = []
    for sheet in sheets.findall("s:sheet", MAIN_NS):
        name = sheet.get("name")
        rel_id = sheet.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        state = sheet.get("state")
        if not rel_id:
            continue
        candidates.append((name or "", rel_id, state))
    if not candidates:
        raise RuntimeError("No sheets found in workbook")
    target_rel_id: Optional[str] = None
    if preferred_name:
        for name, rel_id, state in candidates:
            if name == preferred_name:
                target_rel_id = rel_id
                break
    if target_rel_id is None:
        for name, rel_id, state in candidates:
            if state != "hidden":
                target_rel_id = rel_id
                break
    if target_rel_id is None:
        target_rel_id = candidates[0][1]
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    for rel in rels_root.findall("rel:Relationship", RELS_NS):
        if rel.get("Id") == target_rel_id:
            target = rel.get("Target")
            if not target:
                break
            if not target.startswith("/"):
                return f"xl/{target}"
            return target.lstrip("/")
    raise RuntimeError(f"Failed to resolve target for relationship {target_rel_id}")


def parse_cell_value(cell: ET.Element, shared_strings: List[str]) -> Optional[str]:
    """
    Достаёт текстовое значение ячейки из её XML-элемента.

    Типы ячеек:
        `s`         → индекс в shared_strings.
        `inlineStr` → все `<t>` внутри ячейки склеиваются.
        иначе       → значение `<v>` как есть.
    """
    cell_type = cell.get("t")
    if cell_type == "s":
        value_node = cell.find("s:v", MAIN_NS)
        if value_node is None or value_node.text is None:
            return ""
        idx = int(value_node.text)
        return shared_strings[idx] if 0 <= idx < len(shared_strings) else ""
    if cell_type == "inlineStr":
        texts = [node.text or "" for node in cell.findall(".//s:t", MAIN_NS)]
        return "".join(texts)
    value_node = cell.find("s:v", MAIN_NS)
    if value_node is None or value_node.text is None:
        return ""
    return value_node.text


def read_rows(
    zf: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: List[str],
) -> tuple[List[Row], List[MergeRange]]:
    """
    Читает все строки листа и его merge-ranges.

    Вход:
        zf: открытый zip.
        sheet_path: путь к xml листа (от `resolve_sheet_path`).
        shared_strings: таблица общих строк (от `load_shared_strings`).

    Выход:
        `(rows, merge_ranges)` — отсортированные строки и список merge-диапазонов.
    """
    sheet_xml = zf.read(sheet_path)
    sheet_root = ET.fromstring(sheet_xml)
    merge_ranges: List[MergeRange] = []
    merge_cells = sheet_root.find("s:mergeCells", MAIN_NS)
    if merge_cells is not None:
        for merge in merge_cells.findall("s:mergeCell", MAIN_NS):
            ref = merge.get("ref")
            if not ref:
                continue
            start, end = ref.split(":")
            start_col, start_row = split_cell_reference(start)
            end_col, end_row = split_cell_reference(end)
            merge_ranges.append(MergeRange(start_col, start_row, end_col, end_row))
    rows: List[Row] = []
    sheet_data = sheet_root.find("s:sheetData", MAIN_NS)
    if sheet_data is None:
        return (rows, merge_ranges)
    for row_node in sheet_data.findall("s:row", MAIN_NS):
        idx = int(row_node.get("r"))
        cells: Dict[str, str] = {}
        for cell in row_node.findall("s:c", MAIN_NS):
            ref = cell.get("r")
            if not ref:
                continue
            col, _ = split_cell_reference(ref)
            value = parse_cell_value(cell, shared_strings)
            if value is None:
                continue
            value = value.replace("\r", "").strip()
            if value:
                cells[col] = value
        if cells:
            rows.append(Row(idx, cells))
    rows.sort(key=lambda r: r.index)
    return (rows, merge_ranges)
# END_XLSX_READERS


# START_STRUCTURE_DETECTION
# PURPOSE: Эвристики распознавания структуры опросника — summary-строк, шапок, вопросов, схемы колонок.
# INPUTS: Списки строк/merges, отдельные значения.
# OUTPUTS: Признаки (bool/sets) и `ColumnLayout`.
# KEYWORDS: heuristics, summary-row, question-detection, column-layout.
def detect_summary_rows(merges: Iterable[MergeRange]) -> set[int]:
    """
    Находит строки-саммари: merge-range начинается в колонке `B` и тянется минимум до `J`.

    Вход: merge-ranges листа.
    Выход: Множество 1-based номеров строк, которые интерпретируем как summary.
    """
    summary_starts = set()
    for mr in merges:
        if mr.start_col.upper() == "B" and mr.end_col_index >= column_to_index("J"):
            summary_starts.add(mr.start_row)
    return summary_starts


def looks_like_question(cell_value: str) -> bool:
    """
    True, если значение ячейки начинается с цифры — типовой признак пронумерованного вопроса.
    """
    trimmed = cell_value.strip()
    if not trimmed:
        return False
    return trimmed[0].isdigit()


def is_header_marker(row: Row) -> bool:
    """
    Признак «строка — часть шапки секции»: в колонке B нет номера вопроса, но есть
    непустые ячейки в колонках E и правее (имена респондентов обычно там).
    """
    number_cell = (row.get("B") or "").strip()
    if number_cell and looks_like_question(number_cell):
        return False
    for col in row.cells:
        if column_to_index(col) >= column_to_index("E"):
            return True
    return False


def normalize_header_value(value: str) -> str:
    """Убирает NBSP, CR/LF, двоеточие на конце и схлопывает пробелы."""
    cleaned = value.replace("\xa0", " ").replace("\r", " ").replace("\n", " ")
    cleaned = cleaned.strip().rstrip(":")
    return " ".join(cleaned.split())


def is_numeric_text(text: str) -> bool:
    """True, если строка парсится как float (с запятой как разделителем)."""
    if not text:
        return False
    normalized = text.replace(",", ".")
    try:
        float(normalized)
        return True
    except ValueError:
        return False


def detect_column_layout(header_rows: List[Row]) -> ColumnLayout:
    """
    По строкам шапки определяет: кто респондент, где колонка «Всего», где «Среднее».

    Логика:
        1. Просматривает строки шапки, собирая первую непустую нечисловую метку в каждой
           колонке начиная с `E`.
        2. Метки с «балл» пропускаются (это заголовки шкалы, не респондент).
        3. Метки с `всего/итог/total/общее/суммар` → total_column.
        4. Метки с `средн/average/avg/mean` → average_column.
        5. Остальные → respondents (словарь `letter → имя`).
    """
    candidate_labels: Dict[str, str] = {}
    for row in header_rows:
        for col, raw_value in row.cells.items():
            if column_to_index(col) < column_to_index("E"):
                continue
            normalized = normalize_header_value(raw_value or "")
            if not normalized:
                continue
            lowered = normalized.lower().replace("ё", "е")
            if "балл" in lowered:
                continue
            if is_numeric_text(normalized):
                continue
            candidate_labels.setdefault(col, normalized)
    respondents: Dict[str, str] = {}
    total_column: Optional[str] = None
    total_label: Optional[str] = None
    average_column: Optional[str] = None
    average_label: Optional[str] = None
    for col, label in sorted(candidate_labels.items(), key=lambda item: column_to_index(item[0])):
        lowered = label.lower().replace("ё", "е")
        if any(token in lowered for token in ("всего", "итог", "итого", "overall", "total", "общее", "суммар")):
            total_column = col
            total_label = label
            continue
        if any(token in lowered for token in ("средн", "average", "avg", "mean")):
            average_column = col
            average_label = label
            continue
        respondents[col] = label
    return ColumnLayout(
        respondents=respondents,
        total_column=total_column,
        total_label=total_label,
        average_column=average_column,
        average_label=average_label,
    )
# END_STRUCTURE_DETECTION


# START_ASSEMBLE
# PURPOSE: Сборка распознанных элементов в итоговую структуру `{meta, sections}`.
# INPUTS: Список строк + summary-индексы.
# OUTPUTS: Dict с серализуемой структурой для сохранения в JSON.
# KEYWORDS: assemble, sections, questions, scores.
def assemble_structure(rows: List[Row], summary_rows: set[int]) -> Dict[str, object]:
    """
    Проходит по строкам, делит на секции и собирает вопросы.

    Логика по строке:
        - Если индекс строки в `summary_rows` → её текст идёт в `summary` текущей секции.
        - Если в E стоит «Баллы» → начинается новая секция.
        - Если секция ещё не начата → строка идёт в `meta`.
        - Если у секции ещё нет title и в B есть текст → это title.
        - Если строка похожа на шапку → добавляется в header_rows.
        - Если в B есть номер вопроса → это вопрос, scores берём из respondents‑колонок.
        - Иначе это комментарий/нота, добавляется к `notes` последнего вопроса.

    Выход:
        Dict `{meta, sections}`, где sections — список `{title, headers, questions, summary}`.
    """
    meta_info: List[Dict[str, object]] = []
    sections: List[Section] = []
    current_section: Optional[Section] = None
    last_question: Optional[Dict[str, object]] = None
    for row in rows:
        idx = row.index
        if idx in summary_rows:
            if current_section is not None:
                summary_text = row.get("B") or row.get("C") or row.get("D")
                if summary_text:
                    current_section.summary = summary_text.strip()
            continue
        if row.get("E") == "Баллы":
            if current_section is not None:
                sections.append(current_section)
            title = row.get("B") or None
            current_section = Section(
                title=title.strip() if title else None,
                header_rows=[row],
                questions=[],
                summary=None,
            )
            last_question = None
            continue
        if current_section is None:
            text = row.text
            if text:
                meta_info.append({"row": idx, "text": text})
            continue
        if current_section.title is None and row.get("B"):
            current_section.title = row.get("B").strip()
            current_section.header_rows.append(row)
            continue
        if is_header_marker(row):
            current_section.header_rows.append(row)
            continue
        number_cell = row.get("B")
        if number_cell and looks_like_question(number_cell):
            if current_section.column_layout is None:
                current_section.column_layout = detect_column_layout(current_section.header_rows)
            layout = current_section.column_layout
            scores: Dict[str, str] = {}
            for col, label in layout.respondents.items():
                scores[label] = row.get(col, "").strip()
            if layout.total_column:
                scores["Всего"] = row.get(layout.total_column, "").strip()
            if layout.average_column:
                scores["Средняя"] = row.get(layout.average_column, "").strip()
            question_entry = {
                "row": idx,
                "number": number_cell.strip(),
                "question": row.get("C").strip() if row.get("C") else "",
                "problem_comment": row.get("D").strip() if row.get("D") else "",
                "scores": scores,
                "notes": [],
            }
            current_section.questions.append(question_entry)
            last_question = question_entry
            continue
        note_text = row.get("C") or row.get("D")
        if note_text and last_question is not None:
            last_question["notes"].append({"row": idx, "text": note_text.strip()})
            continue
    if current_section is not None:
        sections.append(current_section)
    serialisable_sections: List[Dict[str, object]] = []
    for section in sections:
        header_texts = [row.text for row in section.header_rows if row.text]
        serialisable_sections.append({
            "title": section.title,
            "headers": header_texts,
            "questions": section.questions,
            "summary": section.summary,
        })
    return {"meta": meta_info, "sections": serialisable_sections}
# END_ASSEMBLE


# START_PUBLIC_ENTRY
# PURPOSE: Публичный entry пакета — единственное, что вызывается из main.py.
# INPUTS: Путь к xlsx, опциональное имя листа.
# OUTPUTS: Dict `{meta, sections}`.
# KEYWORDS: public-api, parse-excel.
def parse_excel_to_json(excel_path: Path, sheet_name: Optional[str] = None) -> Dict[str, object]:
    """
    Парсит xlsx-драйверов в структурированный словарь.

    Вход:
        excel_path: Путь к .xlsx файлу.
        sheet_name: Имя целевого листа. None → первый видимый.

    Выход:
        Dict `{meta: [...], sections: [...]}`.

    Ошибки:
        `ValueError` — если файл это временный файл Excel (`~$...`) или битый zip.
    """
    if excel_path.name.startswith("~$"):
        raise ValueError(f"Temporary Excel file detected: {excel_path.name}")
    try:
        with zipfile.ZipFile(excel_path) as zf:
            shared_strings = load_shared_strings(zf)
            sheet_path = resolve_sheet_path(zf, sheet_name)
            rows, merges = read_rows(zf, sheet_path, shared_strings)
    except zipfile.BadZipFile as exc:
        raise ValueError(f"{excel_path} is not a valid XLSX archive") from exc
    summary_rows = detect_summary_rows(merges)
    return assemble_structure(rows, summary_rows)
# END_PUBLIC_ENTRY
