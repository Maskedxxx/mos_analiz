# START_MODULE_CONTRACT
# PURPOSE: Парсеры doc_type «КПСЦ — Карта потока создания ценности» (xlsx). 9 независимых парсеров на 9 листов книги + общий `find_sheet` для нечёткого поиска листов (латиница⇄кириллица).
# INPUTS: Путь к xlsx и Workbook-объекты для каждого парсера; по конкретным листам читаются фиксированные ячейки и диапазоны.
# OUTPUTS: 9 структурированных словарей — по одному на парсер.
# KEYWORDS: kpsc, xlsx, structured-parsing, openpyxl.
# LINKS: main.py::KPSC_PARSER_MODULE_DISPATCH, main.py::kpsc_run_validations, doc_configs/kpsc/.
# RATIONALE: Каждый парсер знает свой лист КПСЦ и возвращает тот формат, который ожидают соответствующие валидаторы. Приватные хелперы оставлены с префиксами имени парсера, чтобы не конфликтовали между секциями одного файла.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import re
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional
from xml.etree import ElementTree

import openpyxl
import posixpath
import zipfile
from openpyxl import Workbook, load_workbook
from openpyxl.chart._chart import ChartBase
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries
from openpyxl.worksheet.worksheet import Worksheet
# END_IMPORTS

# START_SHEET_FINDER
# PURPOSE: Нечёткий поиск листа xlsx по ключевым словам. Нужен всем парсерам KPSC — в разных файлах названия листов отличаются по кейсу/транслитерации (латиница⇄кириллица).
# INPUTS: Workbook, список ключевых слов, опциональные exclude/prefer.
# OUTPUTS: Worksheet или None (find_sheet) / ValueError (find_sheet_or_raise).
# KEYWORDS: sheet-finder, fuzzy-match, kpsc.
kpsc_sheet_finder__LATIN_TO_CYRILLIC = str.maketrans({'A': 'А', 'B': 'В', 'C': 'С', 'E': 'Е', 'H': 'Н', 'K': 'К', 'M': 'М', 'O': 'О', 'P': 'Р', 'T': 'Т', 'X': 'Х', 'a': 'а', 'c': 'с', 'e': 'е', 'o': 'о', 'p': 'р', 'x': 'х'})

def kpsc_sheet_finder__normalize(name: str) -> str:
    """Нормализация имени листа: strip, lower, Latin→Cyrillic, убираем спец-символы."""
    name = name.strip().lower()
    name = name.translate(kpsc_sheet_finder__LATIN_TO_CYRILLIC)
    return name

def find_sheet(wb: Workbook, keywords: List[str], *, exclude_keywords: Optional[List[str]]=None, prefer_keywords: Optional[List[str]]=None) -> Optional[Worksheet]:
    """
    Нечёткий поиск листа в workbook по ключевым словам.

    Алгоритм:
    1. Нормализуем имена листов (strip, lower, Latin→Cyrillic)
    2. Ищем листы, содержащие ВСЕ keywords
    3. Исключаем листы с exclude_keywords
    4. Из оставшихся предпочитаем с prefer_keywords
    5. Из финальных кандидатов выбираем лист с максимумом данных

    Args:
        wb: openpyxl Workbook
        keywords: обязательные подстроки (нормализованные, lowercase)
        exclude_keywords: исключающие подстроки (если есть — лист отбрасывается)
        prefer_keywords: предпочтительные подстроки (приоритет при нескольких кандидатах)

    Returns:
        Worksheet или None, если не найден
    """
    if exclude_keywords is None:
        exclude_keywords = []
    if prefer_keywords is None:
        prefer_keywords = []
    kw_norm = [kpsc_sheet_finder__normalize(kw) for kw in keywords]
    excl_norm = [kpsc_sheet_finder__normalize(ek) for ek in exclude_keywords]
    pref_norm = [kpsc_sheet_finder__normalize(pk) for pk in prefer_keywords]
    candidates = []
    for sheet_name in wb.sheetnames:
        name_norm = kpsc_sheet_finder__normalize(sheet_name)
        if not all((kw in name_norm for kw in kw_norm)):
            continue
        if any((ek in name_norm for ek in excl_norm)):
            continue
        candidates.append(sheet_name)
    if not candidates:
        return None
    if len(candidates) == 1:
        return wb[candidates[0]]
    if pref_norm:
        preferred = []
        for sn in candidates:
            name_norm = kpsc_sheet_finder__normalize(sn)
            if any((pk in name_norm for pk in pref_norm)):
                preferred.append(sn)
        if preferred:
            candidates = preferred
    if len(candidates) == 1:
        return wb[candidates[0]]
    best_sheet = None
    best_count = -1
    for sn in candidates:
        ws = wb[sn]
        count = 0
        for row in ws.iter_rows(min_row=1, max_row=min(20, ws.max_row or 1)):
            for cell in row:
                if cell.value not in (None, ''):
                    count += 1
        if count > best_count:
            best_count = count
            best_sheet = sn
    return wb[best_sheet] if best_sheet else None

def find_sheet_or_raise(wb: Workbook, keywords: List[str], parser_name: str, **kwargs) -> Worksheet:
    """
    find_sheet() с выбросом исключения если лист не найден.

    Args:
        wb: openpyxl Workbook
        keywords: ключевые слова для поиска
        parser_name: имя парсера (для сообщения об ошибке)
        **kwargs: дополнительные аргументы для find_sheet()

    Returns:
        Worksheet

    Raises:
        ValueError: если лист не найден
    """
    ws = find_sheet(wb, keywords, **kwargs)
    if ws is None:
        raise ValueError(f'[{parser_name}] Лист не найден по keywords={keywords} среди {wb.sheetnames}')
    return ws
# END_SHEET_FINDER

# START_PARSE_KPSC_HEADER
# PURPOSE: Парсер шапки листа «КПСЦ» — название, компания, поток, ответственные, даты.
kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW = 15

def kpsc_parse_kpsc_header_collect_cells(ws, max_row: int=kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW, max_col: Optional[int]=None):
    """Собираем все непустые ячейки в заголовочном регионе."""
    if max_col is None:
        max_col = ws.max_column or 50
    cells = []
    comments = []
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if val not in (None, ''):
                cells.append({'coord': f'{get_column_letter(c)}{r}', 'row': r, 'col': c, 'value': val})
            if cell.comment:
                comments.append({'coord': f'{get_column_letter(c)}{r}', 'row': r, 'col': c, 'author': cell.comment.author, 'text': cell.comment.text})
    return (cells, comments)

def kpsc_parse_kpsc_header_collect_merged(ws, max_row: int=kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW, max_col: Optional[int]=None):
    """Собираем merged-диапазоны в заголовочном регионе."""
    if max_col is None:
        max_col = ws.max_column or 50
    merged = []
    for m in ws.merged_cells.ranges:
        if m.min_row <= max_row and m.min_col <= max_col:
            merged.append({'coord': m.coord, 'min_row': m.min_row, 'max_row': m.max_row, 'min_col': m.min_col, 'max_col': m.max_col})
    return merged

def kpsc_parse_kpsc_header__find_label_value(ws, label_keywords: List[str], max_row: int=kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW) -> Optional[Any]:
    """
    Динамический поиск значения по лейблу.

    Ищем ячейку, содержащую одно из label_keywords, затем берём значение
    из ближайшей непустой ячейки справа в той же строке.

    Если лейбл содержит значение inline (например "Наименование потока: Производство..."),
    извлекаем часть после двоеточия.
    """
    for r in range(1, max_row + 1):
        for c in range(1, min(5, (ws.max_column or 5) + 1)):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            v_lower = v.strip().lower()
            for kw in label_keywords:
                if kw not in v_lower:
                    continue
                if ':' in v:
                    parts = v.split(':', 1)
                    inline_val = parts[1].strip()
                    if inline_val:
                        return inline_val
                for vc in range(c + 1, min(c + 4, (ws.max_column or c) + 1)):
                    val = ws.cell(row=r, column=vc).value
                    if val not in (None, ''):
                        return val
                return None
    return None

def kpsc_parse_kpsc_header__find_title(ws, max_row: int=kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW) -> Optional[str]:
    """
    Извлекает заголовок карты потока.

    Ищем в первых строках длинную строку, содержащую ключевые слова
    типа "карта потока", "КПСЦ", "текущее состояние".
    """
    title_keywords = ['карта потока', 'кпсц', 'текущее состояние', 'наименование потока']
    for r in range(1, min(4, max_row + 1)):
        for c in range(1, 4):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and len(v.strip()) > 15:
                v_lower = v.strip().lower()
                if any((tk in v_lower for tk in title_keywords)):
                    return v.strip()
    for r in range(1, 3):
        for c in range(1, 4):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and len(v.strip()) > 15:
                return v.strip()
    return None

def kpsc_parse_kpsc_header__find_organization(ws, max_row: int=3) -> Optional[str]:
    """
    Ищем название организации (ООО/АО/ПАО + наименование) во всех ячейках первых строк.

    Некоторые компании (biznes_otel) размещают ООО в merged-ячейках далеко справа (col 45+),
    поэтому сканируем всю ширину строки.
    """
    import re
    org_pattern = re.compile('(ООО|АО|ПАО|ОАО|ЗАО)\\s*[«"\\\'"].+?[»"\\\'\\"]', re.IGNORECASE)
    max_col = ws.max_column or 50
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str):
                m = org_pattern.search(v)
                if m:
                    return m.group(0)
    return None

def kpsc_parse_kpsc_header_extract_fields(ws) -> Dict[str, Any]:
    """
    Динамическое извлечение полей заголовка КПСЦ.

    Вместо захардкоженных координат ищем лейблы по ключевым словам
    и берём значения из соседних ячеек.
    """
    return {'title': kpsc_parse_kpsc_header__find_title(ws), 'organization': kpsc_parse_kpsc_header__find_organization(ws), 'flow_name': kpsc_parse_kpsc_header__find_label_value(ws, ['поток:', 'наименование потока']), 'responsible': kpsc_parse_kpsc_header__find_label_value(ws, ['ответственн']), 'date_developed': kpsc_parse_kpsc_header__find_label_value(ws, ['дата разработ']), 'date_implementation': kpsc_parse_kpsc_header__find_label_value(ws, ['дата реализ', 'дата достиж']), 'compiled_by': kpsc_parse_kpsc_header__find_label_value(ws, ['составил', 'разработал']), 'takt_time': kpsc_parse_kpsc_header__find_label_value(ws, ['такт', 'время такта', 'takt'])}

def kpsc_parse_kpsc_header__find_takt_time_in_pokazateli(wb) -> Optional[Any]:
    """
    Fallback: ищет 'Время такта' на листе Показатели если не найдено в шапке КПСЦ.
    Сканирует весь лист, ищет строку-лейбл и берёт значение из соседней ячейки справа.
    """
    ws = find_sheet(wb, keywords=['показател'])
    if ws is None:
        return None
    for r in range(1, (ws.max_row or 50) + 1):
        for c in range(1, min(5, (ws.max_column or 5) + 1)):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            if 'время такта' in v.strip().lower() or 'такт' in v.strip().lower():
                for vc in range(c + 1, min(c + 4, (ws.max_column or c) + 1)):
                    val = ws.cell(row=r, column=vc).value
                    if val not in (None, ''):
                        return val
    return None

def kpsc_parse_kpsc_header_build_payload(xlsx: Path, sheet_name: Optional[str]=None):
    """Строит payload из данных header-блока КПСЦ."""
    wb = load_workbook(xlsx, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = find_sheet(wb, keywords=['кпсц'], exclude_keywords=['спагетти', 'укрупн', 'оцифровк'], prefer_keywords=['тс', 'текущ'])
        if ws is None:
            raise ValueError(f'Лист КПСЦ не найден среди {wb.sheetnames}')
    actual_sheet = ws.title
    max_col = ws.max_column or 50
    cells, _comments = kpsc_parse_kpsc_header_collect_cells(ws, max_col=max_col)
    fields = kpsc_parse_kpsc_header_extract_fields(ws)
    if not fields.get('takt_time'):
        fields['takt_time'] = kpsc_parse_kpsc_header__find_takt_time_in_pokazateli(wb)
    return {'meta': {'workbook': str(xlsx), 'sheet': actual_sheet, 'region': f'A1:{get_column_letter(max_col)}{kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW}'}, 'fields': fields}

def parse_kpsc_header(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит верхний блок КПСЦ и сохраняет результат в output_dir/kpsc_header_v2.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_kpsc_header_build_payload(xlsx_path)
    output_path = output_dir / 'kpsc_header_v2.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_kpsc_header_main():
    ap = argparse.ArgumentParser(description='Парсер верхнего блока КПСЦ')
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_kpsc_header_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)
# END_PARSE_KPSC_HEADER

# START_PARSE_KPSC_TABLE1
# PURPOSE: Парсер основной таблицы листа «КПСЦ».
def kpsc_parse_kpsc_table1_find_section_anchor(ws, phrase: str) -> Optional[Tuple[int, int]]:
    """
    Ищем начало секции таблицы КПСЦ с показателями потока.

    Стратегия (по приоритету):
    1. "1. Определение показателей потока" — стандартный формат (biznes_otel)
    2. "Название этапа процесса" — плоская таблица (mapper, sodex)
    3. "Показатель" рядом с "Ед. измерения" — сводная таблица (rotosnab, ruslet)
    """
    phrase_low = phrase.lower()
    for row in ws.iter_rows():
        for cell in row:
            val = cell.value
            if isinstance(val, str) and phrase_low in val.lower():
                if val.strip().startswith('1'):
                    return (cell.row, cell.column)
    for row in ws.iter_rows(max_row=min(15, ws.max_row)):
        for cell in row:
            val = cell.value
            if isinstance(val, str) and 'название этапа' in val.lower():
                return (max(1, cell.row - 1), cell.column)
    for r in range(1, ws.max_row + 1):
        row_texts = []
        first_col = None
        for c in range(1, min(10, (ws.max_column or 10) + 1)):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                row_texts.append(v.lower())
                if first_col is None:
                    first_col = c
        joined = ' '.join(row_texts)
        if 'показатель' in joined and 'ед.' in joined:
            return (max(1, r - 1), first_col or 1)
    return None

def kpsc_parse_kpsc_table1_find_header_row(ws, anchor_row: int) -> int:
    for r in range(anchor_row + 1, anchor_row + 5):
        if any((c.value not in (None, '') for c in ws[r])):
            return r
    return anchor_row + 1

def kpsc_parse_kpsc_table1_find_bottom_row(ws, header_row: int, section_col: int) -> int:
    """Ищем конец таблицы: первую строку, где в колонке section_col начинается следующая секция ("2.")."""
    r = header_row + 1
    while r <= ws.max_row:
        cell_val = ws.cell(row=r, column=section_col).value
        if isinstance(cell_val, str) and cell_val.strip().startswith('2.'):
            return r - 1
        row_vals = [c.value for c in ws[r]]
        if all((v in (None, '') for v in row_vals)) and r > header_row + 2:
            return r - 1
        r += 1
    return ws.max_row

def kpsc_parse_kpsc_table1_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    """
    Границы по строке заголовков:
      left  — первый непустой столбец,
      right — последний столбец в заголовке, где текст содержит 'ед. измерения'
              (регистр не важен); если не найдено, берем правый непустой.
    """
    cells = list(ws[header_row])
    left = None
    right = None
    right_by_text = None
    for c in cells:
        val = c.value
        if val not in (None, ''):
            if left is None:
                left = c.column
            if isinstance(val, str) and 'ед. измерения' in val.lower():
                right_by_text = c.column
            right = c.column
    if right_by_text:
        right = right_by_text
    if left is None:
        left = 1
    if right is None:
        right = left
    return (left, right)

def kpsc_parse_kpsc_table1_build_merged_lookup(ws):
    lookup = {}
    for merge in ws.merged_cells.ranges:
        coord = merge.coord
        min_row, min_col, max_row, max_col = (merge.min_row, merge.min_col, merge.max_row, merge.max_col)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                lookup[r, c] = coord
    return lookup

def kpsc_parse_kpsc_table1_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_kpsc_table1_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            cell = ws.cell(row=r, column=c)
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = cell.value
            row_cells.append({'coord': coord, 'col': c, 'row': r, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_kpsc_table1_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = find_sheet(wb, keywords=['кпсц'], exclude_keywords=['спагетти', 'укрупн', 'оцифровк'], prefer_keywords=['тс', 'текущ'])
        if ws is None:
            raise ValueError(f'Лист КПСЦ не найден среди {wb.sheetnames}')
    actual_sheet = ws.title
    anchor = kpsc_parse_kpsc_table1_find_section_anchor(ws, 'определение показателей потока')
    if not anchor:
        return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'section_title_cell': None}, 'bounds': None, 'rows': []}
    anchor_row, anchor_col = anchor
    header_row = kpsc_parse_kpsc_table1_find_header_row(ws, anchor_row)
    bottom_row = kpsc_parse_kpsc_table1_find_bottom_row(ws, header_row, anchor_col)
    left_col, right_col = kpsc_parse_kpsc_table1_compute_col_bounds(ws, header_row)
    table_rows = kpsc_parse_kpsc_table1_extract_table(ws, header_row, bottom_row, left_col, right_col)
    payload = {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'section_title_cell': f'{get_column_letter(anchor_col)}{anchor_row}'}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col, 'left_letter': get_column_letter(left_col), 'right_letter': get_column_letter(right_col), 'height': bottom_row - header_row + 1, 'width': right_col - left_col + 1}, 'rows': table_rows}
    return payload

def parse_kpsc_table1(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит таблицу '1. Определение показателей потока' и сохраняет в kpsc_table1_v2.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_kpsc_table1_build_payload(xlsx_path)
    output_path = output_dir / 'kpsc_table1_v2.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_kpsc_table1_main():
    parser = argparse.ArgumentParser(description="Парсер таблицы '1. Определение показателей потока'")
    parser.add_argument('-i', '--input', required=True, help='XLSX файл')
    parser.add_argument('-s', '--sheet', default=None, help='Лист (default: auto)')
    parser.add_argument('-o', '--output', default=None, help='JSON файл вывода (stdout если не указан)')
    args = parser.parse_args()
    payload = kpsc_parse_kpsc_table1_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

# END_PARSE_KPSC_TABLE1

# START_PARSE_LEGEND
# PURPOSE: Парсер листа «Легенда» — условные обозначения с картинками.
kpsc_parse_legend_NS = {'wb': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

def kpsc_parse_legend_read_sheet_drawing(xlsx: Path, sheet_name: str):
    with zipfile.ZipFile(xlsx) as z:
        wb_xml = ET.fromstring(z.read('xl/workbook.xml'))
        wb_rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rel_map = {rel.attrib['Id']: rel.attrib['Target'] for rel in wb_rels}
        sheet_target = None
        for sheet in wb_xml.find('wb:sheets', kpsc_parse_legend_NS):
            if sheet.attrib['name'] == sheet_name:
                rid = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
                sheet_target = rel_map[rid]
                break
        rel_path = f'xl/worksheets/_rels/{posixpath.basename(sheet_target)}.rels'
        rels = ET.fromstring(z.read(rel_path))
        drawing_path = None
        for rel in rels:
            if rel.attrib.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing':
                drawing_path = posixpath.normpath(posixpath.join('xl/worksheets', rel.attrib['Target']))
        d_rels_path = posixpath.join(posixpath.dirname(drawing_path), '_rels', posixpath.basename(drawing_path) + '.rels')
        d_rels = ET.fromstring(z.read(d_rels_path)) if d_rels_path in z.namelist() else None
        rel_pic = {rel.attrib['Id']: rel.attrib['Target'] for rel in d_rels} if d_rels is not None else {}
        drawing_root = ET.fromstring(z.read(drawing_path))
    return (drawing_root, rel_pic)

def kpsc_parse_legend_parse_pictures(drawing_root: ET.Element, rel_pic: Dict[str, str]):
    pics = []
    for anc in drawing_root.findall('./', kpsc_parse_legend_NS):
        pic_el = anc.find('xdr:pic', kpsc_parse_legend_NS)
        if pic_el is None:
            continue
        pf = anc.find('xdr:from', kpsc_parse_legend_NS)
        pt = anc.find('xdr:to', kpsc_parse_legend_NS)
        fcol = int(pf.find('xdr:col', kpsc_parse_legend_NS).text)
        frow = int(pf.find('xdr:row', kpsc_parse_legend_NS).text)
        tcol = int(pt.find('xdr:col', kpsc_parse_legend_NS).text)
        trow = int(pt.find('xdr:row', kpsc_parse_legend_NS).text)
        blip = pic_el.find('.//a:blip', kpsc_parse_legend_NS)
        rid = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed') if blip is not None else None
        pics.append({'row': frow + 1, 'col': fcol + 1, 'bbox': {'from': {'row': frow + 1, 'col': fcol + 1}, 'to': {'row': trow + 1, 'col': tcol + 1}}, 'rel_id': rid, 'target': rel_pic.get(rid)})
    return pics

def kpsc_parse_legend_collect_text(ws):
    texts = []
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=2).value
        if val not in (None, ''):
            texts.append({'row': row, 'col': 2, 'text': val})
    return texts

def kpsc_parse_legend_match_pics(texts: List[Dict[str, Any]], pics: List[Dict[str, Any]]):
    matched = []
    texts_order = [t for t in sorted(texts, key=lambda x: x['row']) if t['text'] != 'Расшифровка или пояснение']
    pics_order = sorted(pics, key=lambda x: x['row'])
    n = min(len(texts_order), len(pics_order))
    for i, t in enumerate(texts_order):
        pic = pics_order[i] if i < len(pics_order) else None
        matched.append({'row': t['row'], 'text': t['text'], 'picture': pic})
    return matched

def kpsc_parse_legend_build_payload(xlsx: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['условн', 'обозн'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx), 'sheet': None}, 'pictures': [], 'entries': []}
        actual_sheet = ws.title
    drawing_root, rel_pic = kpsc_parse_legend_read_sheet_drawing(xlsx, actual_sheet)
    pics = kpsc_parse_legend_parse_pictures(drawing_root, rel_pic)
    texts = kpsc_parse_legend_collect_text(ws)
    entries = kpsc_parse_legend_match_pics(texts, pics)
    return {'meta': {'workbook': str(xlsx), 'sheet': actual_sheet}, 'pictures': pics, 'entries': entries}

def parse_legend(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит лист 'Условные обозначения' и сохраняет в legend_v2.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_legend_build_payload(xlsx_path)
    output_path = output_dir / 'legend_v2.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    return payload

def kpsc_parse_legend_main():
    ap = argparse.ArgumentParser(description="Парсер листа 'Условные обозначения'")
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_legend_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)
# END_PARSE_LEGEND

# START_PARSE_LOSS_DIGITIZATION
# PURPOSE: Парсер листа «Оцифровка потерь» с таблицей потерь по операциям.
kpsc_parse_loss_digitization_HEADER_KEYS = ('описание проблемы', 'вид потери')

def kpsc_parse_loss_digitization_find_header_row(ws) -> Optional[int]:
    """Находим строку заголовков по ключевым фразам."""
    for r in range(1, ws.max_row + 1):
        lower_vals = [str(c.value).lower() for c in ws[r] if isinstance(c.value, str)]
        if lower_vals and all((any((key in v for v in lower_vals)) for key in kpsc_parse_loss_digitization_HEADER_KEYS)):
            return r
    return None

def kpsc_parse_loss_digitization_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    """Границы по непустым ячейкам строки заголовков (от первой до последней)."""
    left = None
    right = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v not in (None, ''):
            if left is None:
                left = c
            right = c
    if left is None:
        left = 1
        right = 1
    return (left, right)

def kpsc_parse_loss_digitization_find_bottom_row(ws, header_row: int, left_col: int, right_col: int) -> int:
    """
    Ищем конец таблицы: первая строка после заголовка, где данные отсутствуют
    во всех колонках кроме порядкового номера (left_col). Строки с одним
    номером считаем пустыми.
    """
    for r in range(header_row + 1, ws.max_row + 1):
        has_data = any((ws.cell(row=r, column=c).value not in (None, '') for c in range(left_col + 1, right_col + 1)))
        if not has_data:
            return r - 1
    return ws.max_row

def kpsc_parse_loss_digitization_build_merged_lookup(ws):
    lookup = {}
    for merge in ws.merged_cells.ranges:
        coord = merge.coord
        for r in range(merge.min_row, merge.max_row + 1):
            for c in range(merge.min_col, merge.max_col + 1):
                lookup[r, c] = coord
    return lookup

def kpsc_parse_loss_digitization_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_loss_digitization_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = ws.cell(row=r, column=c).value
            row_cells.append({'coord': coord, 'row': r, 'col': c, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_loss_digitization_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['оцифровк'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx_path), 'sheet': None, 'header_row': None}, 'bounds': None, 'rows': []}
        actual_sheet = ws.title
    header_row = kpsc_parse_loss_digitization_find_header_row(ws)
    if not header_row:
        return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': None}, 'bounds': None, 'rows': []}
    left_col, right_col = kpsc_parse_loss_digitization_compute_col_bounds(ws, header_row)
    bottom_row = kpsc_parse_loss_digitization_find_bottom_row(ws, header_row, left_col, right_col)
    rows = kpsc_parse_loss_digitization_extract_table(ws, header_row, bottom_row, left_col, right_col)
    return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': header_row}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col, 'left_letter': get_column_letter(left_col), 'right_letter': get_column_letter(right_col), 'height': bottom_row - header_row + 1, 'width': right_col - left_col + 1}, 'rows': rows}

def parse_loss_digitization(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит лист 'Оцифровка потерь КПСЦ' и сохраняет в ocifrovka_poteri_v2.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_loss_digitization_build_payload(xlsx_path)
    output_path = output_dir / 'ocifrovka_poteri_v2.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_loss_digitization_main():
    ap = argparse.ArgumentParser(description="Парсер листа 'Оцифровка потерь КПСЦ'")
    ap.add_argument('-i', '--input', required=True, help='Путь к XLSX файлу')
    ap.add_argument('-s', '--sheet', default=None, help='Имя листа (default: auto)')
    ap.add_argument('-o', '--output', help='JSON файл вывода (stdout если не указан)')
    args = ap.parse_args()
    payload = kpsc_parse_loss_digitization_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)
# END_PARSE_LOSS_DIG

# START_PARSE_PA1_CHART
# PURPOSE: Парсер графика на листе «ПА1».
def kpsc_parse_pa1_chart__sheet_name_from_range(rng: str) -> Tuple[str, str]:
    """Возвращает (sheet_name, range_part) из строки вида 'Лист'!$A$1:$B$2.

    Обрабатывает ссылки на внешние книги: '[2]ПА1 '!$A$1 → 'ПА1 '.
    """
    import re
    if '!' in rng:
        sheet, r = rng.split('!', 1)
        sheet = sheet.strip("'")
        sheet = re.sub('^\\[\\d+\\]', '', sheet)
        return (sheet, r)
    return ('', rng)

def kpsc_parse_pa1_chart__values_from_range(wb, sheet_name: str, rng: str):
    ws = wb[sheet_name]
    min_col, min_row, max_col, max_row = range_boundaries(rng)
    values = []
    for r in range(min_row, max_row + 1):
        row = []
        for c in range(min_col, max_col + 1):
            row.append(ws.cell(r, c).value)
        values.append(row)
    if min_row == max_row or min_col == max_col:
        return [v[0] if min_col == max_col else v for v in values] if min_row != max_row else values[0]
    return values

def kpsc_parse_pa1_chart__chart_title(chart: ChartBase) -> Optional[str]:
    t = chart.title
    if t is None:
        return None
    try:
        if t.tx and t.tx.rich and t.tx.rich.p:
            return ''.join((r.t for r in t.tx.rich.p[0].r))
    except Exception:
        pass
    return None

def kpsc_parse_pa1_chart_extract_chart_payload(chart: ChartBase, wb_data, wb_formulas) -> Dict[str, Any]:
    payload: Dict[str, Any] = {'type': type(chart).__name__, 'title': kpsc_parse_pa1_chart__chart_title(chart)}
    if getattr(chart, 'anchor', None) and getattr(chart.anchor, '_from', None):
        a_from = chart.anchor._from
        a_to = chart.anchor.to
        payload['anchor'] = {'from': {'col': a_from.col + 1, 'row': a_from.row + 1}, 'to': {'col': a_to.col + 1, 'row': a_to.row + 1}}
    series_list = []
    for s in chart.series:
        name = None
        if s.title:
            if getattr(s.title, 'v', None):
                name = s.title.v
            elif getattr(s.title, 'strRef', None) and s.title.strRef.strCache and s.title.strRef.strCache.pt:
                name = s.title.strRef.strCache.pt[0].v
        val_ref = getattr(getattr(s, 'val', None), 'numRef', None)
        val_range = val_ref.f if val_ref else None
        values = None
        if val_range:
            sheet_name, rng = kpsc_parse_pa1_chart__sheet_name_from_range(val_range)
            values = kpsc_parse_pa1_chart__values_from_range(wb_data, sheet_name, rng)
        cat_ref_obj = getattr(getattr(s, 'cat', None), 'strRef', None)
        cat_range = cat_ref_obj.f if cat_ref_obj else None
        categories = None
        if cat_range:
            sheet_name, rng = kpsc_parse_pa1_chart__sheet_name_from_range(cat_range)
            categories = kpsc_parse_pa1_chart__values_from_range(wb_data, sheet_name, rng)
        series_list.append({'name': name, 'values_range': val_range, 'values': values, 'categories_range': cat_range, 'categories': categories})
    payload['series'] = series_list
    return payload

def kpsc_parse_pa1_chart__find_pa1_sheet(wb) -> Optional[str]:
    """Находит лист ПА-1 через sheet_finder."""
    ws = find_sheet(wb, keywords=['па'], exclude_keywords=['спагетти', 'кпсц', 'ямадз'])
    return ws.title if ws else None

def kpsc_parse_pa1_chart_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb_formulas = load_workbook(xlsx_path, data_only=False)
    wb_data = load_workbook(xlsx_path, data_only=True)
    if not sheet_name:
        ws_found = find_sheet(wb_formulas, keywords=['па'], exclude_keywords=['спагетти', 'кпсц', 'ямадз'])
        if ws_found is None:
            return {'meta': {'workbook': str(xlsx_path), 'sheet': None}, 'charts': [], 'text_boxes': []}
        sheet_name = ws_found.title
    ws = wb_formulas[sheet_name]
    charts = getattr(ws, '_charts', [])
    charts_payload = [kpsc_parse_pa1_chart_extract_chart_payload(ch, wb_data, wb_formulas) for ch in charts]
    text_boxes: List[Dict[str, Any]] = []
    try:
        with zipfile.ZipFile(xlsx_path) as zf:
            import xml.etree.ElementTree as ET
            ns_wb = {'n': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
            wb_root = ET.fromstring(zf.read('xl/workbook.xml'))
            rid = None
            for sh in wb_root.findall('n:sheets/n:sheet', ns_wb):
                if sh.attrib.get('name') == sheet_name:
                    rid = sh.attrib[f'{{{ns_wb['r']}}}id']
                    break
            if rid:
                wb_rels = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
                sheet_target = None
                for rel in wb_rels:
                    if rel.attrib.get('Id') == rid:
                        sheet_target = rel.attrib['Target']
                        break
                if sheet_target:
                    sheet_rels_path = 'xl/' + sheet_target.replace('worksheets/', 'worksheets/_rels/') + '.rels'
                    if sheet_rels_path in zf.namelist():
                        sheet_rels = ET.fromstring(zf.read(sheet_rels_path))
                        drawing_target = None
                        for rel in sheet_rels:
                            if rel.attrib.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing':
                                drawing_target = rel.attrib['Target']
                                break
                        if drawing_target:
                            drawing_target = drawing_target.lstrip('../')
                            drawing_path = 'xl/' + drawing_target
                            if drawing_path in zf.namelist():
                                ns = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                                drw_root = ET.fromstring(zf.read(drawing_path))
                                for two in drw_root.findall('xdr:twoCellAnchor', ns):
                                    text_parts = [t.text or '' for t in two.findall('.//a:t', ns)]
                                    if text_parts:
                                        frm = two.find('xdr:from', ns)
                                        to = two.find('xdr:to', ns)
                                        text_boxes.append({'text': ''.join(text_parts).strip(), 'anchor': {'from': {'col': int(frm.find('xdr:col', ns).text) + 1, 'row': int(frm.find('xdr:row', ns).text) + 1}, 'to': {'col': int(to.find('xdr:col', ns).text) + 1, 'row': int(to.find('xdr:row', ns).text) + 1}}})
    except Exception:
        pass
    return {'meta': {'workbook': str(xlsx_path), 'sheet': sheet_name}, 'charts': charts_payload, 'text_boxes': text_boxes}

def parse_pa1_chart(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит диаграммы на листе 'ПА-1' и сохраняет в pa1_chart_v3.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_pa1_chart_build_payload(xlsx_path)
    output_path = output_dir / 'pa1_chart_v3.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_pa1_chart_main():
    ap = argparse.ArgumentParser(description="Парсер диаграмм на листе 'ПА-1' (после таблицы)")
    ap.add_argument('-i', '--input', required=True, help='Путь к XLSX')
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output', help='JSON вывод (stdout если не указан)')
    args = ap.parse_args()
    payload = kpsc_parse_pa1_chart_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

# END_PARSE_PA1_CHART

# START_PARSE_PA1_TABLE
# PURPOSE: Парсер таблицы на листе «ПА1».
def kpsc_parse_pa1_table_find_header_row(ws) -> Optional[int]:
    """Ищем строку, где в заголовках встречается слово 'итого'."""
    for r in range(1, ws.max_row + 1):
        if any((isinstance(c.value, str) and 'итого' in c.value.lower() for c in ws[r])):
            return r
    return None

def kpsc_parse_pa1_table_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    """
    Берём минимальный/максимальный столбцы с данными в строках заголовка
    и двух строках ниже (чтобы захватить пустой заголовок первого столбца,
    но заполненные значения 'Замер 1' и т.п.).
    """
    left = None
    right = 0
    last_row = min(ws.max_row, header_row + 2)
    for r in range(header_row, last_row + 1):
        for c in range(1, ws.max_column + 1):
            if ws.cell(row=r, column=c).value not in (None, ''):
                left = c if left is None else min(left, c)
                right = max(right, c)
    if left is None:
        left = 1
        right = 1
    return (left, right)

def kpsc_parse_pa1_table_find_bottom_row(ws, header_row: int, left_col: int, right_col: int) -> int:
    """Первая строка, полностью пустая в пределах таблицы, завершает данные."""
    r = header_row + 1
    while r <= ws.max_row:
        if all((ws.cell(row=r, column=c).value in (None, '') for c in range(left_col, right_col + 1))):
            return r - 1
        r += 1
    return ws.max_row

def kpsc_parse_pa1_table_build_merged_lookup(ws):
    lookup = {}
    for m in ws.merged_cells.ranges:
        coord = m.coord
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                lookup[r, c] = coord
    return lookup

def kpsc_parse_pa1_table_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_pa1_table_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = ws.cell(row=r, column=c).value
            row_cells.append({'coord': coord, 'row': r, 'col': c, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_pa1_table_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['па'], exclude_keywords=['спагетти', 'кпсц', 'ямадз'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx_path), 'sheet': None, 'header_row': None}, 'bounds': None, 'rows': []}
        actual_sheet = ws.title
    header_row = kpsc_parse_pa1_table_find_header_row(ws)
    if not header_row:
        return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': None}, 'bounds': None, 'rows': []}
    left_col, right_col = kpsc_parse_pa1_table_compute_col_bounds(ws, header_row)
    bottom_row = kpsc_parse_pa1_table_find_bottom_row(ws, header_row, left_col, right_col)
    rows = kpsc_parse_pa1_table_extract_table(ws, header_row, bottom_row, left_col, right_col)
    return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': header_row}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col, 'left_letter': get_column_letter(left_col), 'right_letter': get_column_letter(right_col), 'height': bottom_row - header_row + 1, 'width': right_col - left_col + 1}, 'rows': rows}

def parse_pa1_table(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит стартовую таблицу на листе 'ПА-1' и сохраняет в pa1_table_v1.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_pa1_table_build_payload(xlsx_path)
    output_path = output_dir / 'pa1_table_v1.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_pa1_table_main():
    ap = argparse.ArgumentParser(description="Парсер стартовой таблицы на листе 'ПА-1'")
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_pa1_table_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

# END_PARSE_PA1_TABLE

# START_PARSE_POKAZATELI
# PURPOSE: Парсер листа «Показатели» с ключевыми метриками потока.
kpsc_parse_pokazateli_TITLE_PHRASE = 'текущие показатели потока'

def kpsc_parse_pokazateli_find_title(ws) -> int:
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and kpsc_parse_pokazateli_TITLE_PHRASE in v.lower():
                return r
    raise ValueError("Title 'Текущие показатели потока' not found")

def kpsc_parse_pokazateli_find_header_row(ws, title_row: int) -> int:
    for r in range(title_row + 1, title_row + 5):
        if any((ws.cell(row=r, column=c).value not in (None, '') for c in range(1, ws.max_column + 1))):
            return r
    return title_row + 1

def kpsc_parse_pokazateli_find_bottom_row(ws, header_row: int) -> int:
    r = header_row + 1
    while r <= ws.max_row:
        if all((ws.cell(row=r, column=c).value in (None, '') for c in range(1, ws.max_column + 1))):
            return r - 1
        r += 1
    return ws.max_row

def kpsc_parse_pokazateli_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    """
    Граница по строке заголовков: берём первую непустую ячейку и продолжаем вправо,
    пока идут непустые. Если встречаем пустую колонку после начала таблицы — там обрываем.
    Это отсечёт служебные столбцы справа (I,J...).
    """
    left = None
    right = None
    started = False
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v not in (None, ''):
            if left is None:
                left = c
            right = c
            started = True
        elif started:
            break
    if left is None:
        left = 1
        right = 1
    return (left, right)

def kpsc_parse_pokazateli_build_merged_lookup(ws):
    lookup = {}
    for m in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = (m.min_col, m.min_row, m.max_col, m.max_row)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                lookup[r, c] = m.coord
    return lookup

def kpsc_parse_pokazateli_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_pokazateli_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = ws.cell(row=r, column=c).value
            row_cells.append({'coord': coord, 'row': r, 'col': c, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_pokazateli_build_payload(xlsx: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['показател'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx), 'sheet': None, 'title_row': None}, 'bounds': None, 'rows': []}
        actual_sheet = ws.title
    title_row = None
    try:
        title_row = kpsc_parse_pokazateli_find_title(ws)
    except ValueError:
        return {'meta': {'workbook': str(xlsx), 'sheet': actual_sheet, 'title_row': None}, 'bounds': None, 'rows': []}
    header_row = kpsc_parse_pokazateli_find_header_row(ws, title_row)
    bottom_row = kpsc_parse_pokazateli_find_bottom_row(ws, header_row)
    left_col, right_col = kpsc_parse_pokazateli_compute_col_bounds(ws, header_row)
    table_rows = kpsc_parse_pokazateli_extract_table(ws, header_row, bottom_row, left_col, right_col)
    return {'meta': {'workbook': str(xlsx), 'sheet': actual_sheet, 'title_row': title_row}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col}, 'rows': table_rows}

def parse_pokazateli(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит 'Текущие показатели потока' и сохраняет в pokazateli_v3.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_pokazateli_build_payload(xlsx_path)
    output_path = output_dir / 'pokazateli_v3.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_pokazateli_main():
    ap = argparse.ArgumentParser(description="Парсер 'Текущие показатели потока'")
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_pokazateli_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)
# END_PARSE_POKAZATELI

# START_PARSE_SPAGHETTI_PROBLEMS
# PURPOSE: Парсер листа «Спагетти-проблемы» — перечень найденных проблем по маршрутам.
kpsc_parse_spaghetti_problems_HEADER_KEY = 'описание проблемы'

def kpsc_parse_spaghetti_problems_find_header_row(ws) -> Optional[int]:
    for r in range(1, ws.max_row + 1):
        if any((isinstance(c.value, str) and kpsc_parse_spaghetti_problems_HEADER_KEY in c.value.lower() for c in ws[r])):
            return r
    return None

def kpsc_parse_spaghetti_problems_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    left = None
    right = 0
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v not in (None, ''):
            if left is None:
                left = c
            right = c
    if left is None:
        left = 1
        right = 1
    return (left, right)

def kpsc_parse_spaghetti_problems_find_bottom_row(ws, header_row: int, left_col: int, right_col: int) -> int:
    r = header_row + 1
    while r <= ws.max_row:
        if all((ws.cell(row=r, column=c).value in (None, '') for c in range(left_col, right_col + 1))):
            return r - 1
        r += 1
    return ws.max_row

def kpsc_parse_spaghetti_problems_build_merged_lookup(ws):
    lookup = {}
    for m in ws.merged_cells.ranges:
        coord = m.coord
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                lookup[r, c] = coord
    return lookup

def kpsc_parse_spaghetti_problems_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_spaghetti_problems_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = ws.cell(row=r, column=c).value
            row_cells.append({'coord': coord, 'row': r, 'col': c, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_spaghetti_problems_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['спагетти'], exclude_keywords=['диаграмм'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx_path), 'sheet': None, 'header_row': None}, 'bounds': None, 'rows': []}
        actual_sheet = ws.title
    header_row = kpsc_parse_spaghetti_problems_find_header_row(ws)
    if not header_row:
        return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': None}, 'bounds': None, 'rows': []}
    left_col, right_col = kpsc_parse_spaghetti_problems_compute_col_bounds(ws, header_row)
    bottom_row = kpsc_parse_spaghetti_problems_find_bottom_row(ws, header_row, left_col, right_col)
    rows = kpsc_parse_spaghetti_problems_extract_table(ws, header_row, bottom_row, left_col, right_col)
    return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': header_row}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col, 'left_letter': get_column_letter(left_col), 'right_letter': get_column_letter(right_col), 'height': bottom_row - header_row + 1, 'width': right_col - left_col + 1}, 'rows': rows}

def parse_spaghetti_problems(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит лист 'Перечень проблем по спагетти' и сохраняет в spaghetti_problems_v1.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_spaghetti_problems_build_payload(xlsx_path)
    output_path = output_dir / 'spaghetti_problems_v1.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_spaghetti_problems_main():
    ap = argparse.ArgumentParser(description="Парсер листа 'Перечень проблем по спагетти'")
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_spaghetti_problems_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)
# END_PARSE_SPAGHETTI_PROB

# START_PARSE_SPAGHETTI_SHEET
# PURPOSE: Парсер листа «Спагетти» — изображение маршрута + таблица шагов.
kpsc_parse_spaghetti_sheet_HEADER_PHRASES = ['шаги процесса', 'путь', 'перемещени']

def kpsc_parse_spaghetti_sheet_find_header_row(ws) -> Optional[int]:
    """Ищем строку-заголовок таблицы перемещений по ключевым словам."""
    for r in range(1, min(20, ws.max_row + 1)):
        row_vals = [c.value for c in ws[r]]
        for v in row_vals:
            if isinstance(v, str):
                v_low = v.lower()
                if any((phrase in v_low for phrase in kpsc_parse_spaghetti_sheet_HEADER_PHRASES)):
                    return r
    return None

def kpsc_parse_spaghetti_sheet_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    left = None
    right = 0
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v not in (None, ''):
            left = c
            break
    if left is None:
        left = 1
    for r in range(header_row, ws.max_row + 1):
        for c in range(left, ws.max_column + 1):
            if ws.cell(row=r, column=c).value not in (None, ''):
                right = max(right, c)
    if right < left:
        right = left
    return (left, right)

def kpsc_parse_spaghetti_sheet_find_bottom_row(ws, header_row: int, left_col: int, right_col: int) -> int:
    r = header_row + 1
    while r <= ws.max_row:
        if all((ws.cell(row=r, column=c).value in (None, '') for c in range(left_col, right_col + 1))):
            return r - 1
        r += 1
    return ws.max_row

def kpsc_parse_spaghetti_sheet_build_merged_lookup(ws):
    lookup = {}
    for m in ws.merged_cells.ranges:
        coord = m.coord
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                lookup[r, c] = coord
    return lookup

def kpsc_parse_spaghetti_sheet_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_spaghetti_sheet_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = ws.cell(row=r, column=c).value
            row_cells.append({'coord': coord, 'row': r, 'col': c, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_spaghetti_sheet_extract_pre_table(ws, header_row: int):
    cells = []
    for r in range(1, header_row):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v not in (None, ''):
                cells.append({'coord': f'{get_column_letter(c)}{r}', 'row': r, 'col': c, 'value': v})
    return cells

def kpsc_parse_spaghetti_sheet_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['спагетти'], exclude_keywords=['пробл', 'улучш', 'перечень'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx_path), 'sheet': None, 'header_row': None}, 'bounds': None, 'pre_table_cells': [], 'rows': []}
        actual_sheet = ws.title
    header_row = kpsc_parse_spaghetti_sheet_find_header_row(ws)
    if not header_row:
        return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': None}, 'bounds': None, 'pre_table_cells': [], 'rows': []}
    left_col, right_col = kpsc_parse_spaghetti_sheet_compute_col_bounds(ws, header_row)
    bottom_row = kpsc_parse_spaghetti_sheet_find_bottom_row(ws, header_row, left_col, right_col)
    pre_table = kpsc_parse_spaghetti_sheet_extract_pre_table(ws, header_row)
    table_rows = kpsc_parse_spaghetti_sheet_extract_table(ws, header_row, bottom_row, left_col, right_col)
    return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': header_row}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col, 'left_letter': get_column_letter(left_col), 'right_letter': get_column_letter(right_col), 'height': bottom_row - header_row + 1, 'width': right_col - left_col + 1}, 'pre_table_cells': pre_table, 'rows': table_rows}

def parse_spaghetti_sheet(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит лист 'Диаграмма Спагетти' и сохраняет в spaghetti_sheet_v2.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_spaghetti_sheet_build_payload(xlsx_path)
    output_path = output_dir / 'spaghetti_sheet_v2.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_spaghetti_sheet_main():
    ap = argparse.ArgumentParser(description="Парсер листа 'Диаграмма Спагетти' (без диаграммы)")
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_spaghetti_sheet_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)
# END_PARSE_SPAGHETTI_SHEET
