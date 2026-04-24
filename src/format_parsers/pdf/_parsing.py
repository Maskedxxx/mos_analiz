# START_MODULE_CONTRACT
# PURPOSE: Чистая логика пост-обработки PDF-парсинга: reading order (xy-cut), парсинг таблиц из PaddleOCR-VL формата, сборка markdown-страницы, финальная нормализация. Без IO, без внешних сервисов.
# INPUTS: Списки регионов от layout-детектора, сырой текст таблиц от VLM, распознанные регионы с текстом, параметры пост-обработки из `PdfParsingConfig`.
# OUTPUTS: Отсортированные регионы, HTML/markdown таблицы, markdown-страницы, нормализованный текст.
# KEYWORDS: parsing, reading-order, xy-cut, table, markdown, normalizer, pure-logic.
# LINKS: src/format_parsers/pdf/_config.py, src/format_parsers/pdf/parse.py.
# RATIONALE: Клиенты (IO) и чистая логика (этот файл) разделены намеренно — функции здесь можно тестировать без сети, моделей и GPU.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import logging
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

from src.format_parsers.pdf._config import PdfParsingConfig
# END_IMPORTS


# START_LOGGERS
# PURPOSE: Отдельные логгеры по логическим подсистемам этого файла.
reading_order_logger = logging.getLogger(__name__ + ".reading_order")
table_parser_logger = logging.getLogger(__name__ + ".table")
normalizer_logger = logging.getLogger(__name__ + ".normalizer")
# END_LOGGERS


# START_READING_ORDER
# PURPOSE: Сортировка регионов по визуальному порядку чтения (xy-cut).
# INPUTS: Список регионов с полем `bbox` = [x1, y1, x2, y2].
# OUTPUTS: Тот же список, отсортированный — сверху вниз, слева направо, с учётом колонок.
# KEYWORDS: reading-order, xy-cut, recursive, layout.
def sort_by_reading_order(regions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Назначение:
        Сортирует регионы в порядке чтения через рекурсивный xy-cut.

    Вход:
        regions: Список `[{class_name, bbox, ...}, ...]` с координатами.

    Выход:
        Тот же список, переупорядоченный в порядке чтения.

    Логика:
        Делегирует `_xycut`, затем проецирует упорядоченные индексы обратно в регионы.
    """
    if len(regions) <= 1:
        return regions
    indices = list(range(len(regions)))
    ordered_indices = _xycut(indices, regions)
    result = [regions[i] for i in ordered_indices]
    reading_order_logger.info(f"Reading order: {len(result)} элементов отсортировано")
    return result


def _xycut(indices: List[int], regions: List[Dict[str, Any]]) -> List[int]:
    """
    Назначение:
        Рекурсивный xy-cut: чередует деление по Y и по X.

    Вход:
        indices: Индексы регионов для сортировки.
        regions: Все регионы (для доступа к bbox).

    Выход:
        Отсортированный список индексов.

    Логика:
        1. Проекция bbox на Y-ось → ищутся горизонтальные разрывы.
        2. Если найдены ≥2 группы — рекурсивно сортирует каждую, склеивает.
        3. Иначе проекция на X-ось → ищутся вертикальные разрывы (колонки).
        4. Если найдены ≥2 группы — рекурсивно и склеивание.
        5. Иначе сортировка базовым правилом (y, x).
    """
    if len(indices) <= 1:
        return indices
    boxes = [(i, regions[i]["bbox"]) for i in indices]
    y_groups = _split_axis(boxes, axis="y")
    if len(y_groups) > 1:
        result: List[int] = []
        for group in y_groups:
            group_indices = [idx for idx, _ in group]
            result.extend(_xycut(group_indices, regions))
        return result
    x_groups = _split_axis(boxes, axis="x")
    if len(x_groups) > 1:
        result = []
        for group in x_groups:
            group_indices = [idx for idx, _ in group]
            result.extend(_xycut(group_indices, regions))
        return result
    # Базовый случай: дальнейшие разрывы не найдены — сортируем визуально.
    boxes.sort(key=lambda b: (b[1][1], b[1][0]))
    return [idx for idx, _ in boxes]


def _split_axis(boxes: List[tuple], axis: str, gap_ratio: float = 0.02) -> List[List[tuple]]:
    """
    Назначение:
        Разбивает набор bbox по одной оси, если между проекциями есть существенный разрыв.

    Вход:
        boxes: Список `(index, [x1, y1, x2, y2])`.
        axis: `x` или `y` — вдоль какой оси искать разрывы.
        gap_ratio: Минимальный разрыв как доля от общего диапазона оси.

    Выход:
        Список групп. Каждая группа — список `(index, bbox)`.

    Логика:
        1. Проецирует bbox на ось и сортирует по началу проекции.
        2. Для каждого следующего bbox смотрит разрыв от текущего конца группы.
        3. Если разрыв > `gap_ratio * total_range` — начинает новую группу.
    """
    if axis == "y":
        projections = [(b[1][1], b[1][3], b) for b in boxes]
    else:
        projections = [(b[1][0], b[1][2], b) for b in boxes]
    projections.sort(key=lambda p: p[0])
    total_min = min(p[0] for p in projections)
    total_max = max(p[1] for p in projections)
    total_range = total_max - total_min
    if total_range <= 0:
        return [boxes]
    min_gap = total_range * gap_ratio
    groups: List[List[tuple]] = []
    current_group = [projections[0][2]]
    current_end = projections[0][1]
    for i in range(1, len(projections)):
        start_i = projections[i][0]
        end_i = projections[i][1]
        if start_i - current_end > min_gap:
            groups.append(current_group)
            current_group = [projections[i][2]]
            current_end = end_i
        else:
            current_group.append(projections[i][2])
            current_end = max(current_end, end_i)
    groups.append(current_group)
    return groups
# END_READING_ORDER


# START_TABLE_PARSER
# PURPOSE: Парсинг табличного формата PaddleOCR-VL (теги <fcel>, <ecel>, <lcel>, <ucel>, <nl>) в HTML/markdown таблицы с сохранением colspan/rowspan.
# INPUTS: Сырая строка с тегами.
# OUTPUTS: HTML- или markdown-таблица как строка.
# KEYWORDS: table, paddle-ocr-vl, colspan, rowspan, grid.
@dataclass
class Cell:
    """
    Назначение:
        Ячейка таблицы с атрибутами слияния и типом.

    Поля:
        text: Содержимое ячейки.
        cell_type: `filled` / `empty` / `lcel` (расширяет соседа слева) / `ucel`
            (расширяет ячейку сверху).
        colspan, rowspan: Размер слияния.
        covered: Флаг «ячейка поглощена соседней» — при сборке HTML не выводится.
    """

    text: str = ""
    cell_type: str = "empty"
    colspan: int = 1
    rowspan: int = 1
    covered: bool = False


def _tokenize_row(row_str: str) -> List[Cell]:
    """
    Назначение:
        Парсит одну строку PaddleOCR-VL в список `Cell` с сохранением типа.

    Вход:
        row_str: Строка вида `<fcel>Текст<ecel><lcel><fcel>Текст2`.

    Выход:
        Список `Cell` в порядке слева направо.

    Логика:
        Регулярка выделяет теги <fcel>/<ecel>/<lcel>/<ucel>, между тегами — текст ячейки.
        Первая встреченная пара (тег, текст) создаёт Cell; последующие добавляются по
        тому же принципу.
    """
    cells: List[Cell] = []
    tokens = re.split("(<fcel>|<ecel>|<lcel>|<ucel>)", row_str)
    current_type: Optional[str] = None
    current_text = ""
    for token in tokens:
        if token in ("<fcel>", "<ecel>", "<lcel>", "<ucel>"):
            if current_type is not None:
                cells.append(Cell(text=current_text.strip(), cell_type=current_type))
            type_map = {"<fcel>": "filled", "<ecel>": "empty", "<lcel>": "lcel", "<ucel>": "ucel"}
            current_type = type_map[token]
            current_text = ""
        else:
            current_text += token
    if current_type is not None:
        cells.append(Cell(text=current_text.strip(), cell_type=current_type))
    return cells


def _build_grid(raw: str) -> Optional[List[List[Cell]]]:
    """
    Назначение:
        Парсит сырой PaddleOCR-VL текст в сетку `Cell` с уже рассчитанными colspan/rowspan.

    Вход:
        raw: Сырой текст с тегами `<fcel>`, `<ecel>`, `<lcel>`, `<ucel>`, `<nl>`.

    Выход:
        Сетку `Cell` или `None`, если формат не распознан (нет ни `<fcel>`, ни `<ecel>`).

    Логика:
        1. Разбивает по `<nl>` на строки, каждую токенизирует `_tokenize_row`.
        2. Проход по строкам обрабатывает `lcel` — предыдущая ячейка получает `colspan + 1`,
           сама lcel помечается `covered`.
        3. Строит position_map: кто на какой позиции; `ucel` увеличивает `rowspan`
           источника сверху, сама помечается `covered`.
    """
    if "<fcel>" not in raw and "<ecel>" not in raw:
        return None
    rows_raw = re.split("<nl>", raw)
    rows_raw = [r.strip() for r in rows_raw if r.strip()]
    if not rows_raw:
        return None
    grid: List[List[Cell]] = []
    for row_str in rows_raw:
        cells = _tokenize_row(row_str)
        if cells:
            grid.append(cells)
    if not grid:
        return None
    # Проход 1: обработка lcel (расширение соседней слева ячейки).
    for row in grid:
        i = 0
        while i < len(row):
            if row[i].cell_type == "lcel":
                for j in range(i - 1, -1, -1):
                    if row[j].cell_type != "lcel":
                        row[j].colspan += 1
                        row[i].covered = True
                        break
            i += 1
    # Вычисляем логическое число колонок — максимум суммы colspan по не-covered ячейкам.
    max_logical_cols = 0
    for row in grid:
        width = sum(c.colspan for c in row if not c.covered)
        max_logical_cols = max(max_logical_cols, width)
    # Проход 2: position_map для обработки ucel (расширение ячейки сверху).
    position_map: List[List[Optional[Cell]]] = []
    for row_idx, row in enumerate(grid):
        pos_row: List[Optional[Cell]] = [None] * max_logical_cols
        col = 0
        for cell in row:
            if cell.covered:
                continue
            while col < max_logical_cols and pos_row[col] is not None:
                col += 1
            if col >= max_logical_cols:
                break
            if cell.cell_type == "ucel":
                cell.covered = True
                for prev_row_idx in range(row_idx - 1, -1, -1):
                    if prev_row_idx < len(position_map):
                        source = position_map[prev_row_idx][col]
                        if source is not None and not source.covered:
                            source.rowspan += 1
                            break
                        elif source is not None and source.covered:
                            break
                col += 1
            else:
                pos_row[col] = cell
                for cs in range(1, cell.colspan):
                    if col + cs < max_logical_cols:
                        pos_row[col + cs] = cell
                col += cell.colspan
        # Перенос rowspan в текущую строку: если в предыдущей строке ячейка имела
        # rowspan > 1 и здесь её не перекрыла новая, она «продолжается».
        if position_map:
            prev = position_map[-1]
            for c in range(max_logical_cols):
                if pos_row[c] is None and prev[c] is not None:
                    src = prev[c]
                    if not src.covered and src.rowspan > 1:
                        remaining = src.rowspan - (row_idx - _find_origin_row(position_map, src))
                        if remaining > 0:
                            pos_row[c] = src
        position_map.append(pos_row)
    n_rows = len(grid)
    n_cols = max_logical_cols
    table_parser_logger.info(f"Таблица: {n_rows} строк × {n_cols} колонок")
    return grid


def _find_origin_row(position_map: List[List[Optional[Cell]]], cell: Cell) -> int:
    """
    Назначение:
        Ищет номер строки, в которой ячейка впервые появилась.

    Вход:
        position_map: Текущий position_map до текущей строки.
        cell: Искомая ячейка.

    Выход:
        Индекс первой строки, где она встречается; 0 если не найдено.

    Логика:
        Линейный проход — используется при вычислении остатка rowspan.
    """
    for row_idx, row in enumerate(position_map):
        if cell in row:
            return row_idx
    return 0


def parse_paddle_table_html(raw: str) -> str:
    """
    Назначение:
        Конвертирует табличный формат PaddleOCR-VL в HTML-таблицу с colspan/rowspan.

    Вход:
        raw: Сырой текст с тегами.

    Выход:
        HTML `<table>...</table>`. Если формат не распознан — возвращает `raw` как есть.

    Логика:
        Строит сетку через `_build_grid`, затем собирает HTML с учётом `covered`, `colspan`,
        `rowspan`.
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


def parse_paddle_table_markdown(raw: str) -> str:
    """
    Назначение:
        Конвертирует табличный формат PaddleOCR-VL в markdown-таблицу.

    Вход:
        raw: Сырой текст с тегами.

    Выход:
        Markdown-таблица. Если формат не распознан — возвращает `raw` как есть.

    Логика:
        Markdown не поддерживает colspan/rowspan; сливающиеся ячейки становятся пустыми.
        Нормализует длину строк по максимальному числу колонок, склеивает `| ... |` формат.
    """
    grid = _build_grid(raw)
    if grid is None:
        return raw
    max_cols = 0
    text_rows: List[List[str]] = []
    for row in grid:
        texts: List[str] = []
        for cell in row:
            if cell.covered:
                texts.append("")
            else:
                texts.append(cell.text)
                for _ in range(cell.colspan - 1):
                    texts.append("")
        max_cols = max(max_cols, len(texts))
        text_rows.append(texts)
    for row in text_rows:
        while len(row) < max_cols:
            row.append("")
    lines: List[str] = []
    lines.append("| " + " | ".join(text_rows[0]) + " |")
    lines.append("|" + "|".join(["---"] * max_cols) + "|")
    for row in text_rows[1:]:
        lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def is_table_format(text: str) -> bool:
    """
    Назначение:
        Проверяет, содержит ли текст табличный формат PaddleOCR-VL.

    Вход:
        text: Текст от VLM.

    Выход:
        True, если в тексте есть хотя бы один из характерных тегов.

    Логика:
        Признак — наличие `<fcel>`, или `<ecel>+<nl>`, или `<lcel>`.
    """
    return "<fcel>" in text or ("<ecel>" in text and "<nl>" in text) or "<lcel>" in text
# END_TABLE_PARSER


# START_PAGE_MARKDOWN
# PURPOSE: Сборка markdown одной страницы из уже распознанных регионов, с учётом правил из `PdfParsingConfig`.
# INPUTS: Список распознанных регионов, `PdfParsingConfig`.
# OUTPUTS: Markdown-текст страницы (без HTML-таблиц, которые вставляются отдельно через `merge_text_and_tables`).
# KEYWORDS: page, markdown, class-wrapper, class-prompts.
def build_page_markdown(
    recognized_regions: List[Dict[str, Any]],
    parsing_cfg: PdfParsingConfig,
    page_num: int = 0,
) -> str:
    """
    Назначение:
        Собирает markdown-представление страницы из регионов, уже распознанных VLM.

    Вход:
        recognized_regions: Список `[{class_name, text, bbox, score}, ...]`, уже
            отсортированный в порядке чтения.
        parsing_cfg: Конфиг пост-обработки (промпты, markdown-обёртки, формат таблиц).
        page_num: Номер страницы (для логирования).

    Выход:
        Markdown-текст страницы.

    Логика:
        1. Итерирует регионы в заданном порядке.
        2. Пропускает регионы, для класса которых нет записи в `class_prompts`, и пустые.
        3. Таблицы (класс `Table` + табличный формат) конвертирует в HTML или markdown
           по `table_output_format`.
        4. Оборачивает текст по шаблону `class_md_wrapper[cls]` (fallback `{text}`).
        5. Склеивает блоки `\n\n`.
    """
    parts: List[str] = []
    for region in recognized_regions:
        cls = region["class_name"]
        text = region.get("text", "").strip()
        if parsing_cfg.class_prompts.get(cls) is None:
            normalizer_logger.debug(f"Пропуск: {cls} (нет промпта)")
            continue
        if not text:
            normalizer_logger.debug(f"Пропуск: {cls} (пустой текст)")
            continue
        if cls == "Table" and is_table_format(text):
            if parsing_cfg.table_output_format == "html":
                text = parse_paddle_table_html(text)
            else:
                text = parse_paddle_table_markdown(text)
        wrapper = parsing_cfg.class_md_wrapper.get(cls, "{text}")
        parts.append(wrapper.format(text=text))
    result = "\n\n".join(parts)
    normalizer_logger.info(f"Страница {page_num}: {len(parts)} блоков, {len(result)} символов markdown")
    return result


def merge_text_and_tables(
    full_page_text: str,
    table_entries: List[Dict[str, Any]],
    page_size: Tuple[int, int],
) -> str:
    """
    Назначение:
        Вставляет HTML-таблицы обратно в текст страницы по относительной y-позиции.

    Вход:
        full_page_text: Текст страницы от VLM (с замаскированными таблицами — таблиц там нет).
        table_entries: Список `[{bbox, html}, ...]`, отсортированный по y.
        page_size: Размер страницы в пикселях `(width, height)`.

    Выход:
        Строка с нормализованным текстом и вставленными таблицами.

    Логика:
        1. Если таблиц нет — просто нормализует и возвращает.
        2. Иначе для каждой таблицы вычисляет y-центр и относительное положение на странице.
        3. Переводит в номер строки в `full_page_text` и вставляет html-таблицу.
        4. Вставляет сзаду наперёд, чтобы индексы не сбивались.
    """
    if not table_entries:
        return normalize_paddle_text(full_page_text)
    _, page_h = page_size
    lines = full_page_text.split("\n")
    total_lines = len(lines)
    if total_lines == 0:
        parts = [t["html"] for t in table_entries]
        return "\n\n".join(parts)
    insertions: List[Tuple[int, str]] = []
    for entry in table_entries:
        y1 = entry["bbox"][1]
        y2 = entry["bbox"][3]
        y_center = (y1 + y2) / 2
        relative_pos = y_center / page_h
        insert_at = int(relative_pos * total_lines)
        insert_at = max(0, min(insert_at, total_lines))
        insertions.append((insert_at, entry["html"]))
    # Идём с конца к началу: так каждая вставка не ломает индексы для следующей.
    insertions.sort(key=lambda x: x[0], reverse=True)
    for line_idx, html in insertions:
        lines.insert(line_idx, f"\n{html}\n")
    return normalize_paddle_text("\n".join(lines))


def normalize_paddle_text(text: str) -> str:
    """
    Назначение:
        Финальная нормализация текста страницы после сборки.

    Вход:
        text: Сырой текст страницы (возможно, с HTML-таблицами внутри).

    Выход:
        Нормализованный текст.

    Логика:
        1. Сжимает три и более подряд идущих `\n` в `\n\n`.
        2. Убирает пробелы в конце строк.
        3. Обрезает внешние пробелы/переносы.

    Примечания:
        В отличие от OCR-нормализатора HunyuanOCR, здесь НЕ применяются
        unwrap_latex_artifacts и html_table_to_markdown: PaddleOCR-VL не генерирует
        LaTeX-артефакты, а HTML таблицы сохраняются с colspan/rowspan — они полезны
        для LLM-аудита.
    """
    if not text:
        return ""
    text = re.sub(r"\n{3,}", "\n\n", text)
    text = re.sub(r" +\n", "\n", text)
    return text.strip()
# END_PAGE_MARKDOWN
