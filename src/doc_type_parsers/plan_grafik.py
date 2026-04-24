# START_MODULE_CONTRACT
# PURPOSE: Парсер doc_type «План-график мероприятий» (xlsx). Читает лист «План мероприятий» и вытаскивает структурированные данные: утверждение, даты, шапку таблицы (2 строки заголовков), строки данных, столбцы статусов, формулы, подпись.
# INPUTS: Путь к .xlsx файлу плана-графика.
# OUTPUTS: `PlanGrafikDocument` — словарь с 11 структурированными полями.
# KEYWORDS: plan-grafik, xlsx, structured-parsing, openpyxl.
# LINKS: main.py (plan_grafik_run через SPECIAL_ENGINE_RUNNERS), doc_configs/plan_grafik/.
# RATIONALE: Парсер знает специфику xlsx-структуры плана-графика (конкретные адреса ячеек B8, DM8..DM12, I12/I13, шапка в 16-17 строках, данные с 18-й). Валидаторы плана-графика получают уже структурированный результат.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
from pathlib import Path
from typing import Any, Dict, List, TypedDict

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
# END_IMPORTS


# START_CONTRACT
# PURPOSE: Контракт возвращаемого значения парсера.
# INPUTS: —
# OUTPUTS: TypedDict для IDE-автодополнения и статической проверки.
# KEYWORDS: typeddict, contract, plan-grafik-document.
class PlanGrafikDocument(TypedDict, total=False):
    """
    Назначение:
        Структурированный результат парсинга xlsx плана-графика.

    Поля:
        filename: Имя исходного файла.
        sheets: Список листов книги.
        approval: Блок утверждения из колонки DM (8-12 строки).
        dates: Период действия плана-графика (строки 12-13, колонка I).
        title: Название документа из ячейки B8.
        headers_row16: Шапка таблицы, 1-й уровень заголовков (строка 16). Ключ — буква
            колонки, значение — текст заголовка.
        headers_row17: Шапка таблицы, 2-й уровень заголовков (строка 17).
        data_rows: Строки данных начиная с 18-й (problem, measure, responsible, ...).
        status_columns: Наличие и имена колонок статуса (DJ) и комментариев (DM).
        signature: Блок подписи внизу документа (последние 5 строк, фильтр по ключевым словам).
        formulas: Проверяемые формулы в ячейках DK17, DK18, DL18, DK12, N12 + признак наличия ошибки `#REF`.
        error: Если лист «План мероприятий» не найден — текст ошибки; иначе ключа нет.

    Логика:
        `total=False` — не все поля обязательны (в частности, `error` появляется только
        при отсутствии целевого листа).
    """

    filename: str
    sheets: List[str]
    approval: Dict[str, str]
    dates: Dict[str, Any]
    title: str
    headers_row16: Dict[str, str]
    headers_row17: Dict[str, str]
    data_rows: List[Dict[str, Any]]
    status_columns: Dict[str, Any]
    signature: str
    formulas: Dict[str, Dict[str, Any]]
    error: str
# END_CONTRACT


# START_HELPERS
# PURPOSE: Тонкие утилиты чтения ячейки по координатам/букве.
# INPUTS: Объект листа openpyxl и координаты.
# OUTPUTS: Текстовое значение ячейки (stripped) или пустая строка.
# KEYWORDS: cell, openpyxl, helper.
def _cell_str_by_letter(ws: Any, col_letter: str, row: int) -> str:
    """
    Назначение:
        Возвращает текст ячейки по букве колонки и номеру строки.

    Вход:
        ws: Лист openpyxl.
        col_letter: Буква колонки (например, `B`, `DM`).
        row: 1-based номер строки.

    Выход:
        Строка — stripped-значение ячейки, либо пустая строка если ячейка пустая.
    """
    col_idx = column_index_from_string(col_letter)
    val = ws.cell(row=row, column=col_idx).value
    return str(val).strip() if val else ""


def _cell_str_by_idx(ws: Any, col_idx: int, row: int) -> str:
    """
    Назначение:
        Возвращает текст ячейки по индексу колонки (1-based) и номеру строки.

    Вход:
        ws: Лист openpyxl.
        col_idx: 1-based индекс колонки.
        row: 1-based номер строки.

    Выход:
        Строка — stripped-значение ячейки, либо пустая строка если ячейка пустая.
    """
    val = ws.cell(row=row, column=col_idx).value
    return str(val).strip() if val else ""
# END_HELPERS


# START_PARSE_PLAN_GRAFIK
# PURPOSE: Основная функция парсинга xlsx плана-графика.
# INPUTS: Путь к xlsx.
# OUTPUTS: `PlanGrafikDocument`.
# KEYWORDS: parse, plan-grafik, xlsx, structured.
def parse_plan_grafik(file_path: str) -> PlanGrafikDocument:
    """
    Назначение:
        Парсит xlsx-план-график в структурированный `PlanGrafikDocument`.

    Вход:
        file_path: Путь к .xlsx.

    Выход:
        `PlanGrafikDocument` со всеми извлечёнными полями. Если лист «План мероприятий»
        не найден — возвращает результат с ключом `error` и без остальных данных.

    Логика:
        1. Открывает книгу **без** `data_only` (нужны исходные формулы для проверки).
        2. Если нет листа «План мероприятий» — возвращает `{..., error: ...}`.
        3. Читает фиксированные ячейки: `B8` (title), `DM8..DM12` (блок утверждения),
           `I12/I13` (даты действия).
        4. Сканирует строки 16-17 на шапку таблицы (2 уровня), колонки 1..159.
        5. Сканирует строки 18..min(max_row, 499) на строки данных — первые 7 колонок
           содержат problem_num/problem/measure/responsible/plan_fact/start_date/end_date.
           Полностью пустые строки пропускаются.
        6. Проверяет наличие колонок статуса (DJ) и комментариев (DM) в строке 16.
        7. Считывает формулы из фиксированных ячеек и определяет наличие `#REF`-ошибок.
        8. Ищет блок подписи в последних 5 строках по ключевым словам `подпись`, `___`, `202`.
    """
    wb = load_workbook(file_path, data_only=False)
    sheets = wb.sheetnames

    # Инициализируем контракт с пустыми полями — ниже заполняем то, что смогли извлечь.
    result: PlanGrafikDocument = {
        "filename": Path(file_path).name,
        "sheets": sheets,
        "approval": {},
        "dates": {},
        "title": "",
        "headers_row16": {},
        "headers_row17": {},
        "data_rows": [],
        "status_columns": {},
        "signature": "",
        "formulas": {},
    }

    if "План мероприятий" not in sheets:
        result["error"] = "Лист 'План мероприятий' не найден"
        return result

    ws = wb["План мероприятий"]

    # Фиксированные ячейки — адреса закреплены договорённостью о шаблоне плана-графика.
    result["title"] = _cell_str_by_letter(ws, "B", 8)

    dm_col = column_index_from_string("DM")
    result["approval"] = {
        "marker": _cell_str_by_idx(ws, dm_col, 8),
        "position": _cell_str_by_idx(ws, dm_col, 9),
        "company": _cell_str_by_idx(ws, dm_col, 10),
        "fio": _cell_str_by_idx(ws, dm_col, 11),
        "date_line": _cell_str_by_idx(ws, dm_col, 12),
    }

    i12 = ws.cell(row=12, column=column_index_from_string("I")).value
    i13 = ws.cell(row=13, column=column_index_from_string("I")).value
    result["dates"] = {
        "start": str(i12) if i12 else "",
        "end": str(i13) if i13 else "",
        "start_raw": i12,
        "end_raw": i13,
    }

    # Шапка таблицы: два уровня заголовков в строках 16 и 17. Колонок до 160.
    for col in range(1, 160):
        val = ws.cell(row=16, column=col).value
        if val and str(val).strip():
            result["headers_row16"][get_column_letter(col)] = str(val).strip()
    for col in range(1, 160):
        val = ws.cell(row=17, column=col).value
        if val and str(val).strip():
            result["headers_row17"][get_column_letter(col)] = str(val).strip()

    # Данные: строки с 18-й, лимит 500 (типовой план-график короче).
    max_row = ws.max_row
    for row in range(18, min(max_row + 1, 500)):
        a_val = ws.cell(row=row, column=1).value
        j_val = ws.cell(row=row, column=10).value
        i_val = ws.cell(row=row, column=9).value
        k_val = ws.cell(row=row, column=11).value
        l_val = ws.cell(row=row, column=12).value
        c_val = ws.cell(row=row, column=3).value
        d_val = ws.cell(row=row, column=4).value
        # Полностью пустые строки пропускаем — в плане могут быть разделители.
        if all(v is None for v in [a_val, j_val, i_val, k_val, l_val, c_val, d_val]):
            continue
        result["data_rows"].append({
            "row": row,
            "problem_num": str(a_val).strip() if a_val else "",
            "problem": str(c_val).strip() if c_val else "",
            "measure": str(d_val).strip() if d_val else "",
            "responsible": str(i_val).strip() if i_val else "",
            "plan_fact": str(j_val).strip() if j_val else "",
            "start_date": k_val,
            "end_date": l_val,
        })

    # Наличие колонок статуса и комментариев в шапке (строка 16).
    dj_col = column_index_from_string("DJ")
    dk_col = column_index_from_string("DK")
    dl_col = column_index_from_string("DL")
    dm_col_idx = column_index_from_string("DM")
    result["status_columns"] = {
        "status_header": _cell_str_by_idx(ws, dj_col, 16),
        "comments_header": _cell_str_by_idx(ws, dm_col_idx, 16),
        "status_present": bool(ws.cell(row=16, column=dj_col).value),
        "comments_present": bool(ws.cell(row=16, column=dm_col_idx).value),
    }

    # Проверяемые формулы — фиксированный набор ячеек, которые валидатор сверит
    # на наличие формулы (а не захардкоженного значения) и на отсутствие `#REF`.
    formulas: Dict[str, Dict[str, Any]] = {}
    dk17 = ws.cell(row=17, column=dk_col)
    formulas["DK17"] = {
        "value": str(dk17.value) if dk17.value else "",
        "is_formula": str(dk17.value).startswith("=") if dk17.value else False,
    }
    dk18 = ws.cell(row=18, column=dk_col)
    formulas["DK18"] = {
        "value": str(dk18.value) if dk18.value else "",
        "is_formula": str(dk18.value).startswith("=") if dk18.value else False,
    }
    dl18 = ws.cell(row=18, column=dl_col)
    formulas["DL18"] = {
        "value": str(dl18.value) if dl18.value else "",
        "is_formula": str(dl18.value).startswith("=") if dl18.value else False,
    }
    dk12 = ws.cell(row=12, column=dk_col)
    formulas["DK12"] = {
        "value": str(dk12.value) if dk12.value else "",
        "is_formula": str(dk12.value).startswith("=") if dk12.value else False,
        "has_error": "#REF" in str(dk12.value) if dk12.value else False,
    }
    n12 = ws.cell(row=12, column=column_index_from_string("N"))
    formulas["N12"] = {
        "value": str(n12.value) if n12.value else "",
        "is_formula": str(n12.value).startswith("=") if n12.value else False,
    }
    result["formulas"] = formulas

    # Блок подписи — последние 5 строк, фильтр по ключевым словам `подпись`/`___`/`202`.
    signature_parts: List[str] = []
    for row in range(max(1, max_row - 5), max_row + 1):
        for col in range(1, 160):
            val = ws.cell(row=row, column=col).value
            if val and str(val).strip() and (
                "подпись" in str(val).lower()
                or "___" in str(val)
                or "202" in str(val)
            ):
                signature_parts.append(f"{get_column_letter(col)}{row}: {str(val).strip()}")
    result["signature"] = "\n".join(signature_parts) + ("\n" if signature_parts else "")

    return result
# END_PARSE_PLAN_GRAFIK
