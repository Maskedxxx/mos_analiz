# START_MODULE_CONTRACT
# PURPOSE: Парсеры doc_type «Карточка проекта» (xlsx). Три независимые функции, каждая парсит свой лист книги: основную карточку, методику расчёта показателей и справочник единиц измерения из выпадающего списка.
# INPUTS: Путь к .xlsx файлу карточки, путь к output_dir для сохранения промежуточных JSON.
# OUTPUTS: Три структурированных словаря (каждая функция — свой).
# KEYWORDS: kartochka-proekta, xlsx, structured-parsing, openpyxl.
# LINKS: main.py::KARTOCHKA_PARSER_MODULE_DISPATCH, main.py::kartochka_proekta_run_validations, doc_configs/kartochka_proekta/.
# RATIONALE: Три логически независимых парсера одного и того же файла. Каждый отвечает за свой лист и сохраняет JSON на диск для последующих валидаторов (в валидатор попадает именно распаршенный JSON).
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, TypedDict

import openpyxl
# END_IMPORTS


# START_CONTRACTS
# PURPOSE: TypedDict-контракты возвращаемых значений трёх парсеров. Показывают структуру верхнего уровня; вложенные объекты оставлены как `Dict[str, Any]` — подробности видны в docstring каждого парсера.
# INPUTS: —
# OUTPUTS: Контракты для IDE-автодополнения, mypy и документации.
# KEYWORDS: typeddict, contract, kartochka-proekta.
class DropdownDocument(TypedDict):
    """
    Результат `parse_dropdown` — справочник допустимых единиц измерения.

    Поля:
        categories: Имя категории (из заголовка колонки A1/B1/C1/D1) → список единиц из
            соответствующей колонки начиная со строки 2.
    """

    categories: Dict[str, List[str]]


class KartochkaMainDocument(TypedDict):
    """
    Результат `parse_kartochka_main` — основной лист «Карточка проекта».

    Поля:
        meta: `{workbook, sheet, max_row}` — имя книги, имя листа, кол-во строк.
        header: `{org_name, project_name, signee_position, signee_name, signee_date,
            has_utverzhday}` — шапка документа.
        section1: `{clients, perimeter, owner, boundaries, leader, team}` — секция 1.
        section2: `{key_risk, justification}` — секция 2.
        indicator_dates: `{base_date, target_date, ideal_date}` — даты замеров показателей.
        indicators: Список показателей, каждый — `{number, name, unit, base_value,
            target_value, ideal_value}`.
        events: Список событий, каждое — `{name, start_date, end_date}`.
    """

    meta: Dict[str, Any]
    header: Dict[str, Any]
    section1: Dict[str, Any]
    section2: Dict[str, Any]
    indicator_dates: Dict[str, Any]
    indicators: List[Dict[str, Any]]
    events: List[Dict[str, Any]]


class MetodikaDocument(TypedDict):
    """
    Результат `parse_metodika` — лист «Методика расчёта».

    Поля:
        meta: `{sheet}` — имя найденного листа; `None` если лист не найден.
        header: `{org_name, project_name}` — может быть пустым dict, если лист не найден.
        indicators: Список показателей, каждый — `{name, unit, calc_method, data_source,
            row_start}`.
    """

    meta: Dict[str, Any]
    header: Dict[str, Any]
    indicators: List[Dict[str, Any]]
# END_CONTRACTS


# START_SHARED_HELPERS
# PURPOSE: Общие утилиты для работы с ячейками и поиска листов, используемые всеми тремя парсерами.
# INPUTS: Объекты openpyxl (Workbook, Worksheet) и координаты.
# OUTPUTS: Stripped-строки, значения ячеек, листы или None.
# KEYWORDS: cell, sheet, helper.
def _find_sheet_by_keyword(wb: Any, keyword: str, exclude: Optional[str] = None) -> Any:
    """
    Назначение:
        Ищет лист в книге по подстроке в имени, опционально исключая шаблонные листы.

    Вход:
        wb: Workbook openpyxl.
        keyword: Подстрока для сравнения (регистр не важен).
        exclude: Если задан — листы, в имени которых есть эта подстрока (регистр не важен),
            пропускаются. Обычно используется для исключения `ШАБЛОН`.

    Выход:
        Worksheet openpyxl или `None`, если лист не найден.

    Логика:
        Линейный проход по `wb.sheetnames`. Первый совпавший — возвращается.
    """
    for name in wb.sheetnames:
        if keyword.lower() in name.lower():
            if exclude and exclude.upper() in name.upper():
                continue
            return wb[name]
    return None


def _cell_str(ws: Any, coord: str) -> str:
    """
    Назначение:
        Возвращает текст ячейки по excel-координате (например, `B2`).

    Вход:
        ws: Лист openpyxl.
        coord: Координата ячейки.

    Выход:
        Stripped-строка или пустая строка, если ячейка пустая.
    """
    val = ws[coord].value
    if val is None:
        return ""
    return str(val).strip()


def _cell_value(ws: Any, coord: str) -> Any:
    """
    Назначение:
        Возвращает значение ячейки «как есть» — без конвертации в строку.

    Вход:
        ws: Лист openpyxl.
        coord: Координата ячейки.

    Выход:
        Значение ячейки (число, дата, строка или `None`).

    Логика:
        Нужна в тех местах, где важно сохранить тип (даты, числа для пост-проверок).
    """
    return ws[coord].value


def _format_date(val: Any) -> Optional[str]:
    """
    Назначение:
        Преобразует дату/строку в ISO-строку `YYYY-MM-DD`.

    Вход:
        val: `datetime`, строка или `None`.

    Выход:
        ISO-строка, исходная строка (stripped) или `None`.

    Логика:
        1. `None` → `None`.
        2. `datetime` → strftime.
        3. Иначе — stripped str; пустая строка → `None`.
    """
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s:
        return None
    return s


def _extract_fio(raw: str) -> str:
    """
    Назначение:
        Вытаскивает ФИО из строки вида `____________ И.И. Иванов`.

    Вход:
        raw: Исходная строка.

    Выход:
        Та же строка без групп подчёркиваний, stripped.

    Логика:
        Удаляет подряд идущие `_` через regex.
    """
    return re.sub(r"_+", "", raw).strip()


def _is_field_label(text: str, label: str) -> bool:
    """
    Назначение:
        Проверяет, начинается ли текст с указанной метки поля (регистр не важен).

    Вход:
        text: Текст для проверки.
        label: Ожидаемая метка (например, `единицы измерения:`).

    Выход:
        True/False.
    """
    return text.lower().startswith(label.lower())
# END_SHARED_HELPERS


# START_PARSE_DROPDOWN
# PURPOSE: Парсит лист «выпадающий список» — справочник допустимых единиц измерения по категориям.
# INPUTS: Путь к xlsx, путь к output_dir.
# OUTPUTS: Dict `{categories: {header: [unit, ...], ...}}`. Сохраняет также `dropdown_units.json` в output_dir.
# KEYWORDS: dropdown, units, categories.
def parse_dropdown(xlsx_path: Path, output_dir: Path) -> DropdownDocument:
    """
    Назначение:
        Парсит лист с выпадающим списком единиц измерения.

    Вход:
        xlsx_path: Путь к xlsx карточки проекта.
        output_dir: Куда сохранить `dropdown_units.json`.

    Выход:
        Dict `{categories: Dict[str, List[str]]}`. Если лист не найден — `{categories: {}}`.

    Логика:
        1. Создаёт output_dir если нужно.
        2. Ищет лист по подстроке `выпадающий список`. Нет → сохраняет пустой JSON.
        3. Первая строка (A1..D1) — заголовки категорий; значения строк 2+ —
           единицы измерения соответствующей категории.
        4. Собирает словарь `header → [unit, ...]`, пустые значения пропускает.
        5. Сохраняет результат в `dropdown_units.json` и возвращает его.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    ws = _find_sheet_by_keyword(wb, "выпадающий список")

    if ws is None:
        wb.close()
        result: Dict[str, Any] = {"categories": {}}
        _save_json(result, output_dir / "dropdown_units.json")
        return result

    categories: Dict[str, List[str]] = {}
    columns = ["A", "B", "C", "D"]
    # Заголовки категорий — первая строка в каждой из 4 колонок.
    headers = [(_cell_str(ws, f"{col}1") or None) for col in columns]
    for col, header in zip(columns, headers):
        if header is None:
            continue
        units: List[str] = []
        for row in range(2, ws.max_row + 1):
            val = ws[f"{col}{row}"].value
            if val is not None:
                unit = str(val).strip()
                if unit:
                    units.append(unit)
        categories[header] = units

    wb.close()
    result = {"categories": categories}
    _save_json(result, output_dir / "dropdown_units.json")
    return result
# END_PARSE_DROPDOWN


# START_PARSE_KARTOCHKA_MAIN
# PURPOSE: Парсит основной лист «Карточка проекта»: шапку, секции 1 и 2, показатели, события.
# INPUTS: Путь к xlsx, путь к output_dir.
# OUTPUTS: Структурированный Dict с meta/header/section1/section2/indicators/indicator_dates/events.
# KEYWORDS: kartochka-main, header, indicators, events.
def parse_kartochka_main(xlsx_path: Path, output_dir: Path) -> KartochkaMainDocument:
    """
    Назначение:
        Парсит основной лист карточки проекта в структурированный словарь.

    Вход:
        xlsx_path: Путь к xlsx.
        output_dir: Куда сохранить `kartochka_main.json`.

    Выход:
        Dict с полями `meta`, `header`, `section1`, `section2`, `indicator_dates`,
        `indicators` (список), `events` (список).

    Логика:
        1. Ищет лист «Карточка проекта», исключая шаблоны.
        2. Читает фиксированные ячейки (B2, B4, K3/K4/K7/K8) → header.
        3. C15/C16 — руководитель/команда проекта; удаляет префиксы через regex.
        4. Фиксированные адреса E11..E14 → section1; M11/M13 → section2.
        5. F21/G21/H21 → indicator_dates (base/target/ideal).
        6. Строки 22..39, колонка C — номер показателя (int); D/E/F/G/H — name/unit/values.
           Цикл прерывается, когда C уже не int.
        7. Строки 20..39, колонка K — ключевое событие; Q/S — start/end date.
           Пропуск заголовка «Ключевые события».
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    ws = _find_sheet_by_keyword(wb, "Карточка проекта", exclude="ШАБЛОН")

    # Fallback: если явно не нашли — берём второй лист (если есть) или первый.
    if ws is None:
        if len(wb.sheetnames) > 1:
            ws = wb[wb.sheetnames[1]]
        else:
            ws = wb[wb.sheetnames[0]]

    result: Dict[str, Any] = {
        "meta": {"workbook": Path(xlsx_path).name, "sheet": ws.title, "max_row": ws.max_row},
        "header": {},
        "section1": {},
        "section2": {},
        "indicators": [],
        "indicator_dates": {},
        "events": [],
    }

    # Шапка: название организации, проекта, данные подписанта.
    result["header"] = {
        "org_name": _cell_str(ws, "B2"),
        "project_name": _cell_str(ws, "B4"),
        "signee_position": _cell_str(ws, "K4"),
        "signee_name": _extract_fio(_cell_str(ws, "K7")),
        "signee_date": _cell_str(ws, "K8"),
        "has_utverzhday": "утверждаю" in _cell_str(ws, "K3").lower(),
    }

    # Руководитель/команда проекта — в ячейках прописан префикс, который тут снимаем.
    leader_raw = _cell_str(ws, "C15")
    leader = re.sub(r"^Руководитель\s+проекта\s*:\s*", "", leader_raw, flags=re.IGNORECASE).strip()
    team_raw = _cell_str(ws, "C16")
    team = re.sub(r"^Команда\s+проекта\s*:\s*", "", team_raw, flags=re.IGNORECASE).strip()

    result["section1"] = {
        "clients": _cell_str(ws, "E11"),
        "perimeter": _cell_str(ws, "E12"),
        "owner": _cell_str(ws, "E13"),
        "boundaries": _cell_str(ws, "E14"),
        "leader": leader,
        "team": team,
    }
    result["section2"] = {
        "key_risk": _cell_str(ws, "M11"),
        "justification": _cell_str(ws, "M13"),
    }

    result["indicator_dates"] = {
        "base_date": _format_date(_cell_value(ws, "F21")),
        "target_date": _format_date(_cell_value(ws, "G21")),
        "ideal_date": _format_date(_cell_value(ws, "H21")),
    }

    # Показатели: строки 22..39, стоп — когда C уже не int.
    for row in range(22, 40):
        c_val = _cell_value(ws, f"C{row}")
        if c_val is None:
            break
        try:
            num = int(c_val)
        except (ValueError, TypeError):
            break
        d_val = _cell_str(ws, f"D{row}")
        e_val = _cell_str(ws, f"E{row}")
        f_val = _cell_value(ws, f"F{row}")
        g_val = _cell_value(ws, f"G{row}")
        h_val = _cell_value(ws, f"H{row}")
        # Пропускаем пустые строки в середине таблицы.
        if not d_val and not e_val and f_val is None:
            continue
        result["indicators"].append({
            "number": num,
            "name": d_val if d_val else None,
            "unit": e_val if e_val else None,
            "base_value": f_val,
            "target_value": g_val,
            "ideal_value": h_val,
        })

    # События: строки 20..39; пропускаем строку-заголовок «Ключевые события».
    for row in range(20, 40):
        k_val = _cell_str(ws, f"K{row}")
        q_val = _cell_value(ws, f"Q{row}")
        # Если нет ни имени события, ни даты начала — строка пустая, пропускаем.
        if not k_val and q_val is None:
            continue
        s_val = _cell_value(ws, f"S{row}")
        if k_val and "ключевые события" in k_val.lower():
            continue
        result["events"].append({
            "name": k_val.strip() if k_val else "",
            "start_date": _format_date(q_val),
            "end_date": _format_date(s_val),
        })

    wb.close()
    _save_json(result, output_dir / "kartochka_main.json")
    return result
# END_PARSE_KARTOCHKA_MAIN


# START_PARSE_METODIKA
# PURPOSE: Парсит лист «Методика расчёта» — для каждого показателя читает единицы, способ расчёта, источник данных.
# INPUTS: Путь к xlsx, путь к output_dir.
# OUTPUTS: Dict с `meta`, `header`, `indicators`.
# KEYWORDS: metodika, indicators, calc-method.
def parse_metodika(xlsx_path: Path, output_dir: Path) -> MetodikaDocument:
    """
    Назначение:
        Парсит лист «Методика расчёта» в структуру с показателями и их метаданными.

    Вход:
        xlsx_path: Путь к xlsx.
        output_dir: Куда сохранить `metodika.json`.

    Выход:
        Dict `{meta, header, indicators}`. Если лист не найден — пустой каркас.

    Логика:
        Лист имеет паттерн: название показателя в колонке B, далее три строки
        с метками `Единицы измерения:`, `Способ расчета:`, `Источник данных:` —
        значения этих полей лежат в колонке F.

        Алгоритм:
        1. Находит строку с названием показателя (текст в B, НЕ являющийся одной
           из меток полей и НЕ заголовок секции).
        2. Проверяет, что следующая строка — это метка `Единицы измерения:`.
        3. Если да — считывает три значения (F на +1, +2, +3 строках вниз) и
           добавляет запись в indicators.
        4. Перескакивает на 4 строки вперёд и продолжает.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    ws = _find_sheet_by_keyword(wb, "Методика расчет", exclude="ШАБЛОН")

    if ws is None:
        wb.close()
        result: Dict[str, Any] = {"meta": {"sheet": None}, "header": {}, "indicators": []}
        _save_json(result, output_dir / "metodika.json")
        return result

    result = {
        "meta": {"sheet": ws.title},
        "header": {
            "org_name": _cell_str(ws, "B2"),
            "project_name": _cell_str(ws, "B4"),
        },
        "indicators": [],
    }

    # Метки полей и заголовки, которые надо проскакивать в линейном сканировании.
    field_labels = ("единицы измерения", "способ расчет", "источник данных", "методика расчет")
    skip_keywords = ("карточка проекта", "методика расчет")

    row = 1
    max_row = ws.max_row
    while row <= max_row:
        b_val = _cell_str(ws, f"B{row}")
        if not b_val:
            row += 1
            continue

        b_lower = b_val.lower()
        is_label = any(b_lower.startswith(lbl) for lbl in field_labels)
        is_skip = any(kw in b_lower for kw in skip_keywords)
        if is_label or is_skip:
            row += 1
            continue

        # Следующая строка должна содержать метку «Единицы измерения:» — это подтверждает,
        # что мы стоим на названии показателя, а не на случайном тексте.
        next_b = _cell_str(ws, f"B{row + 1}") if row + 1 <= max_row else ""
        if _is_field_label(next_b, "единицы измерения"):
            indicator_name = re.sub(r":$", "", b_val).strip()
            unit = _cell_str(ws, f"F{row + 1}")
            calc_method = _cell_str(ws, f"F{row + 2}") if row + 2 <= max_row else ""
            data_source = _cell_str(ws, f"F{row + 3}") if row + 3 <= max_row else ""
            result["indicators"].append({
                "name": indicator_name,
                "unit": unit if unit else None,
                "calc_method": calc_method if calc_method else None,
                "data_source": data_source if data_source else None,
                "row_start": row,
            })
            row += 4
            continue

        row += 1

    wb.close()
    _save_json(result, output_dir / "metodika.json")
    return result
# END_PARSE_METODIKA


# START_IO_HELPERS
# PURPOSE: Точечные утилиты ввода-вывода, используемые парсерами выше.
# INPUTS: Словарь результата и путь к JSON-файлу.
# OUTPUTS: Файл на диске.
# KEYWORDS: json-save, io.
def _save_json(data: Dict[str, Any], output_path: Path) -> None:
    """
    Назначение:
        Сохраняет словарь в JSON-файл с ensure_ascii=False и отступами.

    Вход:
        data: Словарь для сериализации.
        output_path: Путь к выходному файлу.

    Выход:
        None.

    Логика:
        Используется всеми тремя парсерами для сохранения промежуточных JSON,
        которые потом читают валидаторы.
    """
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
# END_IO_HELPERS
