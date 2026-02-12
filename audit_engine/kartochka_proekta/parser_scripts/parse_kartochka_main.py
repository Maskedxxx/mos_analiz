#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Парсер листа «Карточка проекта» из XLSX-файла.

Извлекает: заголовок, секцию 1 (вовлечённые лица), секцию 2 (обоснование),
показатели (секция 3), события (секция 4), даты.
Результат → kartochka_main.json
"""

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import openpyxl


def _find_sheet(wb, keyword: str, exclude: str = "ШАБЛОН"):
    """Поиск листа по подстроке в имени, исключая шаблоны."""
    for name in wb.sheetnames:
        if keyword.lower() in name.lower() and exclude.upper() not in name.upper():
            return wb[name]
    # Fallback — второй лист (индекс 1)
    if len(wb.sheetnames) > 1:
        return wb[wb.sheetnames[1]]
    return wb[wb.sheetnames[0]]


def _cell_str(ws, coord: str) -> str:
    """Получить строковое значение ячейки, пустая → ''."""
    val = ws[coord].value
    if val is None:
        return ""
    return str(val).strip()


def _cell_value(ws, coord: str) -> Any:
    """Получить значение ячейки as-is (число, дата, строка, None)."""
    return ws[coord].value


def _format_date(val) -> Optional[str]:
    """Преобразование даты в ISO-строку."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s:
        return None
    return s


def _extract_fio(raw: str) -> str:
    """Извлечение ФИО из строки вида '____________ И.И. Иванов'."""
    # Убираем подчёркивания и лишние пробелы
    cleaned = re.sub(r'_+', '', raw).strip()
    return cleaned


def parse(xlsx_path: Path, output_dir: Path) -> dict:
    """
    Парсит лист 'Карточка проекта' из XLSX-файла.

    Ищет лист по подстроке 'Карточка проекта', исключая 'ШАБЛОН'.
    Возвращает и сохраняет структурированный JSON.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    ws = _find_sheet(wb, "Карточка проекта")

    result = {
        "meta": {
            "workbook": Path(xlsx_path).name,
            "sheet": ws.title,
            "max_row": ws.max_row,
        },
        "header": {},
        "section1": {},
        "section2": {},
        "indicators": [],
        "indicator_dates": {},
        "events": [],
    }

    # === Заголовок ===
    result["header"] = {
        "org_name": _cell_str(ws, "B2"),
        "project_name": _cell_str(ws, "B4"),
        "signee_position": _cell_str(ws, "K4"),
        "signee_name": _extract_fio(_cell_str(ws, "K7")),
        "signee_date": _cell_str(ws, "K8"),
        "has_utverzhday": "утверждаю" in _cell_str(ws, "K3").lower(),
    }

    # === Секция 1: Вовлечённые лица и рамки проекта ===
    # Извлекаем текст, убирая префиксы
    leader_raw = _cell_str(ws, "C15")
    leader = re.sub(r'^Руководитель\s+проекта\s*:\s*', '', leader_raw, flags=re.IGNORECASE).strip()

    team_raw = _cell_str(ws, "C16")
    team = re.sub(r'^Команда\s+проекта\s*:\s*', '', team_raw, flags=re.IGNORECASE).strip()

    result["section1"] = {
        "clients": _cell_str(ws, "E11"),
        "perimeter": _cell_str(ws, "E12"),
        "owner": _cell_str(ws, "E13"),
        "boundaries": _cell_str(ws, "E14"),
        "leader": leader,
        "team": team,
    }

    # === Секция 2: Обоснование ===
    result["section2"] = {
        "key_risk": _cell_str(ws, "M11"),
        "justification": _cell_str(ws, "M13"),
    }

    # === Показатели (секция 3) ===
    # Даты для База/Цель/Идеал
    result["indicator_dates"] = {
        "base_date": _format_date(_cell_value(ws, "F21")),
        "target_date": _format_date(_cell_value(ws, "G21")),
        "ideal_date": _format_date(_cell_value(ws, "H21")),
    }

    # Показатели — начинаем с C22, пока значение является числом
    for row in range(22, 40):
        c_val = _cell_value(ws, f"C{row}")
        # Проверяем что C содержит числовой номер показателя
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

        # Пропускаем пустые placeholder-строки (номер есть, но данных нет)
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

    # === События (секция 4) ===
    # Строки 20-31, столбцы K (название), Q (начало), S (окончание)
    for row in range(20, 40):
        k_val = _cell_str(ws, f"K{row}")
        if not k_val:
            # Прекращаем если строка полностью пуста
            q_check = _cell_value(ws, f"Q{row}")
            if q_check is None:
                continue
        if not k_val and _cell_value(ws, f"Q{row}") is None:
            continue

        q_val = _cell_value(ws, f"Q{row}")
        s_val = _cell_value(ws, f"S{row}")

        # Пропускаем строку-заголовок секции 4
        if k_val and "ключевые события" in k_val.lower():
            continue

        result["events"].append({
            "name": k_val.strip() if k_val else "",
            "start_date": _format_date(q_val),
            "end_date": _format_date(s_val),
        })

    wb.close()

    # Сохраняем результат
    output_file = output_dir / "kartochka_main.json"
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    return result
