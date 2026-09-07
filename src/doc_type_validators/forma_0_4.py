# START_MODULE_CONTRACT
# PURPOSE: Special-движок валидации формы 0.4 «О проекте в цифрах». 15 python-валидаторов (без LLM) над forma_0_4_main.json + раннер run_forma_0_4_special.
# INPUTS: args (target, parse_only?, rule_filter?, session_dir?); doc_configs/forma_0_4/{config.json, validation_rules.json}; данные из src/doc_type_parsers/forma_0_4.py.
# OUTPUTS: AuditResult; validation_report.xlsx в session_dir. Публичный символ: run_forma_0_4_special.
# KEYWORDS: forma_0_4, special-runner, python-validators, xlsx, no-llm, config-thresholds.
# LINKS: main.py (SPECIAL_ENGINE_RUNNERS), src/doc_type_parsers/forma_0_4.py (parse_forma_0_4), src/audit/models.py (AuditResult).
# RATIONALE: Логика 15 валидаторов перенесена из прод-движка audit_engine/forma_0_4 as-is (источник правды). Правила 7 (период) и 10 (прирост выработки) берут пороги из config (min_workdays, min_production_growth_pct) — поэтому валидаторы принимают config. Прод-subprocess консолидирован в прямые вызовы (стиль рефактора).
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import re
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl.utils import get_column_letter

from src.audit.models import AuditResult
from src.doc_type_parsers.forma_0_4 import parse_forma_0_4
# END_IMPORTS


# START_PATHS
_REPO_ROOT = Path(__file__).resolve().parents[2]
_DOC_CONFIGS_DIR = _REPO_ROOT / "doc_configs"
_LOGS_RESULT_DIR = _REPO_ROOT / "logs_result"
_DOC_TYPE = "forma_0_4"
# END_PATHS


# START_HELPERS
def _load_config() -> Dict[str, Any]:
    """Читает doc_configs/forma_0_4/config.json."""
    with open(_DOC_CONFIGS_DIR / _DOC_TYPE / "config.json", "r", encoding="utf-8") as f:
        return json.load(f)


def _load_rules() -> List[Dict[str, Any]]:
    """Загружает все правила из doc_configs/forma_0_4/validation_rules.json."""
    with open(_DOC_CONFIGS_DIR / _DOC_TYPE / "validation_rules.json", "r", encoding="utf-8") as f:
        return json.load(f)["rules"]


def _load_forma_data(parser_outputs_dir: Path) -> dict:
    """Загружает forma_0_4_main.json из директории парсера."""
    with open(parser_outputs_dir / "forma_0_4_main.json", "r", encoding="utf-8") as f:
        return json.load(f)


def _make_result(rule_index: str, rule_title: str, status: str, discrepancy: str = "") -> Dict[str, Any]:
    """Единый формат результата валидатора."""
    return {"rule_index": rule_index, "rule_title": rule_title, "status": status, "discrepancy": discrepancy}
# END_HELPERS


# ==============================================================================
# RULE 1 — Имя файла содержит '0.4' и тематические слова
# ==============================================================================
def _rule_1_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 1 — имя файла книги содержит код мероприятия «0.4» и тематические слова.

    Назначение: проверить, что имя xlsx-файла (data["meta"]["workbook"]) соответствует форме 0.4.
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result со статусом PASS/FAIL.
    Логика:
        1. Имя переводится в нижний регистр; если нет подстроки «0.4» — FAIL.
        2. Если нет ни одной из подстрок «о проекте в цифрах» / «опроектевцифрах» / «проекте в цифр» — FAIL.
        3. Иначе PASS.
    """
    title = rule["rule_title"]
    fname = data.get("meta", {}).get("workbook", "")
    fl = fname.lower()
    if "0.4" not in fl:
        return _make_result("1", title, "FAIL", f"Имя файла '{fname}' не содержит код мероприятия '0.4'.")
    keywords = ["о проекте в цифрах", "опроектевцифрах", "проекте в цифр"]
    if not any(k in fl for k in keywords):
        return _make_result("1", title, "FAIL",
                            f"Имя файла '{fname}' не содержит тематических слов 'О проекте в цифрах'.")
    return _make_result("1", title, "PASS")


# ==============================================================================
# RULE 2 — Шапка соглашения (G2:G4)
# ==============================================================================
_AGREEMENT_NO_RE = re.compile(r"\d{2,3}\s*[\-–—]\s*\d{2,3}\s*[\-–—]\s*\d{4}\s*/\s*ППТ", re.IGNORECASE)


def _rule_2_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 2 — шапка соглашения (G2:G4) заполнена и корректна по формату.

    Назначение: проверить три поля data["header"]: appendix_label (G2), agreement_no (G3), agreement_date_text (G4).
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result; при нарушениях — FAIL с перечнем проблем через «; ».
    Логика (каждая проверка добавляет issue, правило нарушено при любом issue):
        1. G2 пуст → issue; не содержит слова «приложение» → issue.
        2. G3 пуст → issue; не совпадает с _AGREEMENT_NO_RE (формат «XXX-XXX-YYYY/ППТ») → issue.
        3. G4 пуст → issue; не содержит 4 цифр подряд (года) → issue.
    """
    title = rule["rule_title"]
    h = data.get("header", {}) or {}
    issues = []
    appendix = (h.get("appendix_label") or "").strip()
    if not appendix:
        issues.append("G2: 'Приложение №...' не указано")
    elif "приложение" not in appendix.lower():
        issues.append(f"G2: ожидается 'Приложение №...', получено '{appendix}'")
    no = (h.get("agreement_no") or "").strip()
    if not no:
        issues.append("G3: номер соглашения не указан")
    elif not _AGREEMENT_NO_RE.search(no):
        issues.append(f"G3: формат номера ('{no}') не соответствует 'XXX-XXX-YYYY/ППТ'")
    date_text = (h.get("agreement_date_text") or "").strip()
    if not date_text:
        issues.append("G4: дата соглашения не указана")
    elif not re.search(r"\d{4}", date_text):
        issues.append(f"G4: дата '{date_text}' не содержит года (4 цифры)")
    if issues:
        return _make_result("2", title, "FAIL", "; ".join(issues))
    return _make_result("2", title, "PASS")


# ==============================================================================
# RULE 3 — Общая информация о потоке заполнена (B8, C8, D8, E8, F8, H8)
# ==============================================================================
def _rule_3_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 3 — общая информация о потоке заполнена (B8, C8, D8, E8, F8, H8).

    Назначение: проверить обязательные поля data["flow_info"].
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result; FAIL с перечнем пустых ячеек через «; », иначе PASS.
    Логика: issue добавляется, если пусто company (B8), region (C8), flow_name (D8),
        directions (F8), project_start_date_iso (H8), либо share_in_revenue (E8) равно None.
    """
    title = rule["rule_title"]
    fi = data.get("flow_info", {}) or {}
    issues = []
    if not fi.get("company"):
        issues.append("B8: предприятие не указано")
    if not fi.get("region"):
        issues.append("C8: регион не указан")
    if not fi.get("flow_name"):
        issues.append("D8: наименование потока не указано")
    if fi.get("share_in_revenue") is None:
        issues.append("E8: доля в выручке не указана")
    if not fi.get("directions"):
        issues.append("F8: основные направления оптимизации не указаны")
    if not fi.get("project_start_date_iso"):
        issues.append("H8: дата старта проекта не указана / невалидна")
    if issues:
        return _make_result("3", title, "FAIL", "; ".join(issues))
    return _make_result("3", title, "PASS")


# ==============================================================================
# RULE 4 — Дата старта проекта (H8) валидна
# ==============================================================================
def _rule_4_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 4 — дата старта проекта (H8) валидна.

    Назначение: проверить, что парсер распознал дату в H8 (data["flow_info"]["project_start_date_iso"] непусто).
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: FAIL, если ISO-дата отсутствует (ячейка пуста или не распарсилась в дд.мм.гггг), иначе PASS.
    """
    title = rule["rule_title"]
    iso = (data.get("flow_info", {}) or {}).get("project_start_date_iso")
    if not iso:
        return _make_result("4", title, "FAIL",
                            "Ячейка H8 (дата старта проекта) пуста или не парсится в формате дд.мм.гггг.")
    return _make_result("4", title, "PASS")


# ==============================================================================
# RULE 5 — Три показателя присутствуют + Выработка обязательна (ФЦК 0.4-3)
# ==============================================================================
def _rule_5_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 5 — у всех показателей указано наименование, среди них есть «Выработка» (критерий ФЦК 0.4-3).

    Назначение: проверить список data["indicators"].
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result; FAIL с перечнем проблем через «; », иначе PASS.
    Логика:
        1. Для каждого показателя с пустым name → issue с номером показателя и строкой листа.
        2. Наименования склеиваются в нижнем регистре; если нет подстроки «выработк» → issue.
    """
    title = rule["rule_title"]
    inds = data.get("indicators") or []
    issues = []
    for i, ind in enumerate(inds, start=1):
        if not ind.get("name"):
            issues.append(f"Показатель {i} (строка {ind.get('row')}): наименование не указано")
    names = " | ".join((ind.get("name") or "") for ind in inds).lower()
    if "выработк" not in names:
        issues.append("Главный критерий ФЦК 0.4-3: показатель 'Выработка' отсутствует среди наименований.")
    if issues:
        return _make_result("5", title, "FAIL", "; ".join(issues))
    return _make_result("5", title, "PASS")


# ==============================================================================
# RULE 6 — Единицы измерения соответствуют справочнику листа 3
# ==============================================================================
def _normalize(s: str) -> str:
    """Нормализует строку для сравнения: обрезка пробелов, нижний регистр, «ё» → «е»; None → ''."""
    return (s or "").strip().lower().replace("ё", "е")


def _category_for(name: str) -> Optional[str]:
    """Определяет категорию показателя по наименованию: 'time' / 'production' / 'stock' или None, если не распознано."""
    n = _normalize(name)
    if "врем" in n and "процесс" in n:
        return "time"
    if "выработк" in n or "производительност" in n:
        return "production"
    if "запас" in n or "незавершен" in n or "оборачиваемост" in n:
        return "stock"
    return None


def _rule_6_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 6 — единицы измерения показателей соответствуют справочнику листа 3.

    Назначение: сверить unit каждого показателя из data["indicators"] со списком допустимых
        единиц data["allowed_units"][категория].
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result; FAIL с перечнем проблем через «; », иначе PASS.
    Логика (для каждого показателя):
        1. Категория определяется через _category_for(name); если None — показатель пропускается.
        2. Справочник категории пуст → issue.
        3. Единица измерения пуста → issue.
        4. Единица (после _normalize) не совпадает ни с одной из справочника
           (равенство, startswith или вхождение справочной единицы в указанную) → issue.
    """
    title = rule["rule_title"]
    inds = data.get("indicators") or []
    allowed = data.get("allowed_units") or {}
    issues = []
    for ind in inds:
        name = ind.get("name") or ""
        unit = ind.get("unit") or ""
        cat = _category_for(name)
        if cat is None:
            continue
        cat_units = [_normalize(u) for u in (allowed.get(cat) or [])]
        if not cat_units:
            issues.append(f"Строка {ind.get('row')}: справочник для категории '{cat}' пуст в листе 3")
            continue
        unit_n = _normalize(unit)
        if not unit_n:
            issues.append(f"Строка {ind.get('row')}: единица измерения не указана для '{name}'")
            continue
        ok = any(unit_n == cu or unit_n.startswith(cu) or cu in unit_n for cu in cat_units)
        if not ok:
            issues.append(f"Строка {ind.get('row')} ('{name}'): единица '{unit}' не из справочника {allowed.get(cat)}")
    if issues:
        return _make_result("6", title, "FAIL", "; ".join(issues))
    return _make_result("6", title, "PASS")


# ==============================================================================
# RULE 7 — Период измерений ≥ min_workdays (config, дефолт 137 дней / 4.5 мес.)
# ==============================================================================
def _parse_iso(s: Optional[str]) -> Optional[date]:
    """Разбирает строку 'ГГГГ-ММ-ДД' в date; при невалидной строке или None возвращает None."""
    try:
        y, m, d = s.split("-")
        return date(int(y), int(m), int(d))
    except (ValueError, AttributeError):
        return None


def _rule_7_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 7 — период измерений (F10..G10) не короче порога min_workdays из config.

    Назначение: проверить даты начала/конца периода data["period_start_iso"] / data["period_end_iso"].
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title);
        config — config.json, используется ключ min_workdays (дефолт 137 дней ≈ 4.5 мес.).
    Выход: результат _make_result со статусом PASS/FAIL (первое найденное нарушение).
    Логика:
        1. F10 не парсится через _parse_iso → FAIL.
        2. G10 не парсится → FAIL.
        3. G10 <= F10 → FAIL.
        4. Разница в днях меньше min_workdays → FAIL (в сообщении — дни и примерно месяцы).
        5. Иначе PASS.
    """
    title = rule["rule_title"]
    min_days = config.get("min_workdays", 137)
    f10 = _parse_iso(data.get("period_start_iso"))
    g10 = _parse_iso(data.get("period_end_iso"))
    if not f10:
        return _make_result("7", title, "FAIL", "F10: дата начала периода не задана / не парсится.")
    if not g10:
        return _make_result("7", title, "FAIL", "G10: дата конца периода не задана / не парсится.")
    if g10 <= f10:
        return _make_result("7", title, "FAIL", f"G10 ({g10}) должна быть позже F10 ({f10}).")
    delta = (g10 - f10).days
    if delta < min_days:
        return _make_result("7", title, "FAIL",
                            f"Период {delta} дней (~{delta/30.4:.1f} мес.) меньше требуемых {min_days} дней (4.5 мес.).")
    return _make_result("7", title, "PASS")


# ==============================================================================
# RULE 8 — Все 6 числовых значений (F/G строк 11-13) — числа
# ==============================================================================
def _rule_8_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 8 — начальное и конечное значения каждого показателя (F/G строк 11-13) являются числами.

    Назначение: проверить value_start и value_end у каждого показателя data["indicators"].
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result; FAIL с перечнем нечисловых ячеек через «; », иначе PASS.
    Логика: value_start не int/float → issue «F{row}»; value_end не int/float → issue «G{row}».
    """
    title = rule["rule_title"]
    inds = data.get("indicators") or []
    issues = []
    for ind in inds:
        row = ind.get("row")
        if not isinstance(ind.get("value_start"), (int, float)):
            issues.append(f"F{row} ('{ind.get('name')}'): начальное значение не число")
        if not isinstance(ind.get("value_end"), (int, float)):
            issues.append(f"G{row} ('{ind.get('name')}'): конечное значение не число")
    if issues:
        return _make_result("8", title, "FAIL", "; ".join(issues))
    return _make_result("8", title, "PASS")


# ==============================================================================
# RULE 9 — Время протекания процесса убывает (G11 < F11)
# ==============================================================================
def _rule_9_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 9 — время протекания процесса сокращается (G11 < F11).

    Назначение: найти показатель «Время протекания процесса» и сравнить его начальное и конечное значения.
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result со статусом PASS/FAIL (первое найденное нарушение).
    Логика:
        1. Показатель ищется по наименованию, содержащему «врем» и «процесс»; не найден → FAIL.
        2. value_start или value_end не int/float → FAIL.
        3. value_end >= value_start (время не сократилось) → FAIL.
        4. Иначе PASS.
    """
    title = rule["rule_title"]
    inds = data.get("indicators") or []
    target = next(
        (i for i in inds if "врем" in (i.get("name") or "").lower() and "процесс" in (i.get("name") or "").lower()),
        None,
    )
    if not target:
        return _make_result("9", title, "FAIL", "Не найден показатель 'Время протекания процесса' в списке показателей.")
    fs, ge = target.get("value_start"), target.get("value_end")
    if not isinstance(fs, (int, float)) or not isinstance(ge, (int, float)):
        return _make_result("9", title, "FAIL", f"F{target['row']} или G{target['row']}: значения не числа.")
    if ge >= fs:
        return _make_result("9", title, "FAIL",
                            f"Время процесса должно сокращаться: F{target['row']}={fs}, G{target['row']}={ge} (G не меньше F).")
    return _make_result("9", title, "PASS")


# ==============================================================================
# RULE 10 — Прирост Выработки ≥ min_production_growth_pct (config, дефолт 52%)
# ==============================================================================
def _rule_10_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 10 — прирост Выработки не ниже порога min_production_growth_pct из config.

    Назначение: найти показатель «Выработка» и рассчитать относительный прирост (G - F) / F.
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title);
        config — config.json, используется ключ min_production_growth_pct (дефолт 52 %).
    Выход: результат _make_result со статусом PASS/FAIL (первое найденное нарушение).
    Логика:
        1. Показатель ищется по наименованию, содержащему «выработк» или «производительност»; не найден → FAIL.
        2. value_start или value_end не int/float → FAIL.
        3. value_start == 0 (деление невозможно) → FAIL.
        4. Прирост меньше порога (min_pct / 100) → FAIL с фактическим и требуемым процентом.
        5. Иначе PASS.
    """
    title = rule["rule_title"]
    min_pct = float(config.get("min_production_growth_pct", 52))
    threshold = min_pct / 100.0
    inds = data.get("indicators") or []
    target = next(
        (i for i in inds if "выработк" in (i.get("name") or "").lower() or "производительност" in (i.get("name") or "").lower()),
        None,
    )
    if not target:
        return _make_result("10", title, "FAIL", "Не найден показатель 'Выработка'.")
    fs, ge = target.get("value_start"), target.get("value_end")
    if not isinstance(fs, (int, float)) or not isinstance(ge, (int, float)):
        return _make_result("10", title, "FAIL", f"F{target['row']} или G{target['row']}: значения не числа.")
    if fs == 0:
        return _make_result("10", title, "FAIL", f"F{target['row']} = 0 — невозможно рассчитать прирост.")
    growth = (ge - fs) / fs
    if growth < threshold:
        return _make_result("10", title, "FAIL",
                            f"Прирост Выработки {growth*100:.2f}% < требуемых {min_pct:.0f}%. "
                            f"F{target['row']}={fs}, G{target['row']}={ge}.")
    return _make_result("10", title, "PASS")


# ==============================================================================
# RULE 11 — Запасы убывают (G13 < F13)
# ==============================================================================
def _rule_11_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 11 — запасы / незавершённое производство сокращаются (G13 < F13).

    Назначение: найти показатель запасов и сравнить его начальное и конечное значения.
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result со статусом PASS/FAIL (первое найденное нарушение).
    Логика:
        1. Показатель ищется по наименованию, содержащему «запас» или «незавершен»; не найден → FAIL.
        2. value_start или value_end не int/float → FAIL.
        3. value_end >= value_start (запасы не сократились) → FAIL.
        4. Иначе PASS.
    """
    title = rule["rule_title"]
    inds = data.get("indicators") or []
    target = next(
        (i for i in inds if "запас" in (i.get("name") or "").lower() or "незавершен" in (i.get("name") or "").lower()),
        None,
    )
    if not target:
        return _make_result("11", title, "FAIL", "Не найден показатель 'Запасы / Незавершённое производство'.")
    fs, ge = target.get("value_start"), target.get("value_end")
    if not isinstance(fs, (int, float)) or not isinstance(ge, (int, float)):
        return _make_result("11", title, "FAIL", f"F{target['row']} или G{target['row']}: значения не числа.")
    if ge >= fs:
        return _make_result("11", title, "FAIL",
                            f"Запасы должны сокращаться: F{target['row']}={fs}, G{target['row']}={ge} (G не меньше F).")
    return _make_result("11", title, "PASS")


# ==============================================================================
# RULE 12 — Все 6 формул на листе 2 заполнены (B2:B7)
# ==============================================================================
def _rule_12_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 12 — все 6 формул расчёта показателей на листе 2 (B2:B7) заполнены.

    Назначение: проверить список data["formulas_filled"] (row, filled).
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result со статусом PASS/FAIL (первое найденное нарушение).
    Логика:
        1. Список пуст (лист не найден) → FAIL.
        2. Есть записи с filled == False → FAIL с перечнем ячеек «B{row}».
        3. Записей меньше 6 → FAIL.
        4. Иначе PASS.
    """
    title = rule["rule_title"]
    formulas = data.get("formulas_filled") or []
    if not formulas:
        return _make_result("12", title, "FAIL", "Лист 'Формулы расчета показателей' не найден или пуст.")
    empty = [f"B{f['row']}" for f in formulas if not f.get("filled")]
    if empty:
        return _make_result("12", title, "FAIL", f"Не заполнены ячейки формул на листе 2: {', '.join(empty)}")
    if len(formulas) < 6:
        return _make_result("12", title, "FAIL", f"На листе 2 ожидалось 6 формул (B2:B7), найдено {len(formulas)}.")
    return _make_result("12", title, "PASS")


# ==============================================================================
# RULE 13 — Согласие на публикацию в D39
# ==============================================================================
def _rule_13_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 13 — согласие на публикацию (D39) выражено.

    Назначение: проверить текст data["signatures"]["consent"].
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result со статусом PASS/FAIL (первое найденное нарушение).
    Логика (текст в нижнем регистре без пробелов по краям):
        1. Пусто → FAIL.
        2. Нет подстроки «соглас» → FAIL.
        3. Есть подстрока «не соглас» (отрицание) → FAIL.
        4. Иначе PASS.
    """
    title = rule["rule_title"]
    consent = (data.get("signatures", {}) or {}).get("consent") or ""
    consent_l = consent.strip().lower()
    if not consent_l:
        return _make_result("13", title, "FAIL", "D39: согласие на публикацию не указано.")
    if "соглас" not in consent_l:
        return _make_result("13", title, "FAIL", f"D39: текст не содержит слова 'соглас(ен)': '{consent[:80]}'")
    if "не соглас" in consent_l:
        return _make_result("13", title, "FAIL", f"D39: указано отрицание: '{consent}'")
    return _make_result("13", title, "PASS")


# ==============================================================================
# RULE 14 — Дата подписи (G41) и дата документа (B46) заполнены
# ==============================================================================
_PLACEHOLDER_RE_14 = re.compile(r"_{3,}|«\s*___")


def _rule_14_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 14 — дата подписи (G41) и дата документа (B46) заполнены.

    Назначение: проверить поля data["signatures"]: signature_date_iso / signature_date_raw и doc_date_text.
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result; FAIL с перечнем проблем через «; », иначе PASS.
    Логика:
        1. G41: пусты и ISO-дата, и сырое значение → issue.
        2. B46: текст пуст → issue; текст совпадает с _PLACEHOLDER_RE_14 (подчёркивания «___», «« ___») → issue.
    """
    title = rule["rule_title"]
    sig = data.get("signatures", {}) or {}
    issues = []
    sig_iso = sig.get("signature_date_iso")
    sig_raw = sig.get("signature_date_raw")
    if not sig_iso and not sig_raw:
        issues.append("G41: дата подписи не заполнена")
    doc_text = (sig.get("doc_date_text") or "").strip()
    if not doc_text:
        issues.append("B46: дата документа не заполнена")
    elif _PLACEHOLDER_RE_14.search(doc_text):
        issues.append(f"B46: вместо даты документа плейсхолдер: '{doc_text}'")
    if issues:
        return _make_result("14", title, "FAIL", "; ".join(issues))
    return _make_result("14", title, "PASS")


# ==============================================================================
# RULE 15 — ФИО подписанта (B44) не плейсхолдер
# ==============================================================================
def _rule_15_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 15 — ФИО подписанта (B44) указано и не является плейсхолдером.

    Назначение: проверить data["signatures"]["signer_fio"].
    Вход: data — forma_0_4_main.json; rule — запись правила (rule_title); config — не используется.
    Выход: результат _make_result со статусом PASS/FAIL (первое найденное нарушение).
    Логика:
        1. Пусто → FAIL.
        2. После удаления символов «_», «/» и пробелов ничего не осталось (плейсхолдер) → FAIL.
        3. Очищенная строка короче 5 символов → FAIL.
        4. Иначе PASS.
    """
    title = rule["rule_title"]
    fio = ((data.get("signatures", {}) or {}).get("signer_fio") or "").strip()
    if not fio:
        return _make_result("15", title, "FAIL", "B44: ФИО подписанта не указано.")
    clean = re.sub(r"[_/\s]+", "", fio).strip()
    if not clean:
        return _make_result("15", title, "FAIL", f"B44: ФИО подписанта — плейсхолдер с подчёркиваниями: '{fio[:60]}'")
    if len(clean) < 5:
        return _make_result("15", title, "FAIL", f"B44: ФИО подписанта слишком короткое: '{fio}'")
    return _make_result("15", title, "PASS")


# START_VALIDATOR_DISPATCH
# Маппинг rule_index → функция-валидатор. Все валидаторы python (без LLM).
VALIDATOR_DISPATCH: Dict[str, Callable[[dict, dict, dict], dict]] = {
    "1": _rule_1_validate, "2": _rule_2_validate, "3": _rule_3_validate, "4": _rule_4_validate,
    "5": _rule_5_validate, "6": _rule_6_validate, "7": _rule_7_validate, "8": _rule_8_validate,
    "9": _rule_9_validate, "10": _rule_10_validate, "11": _rule_11_validate, "12": _rule_12_validate,
    "13": _rule_13_validate, "14": _rule_14_validate, "15": _rule_15_validate,
}
# END_VALIDATOR_DISPATCH


# START_REPORT
def _create_excel_report(results: List[Tuple[Dict, Dict, float]], report_path: Path) -> Any:
    """Собирает validation_report.xlsx из результатов валидаторов."""
    rows = []
    for rule, res, dur in results:
        rows.append({
            "rule_index": rule["rule_index"],
            "section": rule.get("section", ""),
            "rule_title": rule.get("rule_title", ""),
            "status": res.get("status", "UNKNOWN"),
            "discrepancy": res.get("discrepancy", ""),
            "duration_sec": round(dur, 2),
        })
    df = pd.DataFrame(rows)
    df["_sort"] = df["rule_index"].apply(lambda x: int(x) if str(x).isdigit() else 999)
    df = df.sort_values("_sort").drop(columns=["_sort"])
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Validation Results", index=False)
        ws = writer.sheets["Validation Results"]
        for idx, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).apply(len).max(), len(col))
            ws.column_dimensions[get_column_letter(idx + 1)].width = min(max_len + 2, 60)
    return df
# END_REPORT


# START_FORMA_0_4_RUNNER
def run_forma_0_4_special(args: Any) -> AuditResult:
    """
    Публичный entrypoint для forma_0_4. Регистрируется в SPECIAL_ENGINE_RUNNERS.
    Оркестрирует: 1 парсер → load rules → прогон 15 python-валидаторов → Excel.
    Правила 7/10 используют пороги из config (min_workdays, min_production_growth_pct).
    """
    start_time = time.time()
    target_path = Path(args.target)
    config = _load_config()

    session_dir = (
        Path(args.session_dir) if getattr(args, "session_dir", None)
        else _LOGS_RESULT_DIR / _DOC_TYPE / f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    )
    session_dir.mkdir(parents=True, exist_ok=True)
    parser_outputs_dir = session_dir / "parser_outputs"
    validation_outputs_dir = session_dir / "validation_outputs"
    validation_outputs_dir.mkdir(parents=True, exist_ok=True)

    # === Шаг 1: Парсинг XLSX ===
    parse_forma_0_4(target_path, parser_outputs_dir)

    if getattr(args, "parse_only", False):
        return AuditResult(
            doc_type=_DOC_TYPE, session_dir=session_dir, target_path=str(target_path),
            duration_sec=time.time() - start_time,
        )

    # === Шаг 2: Валидация ===
    data = _load_forma_data(parser_outputs_dir)
    rules = _load_rules()
    if getattr(args, "rule_filter", None):
        rules = [r for r in rules if r["rule_index"] == str(args.rule_filter)]
        if not rules:
            raise ValueError(f"Правило {args.rule_filter} не найдено в forma_0_4")

    results: List[Tuple[Dict, Dict, float]] = []
    for rule in rules:
        rule_index = rule["rule_index"]
        t0 = time.time()
        validate_fn = VALIDATOR_DISPATCH.get(rule_index)
        if validate_fn is None:
            res = _make_result(rule_index, rule.get("rule_title", ""), "MISSING",
                               f"Валидатор для правила {rule_index} не реализован")
        else:
            try:
                res = validate_fn(data, rule, config)
            except Exception as e:
                res = _make_result(rule_index, rule.get("rule_title", ""), "ERROR", f"Исключение: {e}")
        dur = time.time() - t0
        results.append((rule, res, dur))
        with open(validation_outputs_dir / f"validate_{rule_index}.json", "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=2)

    # === Шаг 3: Отчёт ===
    df = _create_excel_report(results, session_dir / "validation_report.xlsx")

    violations = []
    for rule, res, _ in results:
        if res.get("status") == "FAIL":
            violations.append({
                "rule_index": res.get("rule_index", rule.get("rule_index", "?")),
                "rule_title": res.get("rule_title", rule.get("rule_title", "")),
                "section": rule.get("section", ""),
                "discrepancy": res.get("discrepancy", ""),
            })

    return AuditResult(
        violations=violations, doc_type=_DOC_TYPE, session_dir=session_dir,
        duration_sec=time.time() - start_time, rules_checked=len(df), target_path=str(target_path),
    )
# END_FORMA_0_4_RUNNER
