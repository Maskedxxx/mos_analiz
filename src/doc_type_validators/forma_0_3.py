# START_MODULE_CONTRACT
# PURPOSE: Special-движок валидации формы 0.3 «О предприятии в цифрах». 14 python-валидаторов (без LLM) над forma_0_3_main.json + раннер run_forma_0_3_special.
# INPUTS: args (target, parse_only?, rule_filter?, session_dir?); doc_configs/forma_0_3/{config.json, validation_rules.json}; данные из src/doc_type_parsers/forma_0_3.py.
# OUTPUTS: AuditResult; validation_report.xlsx в session_dir. Публичный символ: run_forma_0_3_special.
# KEYWORDS: forma_0_3, special-runner, python-validators, xlsx, no-llm.
# LINKS: main.py (SPECIAL_ENGINE_RUNNERS), src/doc_type_parsers/forma_0_3.py (parse_forma_0_3), src/audit/models.py (AuditResult).
# RATIONALE: Логика 14 валидаторов перенесена из прод-движка audit_engine/forma_0_3 as-is (источник правды). Прод запускал валидаторы как subprocess-скрипты; здесь они консолидированы в функции _rule_N_validate(data, rule) с диспатчем — единый файл, без подпроцессов (стиль рефактора, ср. kpsc.py / kartochka_proekta.py).
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import os
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Tuple

import pandas as pd
from openpyxl.utils import get_column_letter

from src.audit.models import AuditResult
from src.doc_type_parsers.forma_0_3 import parse_forma_0_3
# END_IMPORTS


# START_PATHS
_REPO_ROOT = Path(__file__).resolve().parents[2]
_DOC_CONFIGS_DIR = _REPO_ROOT / "doc_configs"
_LOGS_RESULT_DIR = _REPO_ROOT / "logs_result"
_DOC_TYPE = "forma_0_3"
# END_PATHS


# START_HELPERS
def _load_config() -> Dict[str, Any]:
    """Читает doc_configs/forma_0_3/config.json."""
    with open(_DOC_CONFIGS_DIR / _DOC_TYPE / "config.json", "r", encoding="utf-8") as f:
        return json.load(f)


def _load_rules() -> List[Dict[str, Any]]:
    """Загружает все правила из doc_configs/forma_0_3/validation_rules.json."""
    with open(_DOC_CONFIGS_DIR / _DOC_TYPE / "validation_rules.json", "r", encoding="utf-8") as f:
        return json.load(f)["rules"]


def _load_forma_data(parser_outputs_dir: Path) -> dict:
    """Загружает forma_0_3_main.json из директории парсера."""
    with open(parser_outputs_dir / "forma_0_3_main.json", "r", encoding="utf-8") as f:
        return json.load(f)


def _make_result(rule_index: str, rule_title: str, status: str, discrepancy: str = "") -> Dict[str, Any]:
    """Единый формат результата валидатора."""
    return {"rule_index": rule_index, "rule_title": rule_title, "status": status, "discrepancy": discrepancy}
# END_HELPERS


# ==============================================================================
# RULE 1 — Имя файла содержит '0.3' и тематические слова
# ==============================================================================
def _rule_1_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    fname = data.get("meta", {}).get("workbook", "")
    fl = fname.lower()
    if "0.3" not in fl:
        return _make_result("1", title, "FAIL", f"Имя файла '{fname}' не содержит код мероприятия '0.3'.")
    keywords = ["о предприятии в цифрах", "опредприятиивцифрах", "предприятии в цифр"]
    if not any(k in fl for k in keywords):
        return _make_result("1", title, "FAIL",
                            f"Имя файла '{fname}' не содержит тематических слов 'О предприятии в цифрах'.")
    return _make_result("1", title, "PASS")


# ==============================================================================
# RULE 2 — Структура разделов (заголовки B1, B6, B11)
# ==============================================================================
def _rule_2_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    h = data.get("headers", {}) or {}
    expected = {
        "title (B1)": ("Информация о показателях", h.get("title")),
        "section_general (B6)": ("Общая информация", h.get("section_general")),
        "section_indicators (B11)": ("Информация о целевых показателях", h.get("section_indicators")),
    }
    issues = []
    for label, (must_contain, actual) in expected.items():
        if not actual:
            issues.append(f"{label}: ячейка пуста")
        elif must_contain.lower() not in str(actual).lower():
            issues.append(f"{label}: ожидается фраза '{must_contain}', получено '{actual[:60]}'")
    if issues:
        return _make_result("2", title, "FAIL", "; ".join(issues))
    return _make_result("2", title, "PASS")


# ==============================================================================
# RULE 3 — Юр. форма + наименование в B9
# ==============================================================================
_LEGAL_FORMS = ["ООО", "ЗАО", "АО", "ПАО", "ИП", "ОАО", "ФГУП", "ГУП", "МУП", "АНО", "ОДО"]
_LEGAL_RE = re.compile(r"\b(" + "|".join(_LEGAL_FORMS) + r")\b")


def _rule_3_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    name = (data.get("general_info", {}).get("name") or "").strip()
    if not name:
        return _make_result("3", title, "FAIL", "Ячейка B9 пуста — наименование предприятия не указано.")
    if not _LEGAL_RE.search(name):
        return _make_result("3", title, "FAIL",
                            f"В B9 ('{name}') отсутствует юридическая форма (ООО/ЗАО/АО/ПАО/ИП/...).")
    rest = _LEGAL_RE.sub("", name).strip()
    if len(rest) < 2:
        return _make_result("3", title, "FAIL", f"В B9 ('{name}') есть юр. форма, но нет самого наименования.")
    return _make_result("3", title, "PASS")


# ==============================================================================
# RULE 4 — ИНН в C9 (10 цифр)
# ==============================================================================
_INN_RE = re.compile(r"^\d{10}$")


def _rule_4_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    inn = (data.get("general_info", {}).get("inn") or "").strip()
    if not inn:
        return _make_result("4", title, "FAIL", "Ячейка C9 пуста — ИНН не указан.")
    if inn.endswith(".0"):
        inn = inn[:-2]
    if not _INN_RE.match(inn):
        return _make_result("4", title, "FAIL", f"ИНН в C9 ('{inn}') должен содержать ровно 10 цифр (юрлицо).")
    return _make_result("4", title, "PASS")


# ==============================================================================
# RULE 5 — ОКВЭД-2 в D9 (формат '<код> <описание>', LEN > 10)
# ==============================================================================
_OKVED_CODE_RE = re.compile(r"^\d{2,3}(?:\.\d{1,2}){1,2}\b")


def _rule_5_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    okved = (data.get("general_info", {}).get("okved") or "").strip()
    if not okved:
        return _make_result("5", title, "FAIL", "Ячейка D9 пуста — ОКВЭД не указан.")
    if len(okved) <= 10:
        return _make_result("5", title, "FAIL",
                            f"ОКВЭД в D9 ('{okved}') слишком короткий (<= 10 символов): должен содержать код + описание.")
    if not _OKVED_CODE_RE.match(okved):
        return _make_result("5", title, "FAIL",
                            f"ОКВЭД в D9 ('{okved[:60]}...') не начинается с кода формата 'NN.NN' или 'NN.NN.NN'.")
    return _make_result("5", title, "PASS")


# ==============================================================================
# RULE 6 — Регион в E9 (не пустой)
# ==============================================================================
def _rule_6_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    region = (data.get("general_info", {}).get("region") or "").strip()
    if not region:
        return _make_result("6", title, "FAIL", "Ячейка E9 пуста — регион не указан.")
    if len(region) < 2:
        return _make_result("6", title, "FAIL", f"Регион в E9 ('{region}') слишком короткий — похоже на плейсхолдер.")
    return _make_result("6", title, "PASS")


# ==============================================================================
# RULE 7 — Дата соглашения с ФЦК/РЦК в G9 (валидная дата)
# ==============================================================================
def _rule_7_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    d = (data.get("general_info", {}).get("agreement_date_fck") or "").strip()
    if not d:
        return _make_result("7", title, "FAIL", "Ячейка G9 (дата соглашения с ФЦК/РЦК) пуста или невалидна.")
    return _make_result("7", title, "PASS")


# ==============================================================================
# RULE 8 — Базовый год соответствует дате соглашения (правило 1 апреля)
# ==============================================================================
def _expected_base_year(iso_date: str) -> int:
    d = datetime.strptime(iso_date, "%Y-%m-%d").date()
    return d.year - 1 if d.month < 4 else d.year


def _rule_8_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    gi = data.get("general_info", {}) or {}
    iso = gi.get("agreement_date_fck")
    actual = gi.get("base_year")
    if not iso:
        return _make_result("8", title, "FAIL", "Невозможно вычислить базовый год: дата соглашения G9 не задана.")
    if actual is None:
        return _make_result("8", title, "FAIL", "Ячейка H9 (базовый год) пуста.")
    try:
        expected = _expected_base_year(iso)
    except ValueError:
        return _make_result("8", title, "FAIL", f"Не удалось распарсить дату G9: '{iso}'")
    if int(actual) != expected:
        return _make_result("8", title, "FAIL",
                            f"Базовый год в H9 = {actual}, но по дате соглашения {iso} "
                            f"должен быть {expected} (правило: до 1 апреля → год-1, с 1 апреля → текущий год).")
    return _make_result("8", title, "PASS")


# ==============================================================================
# RULE 9 — Структура годов в G12:K12 (БГ-1, БГ, БГ+1, БГ+2, БГ+3)
# ==============================================================================
def _rule_9_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    bg = (data.get("general_info") or {}).get("base_year")
    years = data.get("years") or []
    if bg is None:
        return _make_result("9", title, "FAIL", "Невозможно проверить структуру годов: H9 (базовый год) пуст.")
    if len(years) != 5:
        return _make_result("9", title, "FAIL", f"Ожидается 5 годов в G12:K12, найдено: {years}")
    expected = [bg - 1, bg, bg + 1, bg + 2, bg + 3]
    if years != expected:
        diff = [
            f"{cell}: ожидался {exp}, получен {act}"
            for cell, exp, act in zip(["G12", "H12", "I12", "J12", "K12"], expected, years)
            if exp != act
        ]
        return _make_result("9", title, "FAIL", "Структура годов нарушена. " + "; ".join(diff))
    return _make_result("9", title, "PASS")


# ==============================================================================
# RULE 10 — Базовые данные за БГ-1 и БГ заполнены (строки 15-22, колонки G/H)
# ==============================================================================
_RULE_10_ROWS = {
    15: "Выручка", 16: "Прямые расходы", 17: "Косвенные расходы", 18: "Амортизация",
    19: "Расходы на оплату труда", 20: "Страховые взносы", 21: "Налоги в себестоимости", 22: "Численность",
}


def _rule_10_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    ind = data.get("indicators_data") or {}
    issues = []
    for row, label in _RULE_10_ROWS.items():
        for col in ("G", "H"):
            coord = f"{col}{row}"
            cell = ind.get(coord) or {}
            v = cell.get("value")
            if v is None:
                issues.append(f"{coord} ({label}, {col}={'БГ-1' if col=='G' else 'БГ'}) — пусто")
            elif v == 0 and label != "Налоги в себестоимости":
                issues.append(f"{coord} ({label}) = 0 (вероятно не заполнено)")
    if issues:
        return _make_result("10", title, "FAIL", "; ".join(issues[:10]))
    return _make_result("10", title, "PASS")


# ==============================================================================
# RULE 11 — Формулы прогноза в I/J/K для производных показателей
# ==============================================================================
def _rule_11_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    ind = data.get("indicators_data") or {}
    perf = (data.get("performance") or {}).get("formulas") or {}
    idx = (data.get("index") or {}).get("formulas") or {}
    missing = []
    for col in ["G", "H", "I", "J", "K"]:
        coord = f"{col}13"
        if not (ind.get(coord) or {}).get("formula"):
            missing.append(f"{coord} (добавленная стоимость)")
    for col in ["G", "H", "I", "J", "K"]:
        coord = f"{col}14"
        if not (ind.get(coord) or {}).get("formula"):
            missing.append(f"{coord} (прибыль)")
    for coord in ["G23", "H23", "I23", "J23", "K23"]:
        if not perf.get(coord):
            missing.append(f"{coord} (производительность труда)")
    for coord in ["I24", "J24", "K24"]:
        if not idx.get(coord):
            missing.append(f"{coord} (индекс производительности)")
    if missing:
        return _make_result("11", title, "FAIL", f"Отсутствуют формулы прогноза: {', '.join(missing[:10])}")
    return _make_result("11", title, "PASS")


# ==============================================================================
# RULE 12 — ГЛАВНЫЙ КРИТЕРИЙ ФЦК: индекс производительности ≥ 5% (I24,J24,K24)
# ==============================================================================
_RULE_12_TARGET = 0.05
_RULE_12_TOLERANCE = 0.0001


def _rule_12_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    idx_values = (data.get("index") or {}).get("values") or {}
    issues = []
    for coord in ["I24", "J24", "K24"]:
        v = idx_values.get(coord)
        if v is None:
            issues.append(f"{coord}: значение не вычислено")
            continue
        if v + _RULE_12_TOLERANCE < _RULE_12_TARGET:
            issues.append(f"{coord}: {v*100:.2f}% < целевые 5% (требуется ≥ 5%)")
    if issues:
        return _make_result("12", title, "FAIL",
                            "Прирост производительности труда не достигает 5% (главный критерий ФЦК): " + "; ".join(issues))
    return _make_result("12", title, "PASS")


# ==============================================================================
# RULE 13 — Целевые показатели = 5% (I25, J25, K25)
# ==============================================================================
_RULE_13_TARGET = 0.05
_RULE_13_TOLERANCE = 0.0001


def _rule_13_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    targets = data.get("target_index") or {}
    issues = []
    for coord in ["I25", "J25", "K25"]:
        v = targets.get(coord)
        if v is None:
            issues.append(f"{coord}: значение не указано")
            continue
        if abs(v - _RULE_13_TARGET) > _RULE_13_TOLERANCE:
            issues.append(f"{coord}: {v} (ожидается 0.05 = 5%)")
    if issues:
        return _make_result("13", title, "FAIL", "; ".join(issues))
    return _make_result("13", title, "PASS")


# ==============================================================================
# RULE 14 — Подписи (согласие G27, ФИО B31, совпадение дат J28 и B33)
# ==============================================================================
_PLACEHOLDER_RE = re.compile(r"^[\s_*\.\-]+$")


def _rule_14_validate(data: dict, rule: dict) -> dict:
    title = rule["rule_title"]
    sig = data.get("signatures") or {}
    issues = []
    consent = (sig.get("consent") or "").strip().lower()
    if not consent:
        issues.append("G27: согласие на публикацию пусто")
    elif "соглас" not in consent:
        issues.append(f"G27: текст не похож на согласие: '{(sig.get('consent') or '')[:60]}'")
    fio = (sig.get("signer_fio") or "").strip()
    if not fio:
        issues.append("B31: ФИО подписанта пусто")
    else:
        clean = re.sub(r"[_/]+", "", fio).strip()
        if not clean or _PLACEHOLDER_RE.match(clean):
            issues.append(f"B31: ФИО — плейсхолдер ('{fio[:40]}')")
        elif len(clean) < 5:
            issues.append(f"B31: ФИО слишком короткое ('{fio}')")
    sig_iso = sig.get("signature_date_iso")
    doc_iso = sig.get("doc_date_iso")
    if not sig_iso:
        issues.append("J28: дата подписи пуста или не парсится")
    if not doc_iso:
        issues.append(f"B33: дата документа пуста или не парсится ('{sig.get('doc_date_raw') or ''}')")
    if sig_iso and doc_iso and sig_iso != doc_iso:
        issues.append(f"Даты не совпадают: J28 = {sig_iso}, B33 = {doc_iso}")
    if issues:
        return _make_result("14", title, "FAIL", "; ".join(issues))
    return _make_result("14", title, "PASS")


# START_VALIDATOR_DISPATCH
# Маппинг rule_index → функция-валидатор. Все валидаторы python (без LLM).
VALIDATOR_DISPATCH: Dict[str, Callable[[dict, dict], dict]] = {
    "1": _rule_1_validate, "2": _rule_2_validate, "3": _rule_3_validate, "4": _rule_4_validate,
    "5": _rule_5_validate, "6": _rule_6_validate, "7": _rule_7_validate, "8": _rule_8_validate,
    "9": _rule_9_validate, "10": _rule_10_validate, "11": _rule_11_validate, "12": _rule_12_validate,
    "13": _rule_13_validate, "14": _rule_14_validate,
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


# START_FORMA_0_3_RUNNER
def run_forma_0_3_special(args) -> AuditResult:
    """
    Публичный entrypoint для forma_0_3. Регистрируется в SPECIAL_ENGINE_RUNNERS.
    Оркестрирует: 1 парсер → load rules → прогон 14 python-валидаторов → Excel.
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
    parse_forma_0_3(target_path, parser_outputs_dir)

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
            raise ValueError(f"Правило {args.rule_filter} не найдено в forma_0_3")

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
                res = validate_fn(data, rule)
            except Exception as e:
                res = _make_result(rule_index, rule.get("rule_title", ""), "ERROR", f"Исключение: {e}")
        dur = time.time() - t0
        results.append((rule, res, dur))
        # Сохраняем результат правила (как делал прод-движок).
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
# END_FORMA_0_3_RUNNER
