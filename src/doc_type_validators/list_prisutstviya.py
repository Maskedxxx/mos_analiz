# START_MODULE_CONTRACT
# PURPOSE: Special-движок валидации «Лист присутствия» (модуль 1 «3.1» и модуль 2 «3.4»). 13 валидаторов (10 python + 3 LLM) над lp_main.json + раннер run_list_prisutstviya_special.
# INPUTS: args (target, doc_type, parse_only?, rule_filter?, session_dir?); doc_configs/list_prisutstviya_modul_{1,2}/{config.json, validation_rules.json}; данные из src/doc_type_parsers/list_prisutstviya.py.
# OUTPUTS: AuditResult; validation_report.xlsx. Публичный символ: run_list_prisutstviya_special.
# KEYWORDS: list_prisutstviya, special-runner, python-validators, llm-validators, two-modules, xlsx.
# LINKS: main.py (SPECIAL_ENGINE_RUNNERS), src/doc_type_parsers/list_prisutstviya.py (parse_list_prisutstviya), src/llm/client.py (call_llm), src/audit/models.py (AuditResult).
# RATIONALE: Один движок (engine=list_prisutstviya) обслуживает оба doc_type — раннер параметризован по args.doc_type (как kartochka). Модули различаются config (module, allowed_trainings, filename_keywords). Логика 13 валидаторов перенесена из прода as-is; правила 4/5/7 — LLM (через _llm_yes_no на call_llm, эквивалент прод-llm_yes_no: temp=0, enable_thinking=False, json_object, max_tokens=300). Прод-subprocess консолидирован в прямые вызовы.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Tuple

import pandas as pd
from openpyxl.utils import get_column_letter

from config.llm import LLM_CONFIG
from src.audit.models import AuditResult
from src.doc_type_parsers.list_prisutstviya import parse_list_prisutstviya
from src.llm.client import call_llm
# END_IMPORTS


# START_PATHS
_REPO_ROOT = Path(__file__).resolve().parents[2]
_DOC_CONFIGS_DIR = _REPO_ROOT / "doc_configs"
_LOGS_RESULT_DIR = _REPO_ROOT / "logs_result"
# END_PATHS


# START_HELPERS
def _load_config(doc_type: str) -> Dict[str, Any]:
    """Читает doc_configs/<doc_type>/config.json (modul_1 или modul_2)."""
    with open(_DOC_CONFIGS_DIR / doc_type / "config.json", "r", encoding="utf-8") as f:
        return json.load(f)


def _load_rules(doc_type: str) -> List[Dict[str, Any]]:
    """Загружает все правила из doc_configs/<doc_type>/validation_rules.json."""
    with open(_DOC_CONFIGS_DIR / doc_type / "validation_rules.json", "r", encoding="utf-8") as f:
        data = json.load(f)
    return data["rules"] if isinstance(data, dict) else data


def _load_lp_data(parser_outputs_dir: Path) -> dict:
    """Загружает lp_main.json из директории парсера."""
    with open(parser_outputs_dir / "lp_main.json", "r", encoding="utf-8") as f:
        return json.load(f)


def _make_result(rule_index: str, rule_title: str, status: str, discrepancy: str = "") -> Dict[str, Any]:
    """Единый формат результата валидатора."""
    return {"rule_index": rule_index, "rule_title": rule_title, "status": status, "discrepancy": discrepancy}


def _llm_yes_no(prompt: str, config: dict,
                system: str = "Ты эксперт по проверке документов. Отвечаешь строго в JSON.") -> Dict[str, Any]:
    """Вызывает Qwen и возвращает распарсенный JSON. Эквивалент прод-llm_yes_no:
    temperature=0, enable_thinking=False (в call_llm), response_format=json_object, max_tokens=300."""
    raw = call_llm(
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": prompt},
        ],
        model=config.get("model", LLM_CONFIG.default_model),
        temperature=0.0,
        base_url=config.get("llm_base_url", LLM_CONFIG.base_url),
        max_tokens=300,
        response_format={"type": "json_object"},
    )
    return json.loads(raw)
# END_HELPERS


# ==============================================================================
# RULE 1 — Наличие логотипа (≥1 изображение на листе)
# ==============================================================================
def _rule_1_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 1 — наличие логотипа на листе.

    Назначение: проверить, что на листе присутствия есть хотя бы одно встроенное
    изображение (логотип АНО «Мосстратегия»).
    Вход: data — lp_main.json (используется data["images_count"]); rule — правило
    из validation_rules.json (rule_title); config — config.json модуля (не используется).
    Выход: dict формата _make_result. PASS, если images_count >= 1;
    FAIL, если изображений на листе нет.
    """
    title = rule["rule_title"]
    if data.get("images_count", 0) >= 1:
        return _make_result("1", title, "PASS")
    return _make_result("1", title, "FAIL",
                        "На листе нет ни одного встроенного изображения (логотип АНО Мосстратегия отсутствует).")


# ==============================================================================
# RULE 2 — Имя файла: ключевые слова + название тренинга (config)
# ==============================================================================
def _rule_2_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 2 — имя файла: ключевые слова и название тренинга.

    Назначение: проверить, что имя книги Excel содержит все ключевые слова из
    config["filename_keywords"] и хотя бы одно разрешённое название тренинга
    из config["allowed_trainings"] (регистронезависимо).
    Вход: data — lp_main.json (data["meta"]["workbook"]); rule — правило (rule_title);
    config — config.json модуля (filename_keywords, allowed_trainings, module).
    Выход: dict формата _make_result.
    Логика:
      1. Если в имени файла нет хотя бы одного ключевого слова — FAIL со списком отсутствующих.
      2. Если allowed_trainings задан и ни одно название не входит в имя файла — FAIL.
      3. Иначе PASS.
    """
    title = rule["rule_title"]
    fname = data.get("meta", {}).get("workbook", "")
    fname_low = fname.lower()
    keywords = config.get("filename_keywords", [])
    missing = [kw for kw in keywords if kw.lower() not in fname_low]
    if missing:
        return _make_result("2", title, "FAIL", f"Имя файла '{fname}' не содержит ключевые слова: {missing}")
    allowed = config.get("allowed_trainings", [])
    module_n = config.get("module", 1)
    if allowed and not any(t.lower() in fname_low for t in allowed):
        return _make_result("2", title, "FAIL",
                            f"Имя файла '{fname}' не содержит ни одного из разрешённых названий тренинга модуля {module_n}: {allowed}")
    return _make_result("2", title, "PASS")


# ==============================================================================
# RULE 3 — Название тренинга в C1: 'Модуль N' + разрешённое название (config)
# ==============================================================================
def _rule_3_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 3 — название тренинга в ячейке C1.

    Назначение: проверить, что в C1 указан «Модуль N» (N = config["module"]) и одно из
    разрешённых названий тренинга config["allowed_trainings"] (регистронезависимо).
    Вход: data — lp_main.json (data["header"]["training_name"]); rule — правило (rule_title);
    config — config.json модуля (module, allowed_trainings).
    Выход: dict формата _make_result.
    Логика:
      1. Пустая C1 — FAIL.
      2. Нет подстроки «Модуль N» (в любом регистре) — FAIL.
      3. allowed_trainings задан, но ни одно название не найдено в C1 — FAIL.
      4. Иначе PASS.
    """
    title = rule["rule_title"]
    c1 = (data.get("header", {}).get("training_name") or "").strip()
    if not c1:
        return _make_result("3", title, "FAIL", "Ячейка C1 пуста — название тренинга не указано.")
    module_n = config.get("module", 1)
    if f"Модуль {module_n}" not in c1 and f"модуль {module_n}" not in c1.lower():
        return _make_result("3", title, "FAIL", f"В C1 ('{c1}') нет указания 'Модуль {module_n}'.")
    allowed = config.get("allowed_trainings", [])
    c1_low = c1.lower()
    if allowed and not any(t.lower() in c1_low for t in allowed):
        return _make_result("3", title, "FAIL",
                            f"В C1 ('{c1}') нет ни одного разрешённого названия тренинга модуля {module_n}: {allowed}")
    return _make_result("3", title, "PASS")


# ==============================================================================
# RULE 4 — ФИО тренера в C2 (полное, без сокращений). LLM.
# ==============================================================================
def _rule_4_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 4 — ФИО тренера в ячейке C2 (LLM).

    Назначение: проверить, что в C2 указано полное ФИО (Фамилия Имя Отчество)
    без инициалов, сокращений и плейсхолдеров.
    Вход: data — lp_main.json (data["header"]["trainer_fio"]); rule — правило (rule_title);
    config — config.json модуля (model, llm_base_url для _llm_yes_no).
    Выход: dict формата _make_result.
    Логика:
      1. Пустая C2 — FAIL без вызова LLM.
      2. Иначе промпт с критериями PASS/FAIL отправляется в _llm_yes_no; ожидается JSON
         {"status", "discrepancy"}.
      3. Любой статус, кроме "PASS"/"FAIL", трактуется как FAIL; discrepancy берётся
         из ответа LLM только при FAIL.
    """
    title = rule["rule_title"]
    fio = (data.get("header", {}).get("trainer_fio") or "").strip()
    if not fio:
        return _make_result("4", title, "FAIL", "Ячейка C2 пуста — ФИО тренера не указано.")
    prompt = f"""Проверь, является ли строка полным ФИО (Фамилия Имя Отчество) без сокращений и инициалов.

Строка: "{fio}"

Норма (PASS): три отдельных слова — Фамилия, Имя, Отчество — целиком, без точек-инициалов.
Не норма (FAIL): инициалы (например 'И.И. Иванов', 'Иванов И.И.', 'Иванов Иван И.'), плейсхолдер ('ФИО', 'ФИО тренера'), отсутствие отчества (только 2 слова), пустая строка.

Ответ строго в JSON:
{{"status": "PASS" | "FAIL", "discrepancy": "Краткое описание проблемы если FAIL, иначе пустая строка"}}"""
    res = _llm_yes_no(prompt, config)
    status = res.get("status", "FAIL")
    if status not in ("PASS", "FAIL"):
        status = "FAIL"
    return _make_result("4", title, status, res.get("discrepancy", "") if status == "FAIL" else "")


# ==============================================================================
# RULE 5 — Адрес проведения в C3 (город/улица/дом). LLM.
# ==============================================================================
def _rule_5_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 5 — адрес проведения в ячейке C3 (LLM).

    Назначение: проверить, что в C3 указан адрес с городом, улицей (или иным топонимом)
    и номером дома; плейсхолдер или только регион — нарушение.
    Вход: data — lp_main.json (data["header"]["address"]); rule — правило (rule_title);
    config — config.json модуля (model, llm_base_url для _llm_yes_no).
    Выход: dict формата _make_result.
    Логика:
      1. Пустая C3 — FAIL без вызова LLM.
      2. Иначе промпт с критериями PASS/FAIL отправляется в _llm_yes_no; ожидается JSON
         {"status", "discrepancy"}.
      3. Любой статус, кроме "PASS"/"FAIL", трактуется как FAIL; discrepancy берётся
         из ответа LLM только при FAIL.
    """
    title = rule["rule_title"]
    addr = (data.get("header", {}).get("address") or "").strip()
    if not addr:
        return _make_result("5", title, "FAIL", "Ячейка C3 пуста — адрес проведения не указан.")
    prompt = f"""Проверь адрес проведения мероприятия.

Адрес: "{addr}"

Норма (PASS): указан город, улица (или иной топоним — проспект, переулок и т.п.), номер дома. Допустимы сокращения 'г.' для города (например 'г. Москва') — это устоявшееся сокращение.
Не норма (FAIL): отсутствует город, ИЛИ отсутствует улица, ИЛИ отсутствует номер дома, ИЛИ это плейсхолдер ('Регион, Адрес', 'г. ___'), ИЛИ только название региона без улицы.

Замечание: 'ул.' / 'пр.' / 'пер.' — допустимые стандартные сокращения, НЕ считай их нарушением.

Ответ строго в JSON:
{{"status": "PASS" | "FAIL", "discrepancy": "Что именно не так если FAIL, иначе пустая строка"}}"""
    res = _llm_yes_no(prompt, config)
    status = res.get("status", "FAIL")
    if status not in ("PASS", "FAIL"):
        status = "FAIL"
    return _make_result("5", title, status, res.get("discrepancy", "") if status == "FAIL" else "")


# ==============================================================================
# RULE 6 — ФИО участников (B6:B15): полные, без повторов. Python.
# ==============================================================================
_FIO_RE = re.compile(r"^[А-ЯЁA-Z][а-яёa-z]+(?:-[А-ЯЁA-Z][а-яёa-z]+)?\s+[А-ЯЁA-Z][а-яёa-z]+\s+[А-ЯЁA-Z][а-яёa-z]+$")


def _rule_6_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 6 — ФИО участников (B6:B15): полные, без повторов.

    Назначение: проверить, что у каждого участника ФИО заполнено, соответствует
    шаблону _FIO_RE (три слова с заглавной, допускается двойная фамилия через дефис)
    и не дублируется в других строках (сравнение без учёта регистра).
    Вход: data — lp_main.json (data["participants"], поля row, fio); rule — правило
    (rule_title); config — config.json модуля (не используется).
    Выход: dict формата _make_result.
    Логика:
      1. Пустая таблица участников — FAIL.
      2. Собираются замечания по формату: пустое ФИО или несовпадение с _FIO_RE.
      3. Собираются дубли ФИО с указанием пар строк.
      4. Есть замечания — FAIL с первыми 10 через «; », иначе PASS.
    """
    title = rule["rule_title"]
    parts = data.get("participants", []) or []
    if not parts:
        return _make_result("6", title, "FAIL", "Таблица участников пустая (нет ФИО ни в одной строке B6:B15).")
    bad_format = []
    for p in parts:
        fio = (p.get("fio") or "").strip()
        if not fio:
            bad_format.append(f"строка {p['row']}: ФИО пустое")
            continue
        if not _FIO_RE.match(fio):
            bad_format.append(f"строка {p['row']}: '{fio}' — не похоже на полное Фамилия Имя Отчество (3 слова с заглавных)")
    seen = {}
    duplicates = []
    for p in parts:
        fio = (p.get("fio") or "").strip()
        if not fio:
            continue
        key = fio.lower()
        if key in seen:
            duplicates.append(f"'{fio}' дублируется в строках {seen[key]} и {p['row']}")
        else:
            seen[key] = p["row"]
    issues = bad_format + duplicates
    if issues:
        return _make_result("6", title, "FAIL", "; ".join(issues[:10]))
    return _make_result("6", title, "PASS")


# ==============================================================================
# RULE 7 — Должности (C6:C15) без сокращений. LLM (1 batch-вызов).
# ==============================================================================
def _rule_7_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 7 — должности участников (C6:C15) без сокращений (LLM, один batch-вызов).

    Назначение: проверить, что должности заполнены и записаны полностью, без
    сокращений и аббревиатур («м-р», «нач.», «инж.» и т.п.); опечатки и регистр не проверяются.
    Вход: data — lp_main.json (data["participants"], поля row, position); rule — правило
    (rule_title); config — config.json модуля (model, llm_base_url для _llm_yes_no).
    Выход: dict формата _make_result.
    Логика:
      1. Пустая таблица участников — FAIL.
      2. Есть строки с пустой должностью — FAIL с перечнем строк, без вызова LLM.
      3. Иначе все должности одним списком отправляются в _llm_yes_no; ожидается JSON
         {"status", "discrepancy"}.
      4. Любой статус, кроме "PASS"/"FAIL", трактуется как FAIL; discrepancy берётся
         из ответа LLM только при FAIL.
    """
    title = rule["rule_title"]
    parts = data.get("participants", []) or []
    if not parts:
        return _make_result("7", title, "FAIL", "Таблица участников пустая.")
    rows_with_pos = [(p["row"], (p.get("position") or "").strip()) for p in parts]
    empty = [str(r) for r, v in rows_with_pos if not v]
    if empty:
        return _make_result("7", title, "FAIL", f"Должность не указана в строках: {', '.join(empty)}")
    listing = "\n".join(f'  - строка {r}: "{v}"' for r, v in rows_with_pos)
    prompt = f"""Проверь должности сотрудников. Должны быть указаны полностью, без сокращений и аббревиатур.

Должности:
{listing}

Норма (PASS): полные слова — 'мастер', 'начальник', 'инженер', 'руководитель отдела', 'главный специалист'.
Не норма (FAIL): сокращения — 'м-р' (вместо 'мастер'), 'нач.' / 'нач-к' (вместо 'начальник'), 'инж.' (вместо 'инженер'), 'рук-ль', 'гл. спец.', 'зам.', аббревиатуры неочевидного смысла.

Опечатки в самих словах (например 'Начельник' вместо 'Начальник') — это НЕ сокращение, не проверяй опечатки.
Регистр (с заглавной или нет) — не важен.

Ответ строго в JSON:
{{
  "status": "PASS" | "FAIL",
  "discrepancy": "Перечисление строк с сокращениями (например 'строка 7: \\"м-р\\" — сокращение от мастер'). Если PASS — пустая строка."
}}"""
    res = _llm_yes_no(prompt, config)
    status = res.get("status", "FAIL")
    if status not in ("PASS", "FAIL"):
        status = "FAIL"
    return _make_result("7", title, status, res.get("discrepancy", "") if status == "FAIL" else "")


# ==============================================================================
# RULE 8 — Организация (D6:D15): юр. форма + название, одинаково везде. Python.
# ==============================================================================
_LEGAL_FORMS = ["ООО", "ЗАО", "АО", "ПАО", "ИП", "ОАО", "ФГУП", "ГУП", "МУП", "АНО"]
_LEGAL_RE = re.compile(r"\b(" + "|".join(_LEGAL_FORMS) + r")\b", re.IGNORECASE)


def _rule_8_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 8 — организация участников (D6:D15): юр. форма и единообразие.

    Назначение: проверить, что организация заполнена в каждой строке, содержит
    юридическую форму из _LEGAL_FORMS (ООО/ЗАО/АО/ПАО/ИП/…) и одинакова во всех строках
    (сравнение без учёта регистра и краевых пробелов).
    Вход: data — lp_main.json (data["participants"], поля row, org); rule — правило
    (rule_title); config — config.json модуля (не используется).
    Выход: dict формата _make_result.
    Логика:
      1. Пустая таблица участников — FAIL.
      2. Есть строки с пустой организацией — FAIL с перечнем строк.
      3. Есть строки без юр. формы (_LEGAL_RE не находит) — FAIL с перечнем строк.
      4. Более одного уникального названия — FAIL с примерами (до 5 строк).
      5. Иначе PASS.
    """
    title = rule["rule_title"]
    parts = data.get("participants", []) or []
    if not parts:
        return _make_result("8", title, "FAIL", "Таблица участников пустая.")
    rows_with_org = [(p["row"], (p.get("org") or "").strip()) for p in parts]
    empty = [str(r) for r, v in rows_with_org if not v]
    if empty:
        return _make_result("8", title, "FAIL", f"Организация не указана в строках: {', '.join(empty)}")
    no_legal = [str(r) for r, v in rows_with_org if not _LEGAL_RE.search(v)]
    if no_legal:
        return _make_result("8", title, "FAIL",
                            f"В строках {', '.join(no_legal)} в организации нет юридической формы (ООО/ЗАО/АО/ПАО/ИП и т.п.).")
    uniq = {v.strip().lower(): r for r, v in rows_with_org}
    if len(uniq) > 1:
        examples = "; ".join(f'строка {r}: "{[v for rr, v in rows_with_org if rr == r][0]}"' for r in list(uniq.values())[:5])
        return _make_result("8", title, "FAIL", f"Организация различается между строками: {examples}")
    return _make_result("8", title, "PASS")


# ==============================================================================
# RULE 9 — ИНН (E6:E15): 10/12 цифр, одинаково везде. Python.
# ==============================================================================
_INN_RE = re.compile(r"^\d{10}$|^\d{12}$")


def _rule_9_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 9 — ИНН участников (E6:E15): формат и единообразие.

    Назначение: проверить, что ИНН заполнен в каждой строке, состоит ровно из 10
    или 12 цифр (_INN_RE) и одинаков во всех строках.
    Вход: data — lp_main.json (data["participants"], поля row, inn); rule — правило
    (rule_title); config — config.json модуля (не используется).
    Выход: dict формата _make_result.
    Логика:
      1. Пустая таблица участников — FAIL.
      2. Есть строки с пустым ИНН — FAIL с перечнем строк.
      3. Есть ИНН, не совпадающие с _INN_RE — FAIL с примерами (до 5).
      4. Более одного уникального ИНН — FAIL с перечнем значений.
      5. Иначе PASS.
    """
    title = rule["rule_title"]
    parts = data.get("participants", []) or []
    if not parts:
        return _make_result("9", title, "FAIL", "Таблица участников пустая.")
    rows_with_inn = [(p["row"], str(p.get("inn") or "").strip()) for p in parts]
    empty = [str(r) for r, v in rows_with_inn if not v]
    if empty:
        return _make_result("9", title, "FAIL", f"ИНН не указан в строках: {', '.join(empty)}")
    bad = [(r, v) for r, v in rows_with_inn if not _INN_RE.match(v)]
    if bad:
        examples = "; ".join(f"строка {r}: '{v}'" for r, v in bad[:5])
        return _make_result("9", title, "FAIL", f"ИНН должен содержать 10 или 12 цифр. Некорректные: {examples}")
    uniq = set(v for _, v in rows_with_inn)
    if len(uniq) > 1:
        return _make_result("9", title, "FAIL", f"ИНН различается между строками: {sorted(uniq)}")
    return _make_result("9", title, "PASS")


# ==============================================================================
# RULE 10 — Регион (F6:F15) = expected_region (config). Python.
# ==============================================================================
def _normalize_region(s: str) -> str:
    """Нормализует название региона для сравнения: trim, нижний регистр, без точек и пробелов."""
    return s.strip().lower().replace(".", "").replace(" ", "")


def _rule_10_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 10 — регион участников (F6:F15) равен ожидаемому.

    Назначение: проверить, что регион заполнен в каждой строке и после нормализации
    (_normalize_region) совпадает с config["expected_region"] (по умолчанию «г. Москва»).
    Вход: data — lp_main.json (data["participants"], поля row, region); rule — правило
    (rule_title); config — config.json модуля (expected_region).
    Выход: dict формата _make_result.
    Логика:
      1. Пустая таблица участников — FAIL.
      2. Для каждой строки: пустой регион или несовпадение с ожидаемым — замечание.
      3. Есть замечания — FAIL с первыми 10 через «; », иначе PASS.
    """
    title = rule["rule_title"]
    parts = data.get("participants", []) or []
    if not parts:
        return _make_result("10", title, "FAIL", "Таблица участников пустая.")
    expected = config.get("expected_region", "г. Москва")
    expected_norm = _normalize_region(expected)
    bad = []
    for p in parts:
        v = (p.get("region") or "").strip()
        if not v:
            bad.append(f"строка {p['row']}: регион не указан")
        elif _normalize_region(v) != expected_norm:
            bad.append(f"строка {p['row']}: '{v}' (ожидается '{expected}')")
    if bad:
        return _make_result("10", title, "FAIL", "; ".join(bad[:10]))
    return _make_result("10", title, "PASS")


# ==============================================================================
# RULE 11 — E-mail (G6:G15): валидный формат, без повторов. Python.
# ==============================================================================
_EMAIL_RE = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")


def _rule_11_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 11 — e-mail участников (G6:G15): валидный формат, без повторов.

    Назначение: проверить, что e-mail заполнен в каждой строке, соответствует _EMAIL_RE
    (локальная часть@домен.зона без пробелов) и не дублируется (сравнение без учёта регистра).
    Вход: data — lp_main.json (data["participants"], поля row, email); rule — правило
    (rule_title); config — config.json модуля (не используется).
    Выход: dict формата _make_result.
    Логика:
      1. Пустая таблица участников — FAIL.
      2. Для каждой строки: пустой e-mail, невалидный формат или дубль — замечание.
      3. Есть замечания — FAIL с первыми 10 через «; », иначе PASS.
    """
    title = rule["rule_title"]
    parts = data.get("participants", []) or []
    if not parts:
        return _make_result("11", title, "FAIL", "Таблица участников пустая.")
    issues = []
    seen = {}
    for p in parts:
        v = (p.get("email") or "").strip()
        r = p["row"]
        if not v:
            issues.append(f"строка {r}: email не указан")
            continue
        if not _EMAIL_RE.match(v):
            issues.append(f"строка {r}: '{v}' — невалидный email")
            continue
        key = v.lower()
        if key in seen:
            issues.append(f"строка {r}: '{v}' дублируется со строкой {seen[key]}")
        else:
            seen[key] = r
    if issues:
        return _make_result("11", title, "FAIL", "; ".join(issues[:10]))
    return _make_result("11", title, "PASS")


# ==============================================================================
# RULE 12 — Телефон (H6:H15): 10-12 цифр, без повторов. Python.
# ==============================================================================
_DIGITS_ONLY = re.compile(r"\D+")


def _rule_12_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 12 — телефон участников (H6:H15): 10–12 цифр, без повторов.

    Назначение: проверить, что телефон заполнен в каждой строке, после удаления
    всех нецифровых символов (_DIGITS_ONLY) содержит от 10 до 12 цифр и не дублируется
    (сравнение по нормализованным цифрам).
    Вход: data — lp_main.json (data["participants"], поля row, phone); rule — правило
    (rule_title); config — config.json модуля (не используется).
    Выход: dict формата _make_result.
    Логика:
      1. Пустая таблица участников — FAIL.
      2. Для каждой строки: пустой телефон, число цифр вне 10–12 или дубль — замечание.
      3. Есть замечания — FAIL с первыми 10 через «; », иначе PASS.
    """
    title = rule["rule_title"]
    parts = data.get("participants", []) or []
    if not parts:
        return _make_result("12", title, "FAIL", "Таблица участников пустая.")
    issues = []
    seen = {}
    for p in parts:
        raw = str(p.get("phone") or "").strip()
        r = p["row"]
        if not raw:
            issues.append(f"строка {r}: телефон не указан")
            continue
        digits = _DIGITS_ONLY.sub("", raw)
        if len(digits) < 10 or len(digits) > 12:
            issues.append(f"строка {r}: '{raw}' — должно быть 10–12 цифр, найдено {len(digits)}")
            continue
        if digits in seen:
            issues.append(f"строка {r}: телефон '{raw}' дублируется со строкой {seen[digits]}")
        else:
            seen[digits] = r
    if issues:
        return _make_result("12", title, "FAIL", "; ".join(issues[:10]))
    return _make_result("12", title, "PASS")


# ==============================================================================
# RULE 13 — Дата проведения в I5 (дд.мм.гггг или Excel datetime). Python.
# ==============================================================================
_DATE_PLACEHOLDER_PATTERNS = [r"_+", r"\?{2,}", r"^Поставьте"]


def _rule_13_validate(data: dict, rule: dict, config: dict) -> dict:
    """
    Правило 13 — дата проведения в ячейке I5.

    Назначение: проверить, что в I5 указана реальная дата (дд.мм.гггг или Excel-datetime),
    а не пусто/плейсхолдер, и что парсер смог привести её к ISO.
    Вход: data — lp_main.json (data["header"]["date_cell"] — сырой текст ячейки,
    data["header"]["date_iso"] — распознанная дата); rule — правило (rule_title);
    config — config.json модуля (не используется).
    Выход: dict формата _make_result.
    Логика:
      1. Пустая I5 — FAIL.
      2. Сырой текст совпадает с одним из _DATE_PLACEHOLDER_PATTERNS (подчёркивания,
         «??», «Поставьте…») — FAIL «плейсхолдер вместо даты».
      3. date_iso пустой (парсер не распознал дату) — FAIL.
      4. Иначе PASS.
    """
    title = rule["rule_title"]
    header = data.get("header", {})
    raw = (header.get("date_cell") or "").strip()
    iso = (header.get("date_iso") or "").strip()
    if not raw:
        return _make_result("13", title, "FAIL", "Ячейка I5 пуста — дата проведения не указана.")
    for pat in _DATE_PLACEHOLDER_PATTERNS:
        if re.search(pat, raw):
            return _make_result("13", title, "FAIL", f"В ячейке I5 плейсхолдер вместо даты: '{raw}'")
    if not iso:
        return _make_result("13", title, "FAIL",
                            f"Не удалось распарсить дату в I5: '{raw}'. Ожидается формат дд.мм.гггг или Excel-дата.")
    return _make_result("13", title, "PASS")


# START_VALIDATOR_DISPATCH
# Маппинг rule_index → функция-валидатор. Правила 4/5/7 — LLM, остальные python.
VALIDATOR_DISPATCH: Dict[str, Callable[[dict, dict, dict], dict]] = {
    "1": _rule_1_validate, "2": _rule_2_validate, "3": _rule_3_validate, "4": _rule_4_validate,
    "5": _rule_5_validate, "6": _rule_6_validate, "7": _rule_7_validate, "8": _rule_8_validate,
    "9": _rule_9_validate, "10": _rule_10_validate, "11": _rule_11_validate, "12": _rule_12_validate,
    "13": _rule_13_validate,
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


# START_LIST_PRISUTSTVIYA_RUNNER
def run_list_prisutstviya_special(args: Any) -> AuditResult:
    """
    Публичный entrypoint для list_prisutstviya. Регистрируется в SPECIAL_ENGINE_RUNNERS.
    Один движок на оба модуля — doc_type из args выбирает config/rules (modul_1 / modul_2).
    Оркестрирует: 1 парсер → load rules → прогон 13 валидаторов (3 LLM) → Excel.
    """
    start_time = time.time()
    target_path = Path(args.target)
    doc_type = getattr(args, "doc_type", None) or "list_prisutstviya_modul_1"
    config = _load_config(doc_type)

    session_dir = (
        Path(args.session_dir) if getattr(args, "session_dir", None)
        else _LOGS_RESULT_DIR / doc_type / f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    )
    session_dir.mkdir(parents=True, exist_ok=True)
    parser_outputs_dir = session_dir / "parser_outputs"
    validation_outputs_dir = session_dir / "validation_outputs"
    validation_outputs_dir.mkdir(parents=True, exist_ok=True)

    # === Шаг 1: Парсинг XLSX ===
    parse_list_prisutstviya(target_path, parser_outputs_dir)

    if getattr(args, "parse_only", False):
        return AuditResult(
            doc_type=doc_type, session_dir=session_dir, target_path=str(target_path),
            duration_sec=time.time() - start_time,
        )

    # === Шаг 2: Валидация ===
    data = _load_lp_data(parser_outputs_dir)
    rules = _load_rules(doc_type)
    if getattr(args, "rule_filter", None):
        rules = [r for r in rules if r["rule_index"] == str(args.rule_filter)]
        if not rules:
            raise ValueError(f"Правило {args.rule_filter} не найдено в {doc_type}")

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
        violations=violations, doc_type=doc_type, session_dir=session_dir,
        duration_sec=time.time() - start_time, rules_checked=len(df), target_path=str(target_path),
    )
# END_LIST_PRISUTSTVIYA_RUNNER
