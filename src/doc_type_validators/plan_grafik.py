# START_MODULE_CONTRACT
# PURPOSE: Аудит документа «План-график» (xlsx). 9 детерминистских non-LLM проверок (имя файла, блок УТВЕРЖДАЮ, заголовки, даты, ответственные, расчётные колонки, подпись, формулы) + публичный entrypoint `run_plan_grafik_special` для регистрации в `SPECIAL_ENGINE_RUNNERS`.
# INPUTS: xlsx-файл; парсинг через `parse_plan_grafik` (src/doc_type_parsers/plan_grafik.py); конфиг отсутствует (правила захардкожены в валидаторах).
# OUTPUTS: `run_plan_grafik_special(args) -> AuditResult`. Пишет в session_dir: `parsed.json`, `validation_outputs/validate_N.json` (per-rule), `validation_report.xlsx` (6 колонок — формат special-типов).
# KEYWORDS: validators, plan-grafik, non-llm, deterministic, runner.
# LINKS: src/doc_type_parsers/plan_grafik.py (parse_plan_grafik), src/audit/models.py (AuditResult). Отчёт — локальный `_create_excel_report` (как forma/kartochka), НЕ generic save_to_excel.
# RATIONALE:
#   План-график — xlsx со структурированными ячейками. Проверки простые (regex,
#   наличие полей, сравнение дат) — не требуют LLM. По симметрии с drivers.py
#   и kpsc.py весь doc_type-специфичный код (валидаторы + runner) живёт в одном
#   файле.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd
from openpyxl.utils import get_column_letter

from src.audit.models import AuditResult
from src.doc_type_parsers.plan_grafik import parse_plan_grafik
# END_IMPORTS


# START_PATHS
# PURPOSE: parents[2] = repo root (файл лежит в src/doc_type_validators/).
_LOGS_RESULT_DIR = Path(__file__).resolve().parents[2] / "logs_result"
# END_PATHS


def validate_1_filename(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 1: Проверка имени файла."""
    filename = Path(target_path).stem.lower()
    keywords = ['2.6', 'план', 'график']
    missing = []
    for kw in keywords:
        if kw == '2.6':
            if '2.6' not in filename and '2_6' not in filename:
                missing.append(kw)
        elif kw.lower() not in filename:
            missing.append(kw)
    if missing:
        return [{'rule_index': 1, 'rule_title': 'Проверка названия файла', 'Целевой документ': Path(target_path).name, 'Различие': f'Отсутствуют ключевые слова: {', '.join(missing)}'}]
    return []

def validate_2_approval(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 2: Проверка блока УТВЕРЖДАЮ."""
    approval = parsed.get('approval', {})
    violations = []
    marker = approval.get('marker', '')
    if 'утверждаю' not in marker.lower():
        violations.append({'rule_index': 2, 'rule_title': 'Проверка блока УТВЕРЖДАЮ', 'Целевой документ': f'DM8: {marker or 'пусто'}', 'Различие': 'Отсутствует слово «УТВЕРЖДАЮ»'})
    position = approval.get('position', '')
    if not position or (position == 'Генеральный директор' and len(position) < 5):
        pass
    if not position:
        violations.append({'rule_index': 2, 'rule_title': 'Проверка блока УТВЕРЖДАЮ', 'Целевой документ': f'DM9: пусто', 'Различие': 'Не указана должность подписанта'})
    company = approval.get('company', '')
    is_placeholder = not company or ('___' in company and (not re.search('[А-Яа-яA-Za-z]{3,}', company.replace('ООО', '').replace('АО', '').replace('ЗАО', ''))))
    if is_placeholder:
        violations.append({'rule_index': 2, 'rule_title': 'Проверка блока УТВЕРЖДАЮ', 'Целевой документ': f'DM10: {company or 'пусто'}', 'Различие': 'Наименование организации не заполнено (плейсхолдер)'})
    fio = approval.get('fio', '')
    is_fio_placeholder = not fio or 'фио' in fio.lower() or (fio.count('_') > 3 and (not re.search('[А-Яа-я]{2,}', fio.replace('ФИО', ''))))
    if is_fio_placeholder:
        violations.append({'rule_index': 2, 'rule_title': 'Проверка блока УТВЕРЖДАЮ', 'Целевой документ': f'DM11: {fio or 'пусто'}', 'Различие': 'ФИО подписанта не заполнено'})
    return violations

def validate_3_simple_headers(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 3: Проверка простых заголовков таблицы."""
    headers = parsed.get('headers_row16', {})
    if not headers:
        return [{'rule_index': 3, 'rule_title': 'Проверка заголовков таблицы', 'Целевой документ': 'Строка 16 пустая', 'Различие': 'Заголовки таблицы отсутствуют'}]
    all_text = ' '.join(headers.values()).lower()
    required = {'мероприятие': 'Мероприятие', 'ответственн': 'Ответственный', 'начало': 'Начало мероприятия', 'окончани': 'Окончание мероприятия', 'статус': 'Статус'}
    missing = []
    for keyword, name in required.items():
        if keyword not in all_text:
            missing.append(name)
    if missing:
        return [{'rule_index': 3, 'rule_title': 'Проверка заголовков таблицы', 'Целевой документ': ', '.join(headers.values())[:200], 'Различие': f'Отсутствуют столбцы: {', '.join(missing)}'}]
    return []

def validate_4_complex_headers(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 4: Проверка сложных двухуровневых заголовков."""
    h16 = parsed.get('headers_row16', {})
    h17 = parsed.get('headers_row17', {})
    all_text = ' '.join(list(h16.values()) + list(h17.values())).lower()
    required_sub = {'выработк': 'Влияние на показатель выработка', 'запас': 'Влияние на показатель запасы', 'впп': 'Влияние на показатель ВПП', 'проблем': '№ проблемы из КПСЦ', 'комментари': 'Комментарии'}
    missing = []
    for keyword, name in required_sub.items():
        if keyword not in all_text:
            missing.append(name)
    if missing:
        return [{'rule_index': 4, 'rule_title': 'Проверка структуры заголовков', 'Целевой документ': f'Строки 16-17: {len(h16)} + {len(h17)} столбцов', 'Различие': f'Отсутствуют подзаголовки: {', '.join(missing)}'}]
    return []

def validate_5_dates(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 5: Проверка дат (формат + логика)."""
    violations = []
    dates = parsed.get('dates', {})
    start_raw = dates.get('start_raw')
    end_raw = dates.get('end_raw')
    if not start_raw:
        violations.append({'rule_index': 5, 'rule_title': 'Проверка дат мероприятий', 'Целевой документ': 'I12: пусто', 'Различие': 'Дата начала мероприятий не заполнена'})
    if not end_raw:
        violations.append({'rule_index': 5, 'rule_title': 'Проверка дат мероприятий', 'Целевой документ': 'I13: пусто', 'Различие': 'Дата окончания мероприятий не заполнена'})
    if start_raw and end_raw:
        try:
            start_dt = start_raw if isinstance(start_raw, datetime) else datetime.strptime(str(start_raw)[:10], '%Y-%m-%d')
            end_dt = end_raw if isinstance(end_raw, datetime) else datetime.strptime(str(end_raw)[:10], '%Y-%m-%d')
            if end_dt <= start_dt:
                violations.append({'rule_index': 5, 'rule_title': 'Проверка дат мероприятий', 'Целевой документ': f'Начало: {start_dt.strftime('%d.%m.%Y')}, Окончание: {end_dt.strftime('%d.%m.%Y')}', 'Различие': 'Дата окончания не позже даты начала'})
        except (ValueError, TypeError):
            pass
    data_rows = parsed.get('data_rows', [])
    plan_rows_no_dates = []
    plan_rows_bad_logic = []
    for row_data in data_rows:
        if row_data.get('plan_fact', '').lower() != 'план':
            continue
        row_num = row_data.get('row', '?')
        start = row_data.get('start_date')
        end = row_data.get('end_date')
        if not start and (not end):
            plan_rows_no_dates.append(str(row_num))
        elif start and end:
            try:
                s = start if isinstance(start, datetime) else datetime.strptime(str(start)[:10], '%Y-%m-%d')
                e = end if isinstance(end, datetime) else datetime.strptime(str(end)[:10], '%Y-%m-%d')
                if e < s:
                    plan_rows_bad_logic.append(str(row_num))
            except (ValueError, TypeError):
                pass
    if plan_rows_no_dates:
        violations.append({'rule_index': 5, 'rule_title': 'Проверка дат мероприятий', 'Целевой документ': f'Строки без дат: {', '.join(plan_rows_no_dates[:10])}', 'Различие': 'Даты начала/окончания не заполнены для плановых мероприятий'})
    if plan_rows_bad_logic:
        violations.append({'rule_index': 5, 'rule_title': 'Проверка дат мероприятий', 'Целевой документ': f'Строки с нарушением логики: {', '.join(plan_rows_bad_logic[:10])}', 'Различие': 'Дата окончания раньше даты начала'})
    return violations

def validate_6_responsible(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 6: Проверка заполненности ответственных (строки «План»)."""
    data_rows = parsed.get('data_rows', [])
    empty_rows = []
    for row_data in data_rows:
        if row_data.get('plan_fact', '').lower() != 'план':
            continue
        responsible = row_data.get('responsible', '').strip()
        if not responsible:
            row_num = row_data.get('row', '?')
            problem = row_data.get('problem_num', '?')
            empty_rows.append(f'строка {row_num} (проблема №{problem})')
    if empty_rows:
        return [{'rule_index': 6, 'rule_title': 'Проверка заполненности ответственных', 'Целевой документ': f'Пустые: {', '.join(empty_rows[:10])}', 'Различие': 'Не указан ответственный за мероприятие (строки «План»)'}]
    return []

def validate_7_calc_columns(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 7: Проверка наличия расчётных столбцов (Статус, Отклонения, Комментарии)."""
    sc = parsed.get('status_columns', {})
    violations = []
    if not sc.get('status_present'):
        violations.append({'rule_index': 7, 'rule_title': 'Проверка расчётных столбцов', 'Целевой документ': 'Столбец DJ (Статус): отсутствует', 'Различие': 'Столбец «Статус» не найден в заголовках таблицы'})
    if not sc.get('comments_present'):
        violations.append({'rule_index': 7, 'rule_title': 'Проверка расчётных столбцов', 'Целевой документ': 'Столбец DM (Комментарии): отсутствует', 'Различие': 'Столбец «Комментарии» не найден в заголовках таблицы'})
    formulas = parsed.get('formulas', {})
    dk18 = formulas.get('DK18', {})
    dl18 = formulas.get('DL18', {})
    if not dk18.get('is_formula'):
        violations.append({'rule_index': 7, 'rule_title': 'Проверка расчётных столбцов', 'Целевой документ': f'DK18: {dk18.get('value', 'пусто')}', 'Различие': 'Столбец «Отклонение по началу» не содержит формулу'})
    if not dl18.get('is_formula'):
        violations.append({'rule_index': 7, 'rule_title': 'Проверка расчётных столбцов', 'Целевой документ': f'DL18: {dl18.get('value', 'пусто')}', 'Различие': 'Столбец «Отклонение по окончанию» не содержит формулу'})
    return violations

def validate_8_signature(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 8: Проверка блока подписи внизу документа."""
    approval = parsed.get('approval', {})
    date_line = approval.get('date_line', '')
    signature = parsed.get('signature', '').strip()
    if not date_line and (not signature):
        return [{'rule_index': 8, 'rule_title': 'Проверка блока подписи', 'Целевой документ': 'отсутствует', 'Различие': 'Строка подписи с датой не найдена в документе'}]
    return []

def validate_9_formulas(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 9: Проверка целостности формул (не заменены на значения, нет #REF)."""
    formulas = parsed.get('formulas', {})
    violations = []
    dk17 = formulas.get('DK17', {})
    if not dk17.get('is_formula'):
        violations.append({'rule_index': 9, 'rule_title': 'Проверка целостности формул', 'Целевой документ': f'DK17: {dk17.get('value', 'пусто')}', 'Различие': 'Формула длительности заменена на значение или отсутствует'})
    n12 = formulas.get('N12', {})
    if not n12.get('is_formula'):
        violations.append({'rule_index': 9, 'rule_title': 'Проверка целостности формул', 'Целевой документ': f'N12: {n12.get('value', 'пусто')}', 'Различие': 'Формулы Ганта заменены на значения или отсутствуют'})
    for cell_name, cell_data in formulas.items():
        if cell_data.get('has_error'):
            violations.append({'rule_index': 9, 'rule_title': 'Проверка целостности формул', 'Целевой документ': f'{cell_name}: {cell_data.get('value', '')}', 'Различие': 'Формула содержит ошибку #REF! (сломанная ссылка)'})
    return violations


# START_RULES
# PURPOSE: Метаданные 9 правил plan_grafik. Правила захардкожены в валидаторах
# (validation_rules.json у типа НЕТ) — индекс/заголовок/функция заданы здесь.
# `section` пуст: источника секций нет, не выдумываем (правило #0).
_RULES: List[Dict[str, Any]] = [
    {"rule_index": "1", "rule_title": "Проверка названия файла", "section": "", "fn": validate_1_filename},
    {"rule_index": "2", "rule_title": "Проверка блока УТВЕРЖДАЮ", "section": "", "fn": validate_2_approval},
    {"rule_index": "3", "rule_title": "Проверка заголовков таблицы", "section": "", "fn": validate_3_simple_headers},
    {"rule_index": "4", "rule_title": "Проверка структуры заголовков", "section": "", "fn": validate_4_complex_headers},
    {"rule_index": "5", "rule_title": "Проверка дат мероприятий", "section": "", "fn": validate_5_dates},
    {"rule_index": "6", "rule_title": "Проверка заполненности ответственных", "section": "", "fn": validate_6_responsible},
    {"rule_index": "7", "rule_title": "Проверка расчётных столбцов", "section": "", "fn": validate_7_calc_columns},
    {"rule_index": "8", "rule_title": "Проверка блока подписи", "section": "", "fn": validate_8_signature},
    {"rule_index": "9", "rule_title": "Проверка целостности формул", "section": "", "fn": validate_9_formulas},
]
# END_RULES


# START_REPORT
def _create_excel_report(results: List[Tuple[Dict, Dict, float]], report_path: Path) -> Any:
    """Собирает validation_report.xlsx (6 колонок) — единый формат special-типов (как forma/kartochka)."""
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


# START_RUNNER
def run_plan_grafik_special(args):
    """
    Назначение:
        Публичный entrypoint для plan_grafik. Регистрируется в `SPECIAL_ENGINE_RUNNERS`
        в main.py. Парсит xlsx, прогоняет 9 валидаторов, пишет Excel-отчёт.

    Вход:
        args: argparse-Namespace со полями {target, session_dir?, parse_only?}.

    Выход:
        AuditResult.

    Логика:
        1. Создаёт session_dir.
        2. parse_plan_grafik(target). Если падает — пишет ERROR.txt и возвращает пустой AuditResult.
        3. Сохраняет parsed.json. Если parse_only — выходит.
        4. Прогон 9 правил по одному → per-rule результат (PASS/FAIL/ERROR).
        5. validation_outputs/validate_N.json + validation_report.xlsx (6 колонок).
    """

    start_time = time.time()
    target_path = str(args.target)
    session_dir = (
        Path(args.session_dir) if args.session_dir
        else _LOGS_RESULT_DIR / "plan_grafik" / f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    )
    session_dir.mkdir(parents=True, exist_ok=True)
    try:
        parsed = parse_plan_grafik(target_path)
    except Exception as e:
        error_msg = f"Ошибка парсинга: {e}"
        (session_dir / "ERROR.txt").write_text(error_msg, encoding="utf-8")
        return AuditResult(
            violations=[], doc_type="plan_grafik", session_dir=session_dir,
            duration_sec=time.time() - start_time, rules_checked=0, target_path=target_path,
        )
    with open(session_dir / "parsed.json", "w", encoding="utf-8") as f:
        json.dump(parsed, f, ensure_ascii=False, indent=2, default=str)
    if getattr(args, "parse_only", False):
        return AuditResult(
            doc_type="plan_grafik", session_dir=session_dir,
            duration_sec=time.time() - start_time, target_path=target_path,
        )
    # Прогон 9 правил по одному: PASS / FAIL / ERROR + per-rule validate_N.json (формат special-типов).
    validation_outputs_dir = session_dir / "validation_outputs"
    validation_outputs_dir.mkdir(parents=True, exist_ok=True)
    results: List[Tuple[Dict, Dict, float]] = []
    for rule in _RULES:
        t0 = time.time()
        try:
            viols = rule["fn"](parsed, target_path)
            if viols:
                status = "FAIL"
                # discrepancy = сводка нарушений (ячейка/маркер — суть проблемы).
                discrepancy = "; ".join(
                    " — ".join(p for p in (v.get("Целевой документ", ""), v.get("Различие", "")) if p)
                    for v in viols
                )
            else:
                status = "PASS"
                discrepancy = ""
        except Exception as e:
            status, discrepancy = "ERROR", f"Исключение: {e}"
        dur = time.time() - t0
        res = {
            "rule_index": rule["rule_index"], "rule_title": rule["rule_title"],
            "status": status, "discrepancy": discrepancy,
        }
        results.append((rule, res, dur))
        with open(validation_outputs_dir / f"validate_{rule['rule_index']}.json", "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=2)

    _create_excel_report(results, session_dir / "validation_report.xlsx")

    violations = [
        {"rule_index": res["rule_index"], "rule_title": res["rule_title"], "Различие": res["discrepancy"]}
        for _, res, _ in results if res["status"] == "FAIL"
    ]
    return AuditResult(
        violations=violations, doc_type="plan_grafik", session_dir=session_dir,
        duration_sec=time.time() - start_time, rules_checked=len(_RULES), target_path=target_path,
    )
# END_RUNNER
