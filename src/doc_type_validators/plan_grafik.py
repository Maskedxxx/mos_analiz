# START_MODULE_CONTRACT
# PURPOSE: Аудит документа «План-график» (xlsx). 9 детерминистских non-LLM проверок (имя файла, блок УТВЕРЖДАЮ, заголовки, даты, ответственные, расчётные колонки, подпись, формулы) + публичный entrypoint `run_plan_grafik_special` для регистрации в `SPECIAL_ENGINE_RUNNERS`.
# INPUTS: xlsx-файл; парсинг через `parse_plan_grafik` (src/doc_type_parsers/plan_grafik.py); конфиг отсутствует (правила захардкожены в валидаторах).
# OUTPUTS: `run_plan_grafik_special(args) -> AuditResult`. Также пишет в session_dir: `parsed.json` + `validation_report.xlsx`.
# KEYWORDS: validators, plan-grafik, non-llm, deterministic, runner.
# LINKS: src/doc_type_parsers/plan_grafik.py (parse_plan_grafik), main.py (save_to_excel + AuditResult — лениво импортируются).
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
from typing import Any, Dict, List

from src.audit.excel_reporter import save_to_excel
from src.audit.models import AuditResult
from src.doc_type_parsers.plan_grafik import parse_plan_grafik
# END_IMPORTS


# START_PATHS
# PURPOSE: parents[2] = repo root (файл лежит в src/doc_type_validators/).
_LOGS_RESULT_DIR = Path(__file__).resolve().parents[2] / "logs_result"
# END_PATHS


def run_all_validators(parsed: Dict[str, Any], target_path: str) -> List[Dict[str, Any]]:
    """
    Запускает все 9 валидаторов последовательно.

    Args:
        parsed: результат парсинга из parser.py
        target_path: путь к файлу (для проверки имени)

    Returns:
        Список нарушений
    """
    violations = []
    validators = [validate_1_filename, validate_2_approval, validate_3_simple_headers, validate_4_complex_headers, validate_5_dates, validate_6_responsible, validate_7_calc_columns, validate_8_signature, validate_9_formulas]
    for validator in validators:
        try:
            result = validator(parsed, target_path)
            violations.extend(result)
        except Exception as e:
            violations.append({'rule_index': 0, 'rule_title': f'Ошибка валидатора {validator.__name__}', 'Целевой документ': str(e), 'Различие': 'Внутренняя ошибка валидатора'})
    return violations

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
        4. run_all_validators(parsed, target_path) → нарушения.
        5. save_to_excel — отчёт.
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
    violations = run_all_validators(parsed, target_path)
    save_to_excel(violations, str(session_dir / "validation_report.xlsx"))
    return AuditResult(
        violations=violations, doc_type="plan_grafik", session_dir=session_dir,
        duration_sec=time.time() - start_time, rules_checked=9, target_path=target_path,
    )
# END_RUNNER
