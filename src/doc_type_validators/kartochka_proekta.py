# START_MODULE_CONTRACT
# PURPOSE: Валидаторы документа «Карточка проекта». 11 проверок правил 1–11 по распарсенным JSON-ам из `src/doc_type_parsers/kartochka_proekta.py`.
# INPUTS: parser_outputs_dir с JSON-ами (kartochka_main.json, metodika.json, dropdown_units.json); rule из validation_rules.json (путь через os.environ['VALIDATION_RULES_PATH']); LLM-сервис через src.llm.client.call_llm (для правил 3 и 6 — семантические проверки).
# OUTPUTS: dict {rule_index, rule_title, status=PASS/FAIL, discrepancy}. Публичные символы: `KARTOCHKA_VALIDATOR_MODULE_DISPATCH`, `run_kartochka_validator_module`.
# KEYWORDS: validators, kartochka-proekta, llm, json-verdict.
# LINKS: src/llm/client.py (call_llm), config/llm.py (LLM_CONFIG), src/doc_type_parsers/kartochka_proekta.py (источник данных), doc_configs/kartochka_proekta/validation_rules.json.
# RATIONALE:
#   Интерфейс валидатора отличается от KPSC: `validate(*data_args, rule, api_key)`
#   возвращает вердикт одним вызовом (вместо extract → build_prompt → call_llm →
#   save_result). Часть правил — non-LLM (filename, даты, числа), часть — LLM
#   (семантика названий/обоснований). Раннер `run_kartochka_validator_module`
#   использует inspect для подбора аргументов под сигнатуру `validate`.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import inspect
import json
import os
import re
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict, List

from config.llm import LLM_CONFIG
from src.llm.client import call_llm
# END_IMPORTS


# START_SHARED_HELPERS
# PURPOSE: Общие хелперы для всех 11 валидаторов карточки проекта. Убирают дубли load_rule/load_data/JSON-IO и собирают SimpleNamespace через фабрику.
KARTOCHKA_VALIDATION_RULES_PATH = (
    Path(__file__).resolve().parents[2]
    / "doc_configs"
    / "kartochka_proekta"
    / "validation_rules.json"
)


def _get_validation_rules_path() -> Path:
    """Путь к validation_rules.json (env override через VALIDATION_RULES_PATH)."""
    return Path(os.environ.get("VALIDATION_RULES_PATH", str(KARTOCHKA_VALIDATION_RULES_PATH)))


def _read_json_file(path: Path) -> Any:
    """Читает JSON-файл в UTF-8."""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _write_json_file(path: Path, data: Dict[str, Any]) -> None:
    """Пишет JSON-файл в UTF-8 с indent=2."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _load_rule(rule_index: str) -> dict:
    """Загружает правило проверки из validation_rules.json по rule_index."""
    data = _read_json_file(_get_validation_rules_path())
    for rule in data["rules"]:
        if rule["rule_index"] == rule_index:
            return rule
    raise ValueError(f"Правило {rule_index} не найдено")


def _load_parser_json(parser_outputs_dir: Path, filename: str) -> Any:
    """Загружает один JSON парсера из parser_outputs_dir."""
    return _read_json_file(parser_outputs_dir / filename)


def _load_parser_jsons(parser_outputs_dir: Path, filenames: tuple[str, ...]) -> tuple:
    """Загружает несколько JSON-ов парсера в порядке filenames."""
    return tuple(_load_parser_json(parser_outputs_dir, filename) for filename in filenames)


def _build_validate_args(
    validate_fn: Any,
    loaded_data: Any,
    rule: dict,
    api_key: str,
    rule_index: str,
) -> List[Any]:
    """Собирает позиционные аргументы validate(...) по сигнатуре конкретного валидатора."""
    pending_data = list(loaded_data) if isinstance(loaded_data, tuple) else [loaded_data]
    call_args: List[Any] = []
    for param in inspect.signature(validate_fn).parameters.values():
        if param.name == "rule":
            call_args.append(rule)
        elif param.name == "api_key":
            call_args.append(api_key)
        else:
            if not pending_data:
                raise ValueError(f"Недостаточно данных для валидатора {rule_index}")
            call_args.append(pending_data.pop(0))
    return call_args


def _make_validator_module(
    rule_index: str,
    rule_title: str,
    load_data_fn: Any,
    validate_fn: Any,
    **extra_attrs: Any,
) -> SimpleNamespace:
    """Собирает SimpleNamespace валидатора с обязательным контрактом {RULE_INDEX, RULE_TITLE, load_rule, load_data, validate}."""
    return SimpleNamespace(
        RULE_INDEX=rule_index,
        RULE_TITLE=rule_title,
        load_rule=_load_rule,
        load_data=load_data_fn,
        validate=validate_fn,
        **extra_attrs,
    )
# END_SHARED_HELPERS


# ==============================================================================
# RULE 1 — Название файла
# ==============================================================================
_RULE_1_INDEX = '1'
_RULE_1_TITLE = 'Название файла'

def _rule_1_load_data(parser_outputs_dir: Path) -> Any:
    """Загрузка JSON-ов парсера для правила 1."""
    return _load_parser_json(parser_outputs_dir, "kartochka_main.json")

def _rule_1_validate(data: dict, rule: dict) -> dict:
    """Проверка названия файла.

    Проверяет наличие ключевых слов «Карточка» и «проект» в имени файла.
    Наименование предприятия НЕ проверяется — оно динамическое.
    """
    filename = data.get('meta', {}).get('workbook', '')
    filename_lower = filename.lower()
    keywords = ['карточка', 'проект']
    missing = [kw for kw in keywords if kw not in filename_lower]
    if missing:
        return {'rule_index': _RULE_1_INDEX, 'rule_title': _RULE_1_TITLE, 'status': 'FAIL', 'discrepancy': f"Имя файла '{filename}' не содержит ключевые слова: {missing}"}
    return {'rule_index': _RULE_1_INDEX, 'rule_title': _RULE_1_TITLE, 'status': 'PASS', 'discrepancy': ''}

RULE_1 = _make_validator_module(
    rule_index='1',
    rule_title='Название файла',
    load_data_fn=_rule_1_load_data,
    validate_fn=_rule_1_validate,
)


# ==============================================================================
# RULE 2 — Вид организации и название предприятия
# ==============================================================================
_RULE_2_INDEX = '2'
_RULE_2_TITLE = 'Вид организации и название предприятия'

def _rule_2_load_data(parser_outputs_dir: Path) -> Any:
    """Загрузка JSON-ов парсера для правила 2."""
    return _load_parser_json(parser_outputs_dir, "kartochka_main.json")

def _rule_2_validate(data: dict, rule: dict) -> dict:
    """Проверка юридической формы и названия предприятия."""
    org_name = data.get('header', {}).get('org_name', '')
    errors = []
    if not org_name:
        errors.append('Поле org_name (B2) пустое — нет названия организации')
    else:
        jur_form_match = re.search('\\b(ООО|ЗАО|АО|ПАО|ОАО|ИП)\\b', org_name)
        if not jur_form_match:
            errors.append(f"Юридическая форма (ООО/ЗАО/АО/ПАО/ОАО/ИП) не найдена в '{org_name}'")
        name_match = re.search('["\\«](.+?)["\\»]', org_name)
        if not name_match or not name_match.group(1).strip():
            after_form = re.sub('^(ООО|ЗАО|АО|ПАО|ОАО|ИП)\\s*', '', org_name).strip()
            after_form = after_form.strip('"«»\'" ')
            if not after_form:
                errors.append('Наименование предприятия пустое после юридической формы')
    if errors:
        return {'rule_index': _RULE_2_INDEX, 'rule_title': _RULE_2_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': _RULE_2_INDEX, 'rule_title': _RULE_2_TITLE, 'status': 'PASS', 'discrepancy': ''}

RULE_2 = _make_validator_module(
    rule_index='2',
    rule_title='Вид организации и название предприятия',
    load_data_fn=_rule_2_load_data,
    validate_fn=_rule_2_validate,
)


# ==============================================================================
# RULE 3 — Название потока
# ==============================================================================
_RULE_3_INDEX = '3'
_RULE_3_TITLE = 'Название потока'

def _rule_3_load_data(parser_outputs_dir: Path) -> Any:
    """Загрузка JSON-ов парсера для правила 3."""
    return _load_parser_json(parser_outputs_dir, "kartochka_main.json")

def _rule_3_check_semantic(project_name: str, rule: dict, api_key: str) -> dict:
    """LLM-проверка осмысленности названия проекта/потока (через единый клиент)."""
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    prompt = f'Ты эксперт по проверке документов «Карточка проекта» в рамках бережливого производства.\n\nТРЕБОВАНИЕ:\n{rule['requirement_expert']}\n\nКРИТЕРИИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nНазвание проекта/потока: "{project_name}"\n\nЗАДАНИЕ:\nПроверь, является ли название проекта осмысленным текстом, описывающим реальный проект или поток.\nНе является осмысленным: placeholder ("Название проекта"), набор символов ("ааааа"), слишком общий текст ("тест").\nЯвляется осмысленным: конкретное описание проекта ("Оптимизация производства приборов учёта").\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{_RULE_3_INDEX}",\n  "rule_title": "{_RULE_3_TITLE}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n'
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def _rule_3_validate(data: dict, rule: dict, api_key: str) -> dict:
    """Проверка названия проекта/потока без сетевой зависимости."""
    project_name = data.get('header', {}).get('project_name', '')
    if not project_name:
        return {'rule_index': _RULE_3_INDEX, 'rule_title': _RULE_3_TITLE, 'status': 'FAIL', 'discrepancy': 'Поле project_name (B4) пустое — нет названия проекта/потока'}
    project_name = project_name.strip()
    if len(project_name) <= 5:
        return {'rule_index': _RULE_3_INDEX, 'rule_title': _RULE_3_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта слишком короткое ({len(project_name)} симв.): '{project_name}'"}
    normalized = ' '.join(project_name.lower().split())
    placeholder_values = {'название проекта', 'название потока', 'проект', 'поток', 'тест', 'test', 'aaaaa', 'aaaa', 'qwerty'}
    letters_only = ''.join((ch for ch in normalized if ch.isalpha()))
    if normalized in placeholder_values:
        return {'rule_index': _RULE_3_INDEX, 'rule_title': _RULE_3_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта выглядит как placeholder: '{project_name}'"}
    if not letters_only:
        return {'rule_index': _RULE_3_INDEX, 'rule_title': _RULE_3_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта не содержит осмысленного текста: '{project_name}'"}
    unique_letters = set(letters_only)
    if len(unique_letters) <= 2:
        return {'rule_index': _RULE_3_INDEX, 'rule_title': _RULE_3_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта похоже на набор повторяющихся символов: '{project_name}'"}
    return {'rule_index': _RULE_3_INDEX, 'rule_title': _RULE_3_TITLE, 'status': 'PASS', 'discrepancy': ''}

RULE_3 = _make_validator_module(
    rule_index='3',
    rule_title='Название потока',
    load_data_fn=_rule_3_load_data,
    validate_fn=_rule_3_validate,
    check_semantic=_rule_3_check_semantic,
)


# ==============================================================================
# RULE 4 — Должность, ФИО подписанта, дата и подпись
# ==============================================================================
_RULE_4_INDEX = '4'
_RULE_4_TITLE = 'Должность, ФИО подписанта, дата и подпись'

def _rule_4_load_data(parser_outputs_dir: Path) -> Any:
    """Загрузка JSON-ов парсера для правила 4."""
    return _load_parser_json(parser_outputs_dir, "kartochka_main.json")

def _rule_4_validate(data: dict) -> dict:
    """Проверка подписанта."""
    header = data.get('header', {})
    errors = []
    position = header.get('signee_position', '')
    if not position:
        errors.append('Должность подписанта (K4) не заполнена')
    name = header.get('signee_name', '')
    if not name:
        errors.append('ФИО подписанта (K7) не заполнено')
    else:
        words = re.findall('[А-Яа-яЁёA-Za-z]+\\.?', name)
        if len(words) < 2:
            errors.append(f"ФИО подписанта содержит менее 2 слов: '{name}'")
    date_val = header.get('signee_date', '')
    if not date_val:
        errors.append('Дата подписания (K8) не заполнена')
    elif '___' in date_val or '20__' in date_val:
        errors.append(f"Дата подписания содержит незаполненные placeholder: '{date_val}'")
    if errors:
        return {'rule_index': _RULE_4_INDEX, 'rule_title': _RULE_4_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': _RULE_4_INDEX, 'rule_title': _RULE_4_TITLE, 'status': 'PASS', 'discrepancy': ''}

RULE_4 = _make_validator_module(
    rule_index='4',
    rule_title='Должность, ФИО подписанта, дата и подпись',
    load_data_fn=_rule_4_load_data,
    validate_fn=_rule_4_validate,
)


# ==============================================================================
# RULE 5 — Обязательные поля секции 1
# ==============================================================================
_RULE_5_INDEX = '5'
_RULE_5_TITLE = 'Обязательные поля секции 1'
_RULE_5_REQUIRED_FIELDS = {'clients': 'Клиенты процесса', 'perimeter': 'Периметр проекта', 'owner': 'Владелец процесса', 'boundaries': 'Границы процесса', 'leader': 'Руководитель проекта', 'team': 'Команда проекта'}
_RULE_5_FIO_FIELDS = ['owner', 'leader', 'team']

def _rule_5_load_data(parser_outputs_dir: Path) -> Any:
    """Загрузка JSON-ов парсера для правила 5."""
    return _load_parser_json(parser_outputs_dir, "kartochka_main.json")

def _rule_5_check_fio_format(text: str) -> bool:
    """Проверяет наличие паттерна 'ФИО - должность' (разделитель — тире)."""
    return bool(re.search('.+\\s*[-–—]\\s*.+', text))

def _rule_5_validate(data: dict) -> dict:
    """Проверка обязательных полей секции 1."""
    section1 = data.get('section1', {})
    errors = []
    empty_fields = []
    for field_key, field_name in _RULE_5_REQUIRED_FIELDS.items():
        val = section1.get(field_key, '')
        if not val or not val.strip():
            empty_fields.append(field_name)
    if empty_fields:
        errors.append(f'Пустые поля: {', '.join(empty_fields)}')
    bad_format = []
    for field_key in _RULE_5_FIO_FIELDS:
        val = section1.get(field_key, '')
        if not val:
            continue
        if field_key == 'team':
            members = [m.strip() for m in val.split(',') if m.strip()]
            for member in members:
                if not _rule_5_check_fio_format(member):
                    bad_format.append(f"team: '{member[:50]}'")
                    break
        elif not _rule_5_check_fio_format(val):
            bad_format.append(f"{_RULE_5_REQUIRED_FIELDS[field_key]}: '{val[:50]}'")
    if bad_format:
        errors.append(f'Неверный формат ФИО-должность: {'; '.join(bad_format)}')
    if errors:
        return {'rule_index': _RULE_5_INDEX, 'rule_title': _RULE_5_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': _RULE_5_INDEX, 'rule_title': _RULE_5_TITLE, 'status': 'PASS', 'discrepancy': ''}

RULE_5 = _make_validator_module(
    rule_index='5',
    rule_title='Обязательные поля секции 1',
    load_data_fn=_rule_5_load_data,
    validate_fn=_rule_5_validate,
    _check_fio_format=_rule_5_check_fio_format,
    REQUIRED_FIELDS=_RULE_5_REQUIRED_FIELDS,
    FIO_FIELDS=_RULE_5_FIO_FIELDS,
)


# ==============================================================================
# RULE 6 — Обоснование и ключевой риск
# ==============================================================================
_RULE_6_INDEX = '6'
_RULE_6_TITLE = 'Обоснование и ключевой риск'

def _rule_6_load_data(parser_outputs_dir: Path) -> Any:
    """Загрузка JSON-ов парсера для правила 6."""
    return _load_parser_json(parser_outputs_dir, "kartochka_main.json")

def _rule_6_check_semantic(key_risk: str, justification: str, rule: dict, api_key: str) -> dict:
    """LLM-проверка осмысленности обоснования и ключевого риска (через единый клиент)."""
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    prompt = f'Ты эксперт по проверке документов «Карточка проекта» в рамках бережливого производства.\n\nТРЕБОВАНИЕ:\n{rule['requirement_expert']}\n\nКРИТЕРИИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nКлючевой риск (M11): "{key_risk}"\nОбоснование выбора потока (M13): "{justification}"\n\nЗАДАНИЕ:\nПроверь, содержат ли оба поля осмысленный текст:\n- Ключевой риск — должен описывать конкретный риск проекта (например: "Срыв сроков", "Потеря клиентов").\n  НЕ осмысленный: "-", "нет", "риск", набор символов.\n- Обоснование — должно содержать аргументацию выбора потока (например: "Наличие ожидания в потоке, несвоевременная подготовка").\n  НЕ осмысленный: "-", "обоснование", "тест", набор символов.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{_RULE_6_INDEX}",\n  "rule_title": "{_RULE_6_TITLE}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n'
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def _rule_6_validate(data: dict, rule: dict, api_key: str) -> dict:
    """Проверка обоснования и ключевого риска: non-LLM + LLM."""
    section2 = data.get('section2', {})
    errors = []
    key_risk = section2.get('key_risk', '')
    justification = section2.get('justification', '')
    if not key_risk or not key_risk.strip():
        errors.append('Ключевой риск (M11) не заполнен')
    if not justification or not justification.strip():
        errors.append('Обоснование выбора потока (M13) не заполнено')
    if errors:
        return {'rule_index': _RULE_6_INDEX, 'rule_title': _RULE_6_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return _rule_6_check_semantic(key_risk, justification, rule, api_key)

RULE_6 = _make_validator_module(
    rule_index='6',
    rule_title='Обоснование и ключевой риск',
    load_data_fn=_rule_6_load_data,
    validate_fn=_rule_6_validate,
    check_semantic=_rule_6_check_semantic,
)


# ==============================================================================
# RULE 7 — Даты мероприятий
# ==============================================================================
_RULE_7_INDEX = '7'
_RULE_7_TITLE = 'Даты мероприятий'

def _rule_7_load_data(parser_outputs_dir: Path) -> Any:
    """Загрузка JSON-ов парсера для правила 7."""
    return _load_parser_json(parser_outputs_dir, "kartochka_main.json")

def _rule_7_is_stage_event(name: str) -> bool:
    """Проверяет, является ли событие этапным (начинается с '1.', '2.', '3.', '4.')."""
    return bool(re.match('^\\d+\\.', name.strip()))

def _rule_7_validate(data: dict) -> dict:
    """Проверка дат мероприятий."""
    events = data.get('events', [])
    errors = []
    if not events:
        return {'rule_index': _RULE_7_INDEX, 'rule_title': _RULE_7_TITLE, 'status': 'FAIL', 'discrepancy': 'Список мероприятий (events) пуст'}
    for i, event in enumerate(events):
        name = event.get('name', f'Событие #{i + 1}')
        start_date = event.get('start_date')
        end_date = event.get('end_date')
        if not start_date:
            errors.append(f"'{name[:40]}' — нет даты начала")
        if _rule_7_is_stage_event(name) and (not end_date):
            errors.append(f"'{name[:40]}' — этапное событие без даты окончания")
    if errors:
        return {'rule_index': _RULE_7_INDEX, 'rule_title': _RULE_7_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': _RULE_7_INDEX, 'rule_title': _RULE_7_TITLE, 'status': 'PASS', 'discrepancy': ''}

RULE_7 = _make_validator_module(
    rule_index='7',
    rule_title='Даты мероприятий',
    load_data_fn=_rule_7_load_data,
    validate_fn=_rule_7_validate,
    _is_stage_event=_rule_7_is_stage_event,
)


# ==============================================================================
# RULE 8 — Показатели и даты
# ==============================================================================
_RULE_8_INDEX = '8'
_RULE_8_TITLE = 'Показатели и даты'

def _rule_8_load_data(parser_outputs_dir: Path) -> Any:
    """Загрузка JSON-ов парсера для правила 8."""
    return _load_parser_json(parser_outputs_dir, "kartochka_main.json")

def _rule_8_validate(data: dict) -> dict:
    """Проверка показателей и дат."""
    indicators = data.get('indicators', [])
    indicator_dates = data.get('indicator_dates', {})
    errors = []
    if not indicators:
        errors.append('Нет показателей (массив indicators пуст)')
    else:
        required_fields = ['name', 'unit', 'base_value', 'target_value', 'ideal_value']
        for ind in indicators:
            num = ind.get('number', '?')
            missing = []
            for field in required_fields:
                val = ind.get(field)
                if val is None or (isinstance(val, str) and (not val.strip())):
                    missing.append(field)
            if missing:
                errors.append(f'Показатель #{num}: пустые поля [{', '.join(missing)}]')
    date_fields = {'base_date': 'Дата базы', 'target_date': 'Дата цели', 'ideal_date': 'Дата идеала'}
    for field, label in date_fields.items():
        val = indicator_dates.get(field)
        if not val:
            errors.append(f'{label} не заполнена')
    if errors:
        return {'rule_index': _RULE_8_INDEX, 'rule_title': _RULE_8_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': _RULE_8_INDEX, 'rule_title': _RULE_8_TITLE, 'status': 'PASS', 'discrepancy': ''}

RULE_8 = _make_validator_module(
    rule_index='8',
    rule_title='Показатели и даты',
    load_data_fn=_rule_8_load_data,
    validate_fn=_rule_8_validate,
)


# ==============================================================================
# RULE 9 — Единицы измерения (карточка)
# ==============================================================================
_RULE_9_INDEX = '9'
_RULE_9_TITLE = 'Единицы измерения (карточка)'
_RULE_9_CATEGORY_KEYWORDS = {'Время протекания процесса': ['время', 'протекан'], 'Выработка': ['выработк'], 'Незавершенное производство': ['запас', 'нзп', 'незавершен']}
_RULE_9_DEFAULT_CATEGORY = 'Дополнительный показатель'

def _rule_9_load_data(parser_outputs_dir: Path) -> Any:
    """Загрузка JSON-ов парсера для правила 9."""
    return _load_parser_jsons(parser_outputs_dir, ("kartochka_main.json", "dropdown_units.json",))

def _rule_9_detect_category(indicator_name: str) -> str:
    """Определяет категорию показателя по ключевым словам в названии."""
    name_lower = indicator_name.lower()
    for category, keywords in _RULE_9_CATEGORY_KEYWORDS.items():
        if any((kw in name_lower for kw in keywords)):
            return category
    return _RULE_9_DEFAULT_CATEGORY

def _rule_9_validate(kartochka: dict, dropdown: dict) -> dict:
    """Проверка единиц измерения показателей по справочнику."""
    indicators = kartochka.get('indicators', [])
    categories = dropdown.get('categories', {})
    errors = []
    all_units = set()
    all_units_display = []
    for units_list in categories.values():
        for u in units_list:
            normalized = u.lower().strip().rstrip('.')
            all_units.add(normalized)
            all_units_display.append(u)
    for ind in indicators:
        name = ind.get('name', '')
        unit = ind.get('unit', '')
        num = ind.get('number', '?')
        if not unit:
            errors.append(f"Показатель #{num} '{name}': единица измерения не заполнена")
            continue
        unit_normalized = unit.lower().strip().rstrip('.')
        if unit_normalized not in all_units:
            errors.append(f"Показатель #{num} '{name}': единица '{unit}' не найдена в справочнике (допустимые: {', '.join(all_units_display[:10])}...)")
    if errors:
        return {'rule_index': _RULE_9_INDEX, 'rule_title': _RULE_9_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': _RULE_9_INDEX, 'rule_title': _RULE_9_TITLE, 'status': 'PASS', 'discrepancy': ''}

RULE_9 = _make_validator_module(
    rule_index='9',
    rule_title='Единицы измерения (карточка)',
    load_data_fn=_rule_9_load_data,
    validate_fn=_rule_9_validate,
    _detect_category=_rule_9_detect_category,
    CATEGORY_KEYWORDS=_RULE_9_CATEGORY_KEYWORDS,
    DEFAULT_CATEGORY=_RULE_9_DEFAULT_CATEGORY,
)


# ==============================================================================
# RULE 10 — Единицы измерения (методика расчета)
# ==============================================================================
_RULE_10_INDEX = '10'
_RULE_10_TITLE = 'Единицы измерения (методика расчета)'

def _rule_10_load_data(parser_outputs_dir: Path) -> Any:
    """Загрузка JSON-ов парсера для правила 10."""
    return _load_parser_jsons(parser_outputs_dir, ("kartochka_main.json", "metodika.json",))

def _rule_10_normalize_name(name: str) -> str:
    """Нормализация названия показателя для сопоставления."""
    import re
    name = re.sub('\\(.*?\\)', '', name)
    name = name.replace(':', '').strip().lower()
    return name

def _rule_10_validate(kartochka: dict, metodika: dict) -> dict:
    """Проверка совпадения единиц измерения между методикой и карточкой."""
    k_indicators = kartochka.get('indicators', [])
    m_indicators = metodika.get('indicators', [])
    errors = []
    k_units = {}
    for ind in k_indicators:
        name = ind.get('name', '')
        if name:
            k_units[_rule_10_normalize_name(name)] = ind.get('unit', '')
    for m_ind in m_indicators:
        m_name = m_ind.get('name', '')
        m_unit = m_ind.get('unit')
        if not m_unit:
            m_norm = _rule_10_normalize_name(m_name)
            if m_norm in k_units and k_units[m_norm]:
                errors.append(f"Показатель '{m_name}': единица в методике не заполнена, в карточке — '{k_units[m_norm]}'")
            continue
        m_norm = _rule_10_normalize_name(m_name)
        k_unit = k_units.get(m_norm)
        if k_unit is None:
            continue
        if m_unit.lower().strip() != k_unit.lower().strip():
            errors.append(f"Показатель '{m_name}': единица в методике '{m_unit}' ≠ единица в карточке '{k_unit}'")
    if errors:
        return {'rule_index': _RULE_10_INDEX, 'rule_title': _RULE_10_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': _RULE_10_INDEX, 'rule_title': _RULE_10_TITLE, 'status': 'PASS', 'discrepancy': ''}

RULE_10 = _make_validator_module(
    rule_index='10',
    rule_title='Единицы измерения (методика расчета)',
    load_data_fn=_rule_10_load_data,
    validate_fn=_rule_10_validate,
    _normalize_name=_rule_10_normalize_name,
)


# ==============================================================================
# RULE 11 — Способ расчёта и источник данных
# ==============================================================================
_RULE_11_INDEX = '11'
_RULE_11_TITLE = 'Способ расчёта и источник данных'

def _rule_11_load_data(parser_outputs_dir: Path) -> Any:
    """Загрузка JSON-ов парсера для правила 11."""
    return _load_parser_json(parser_outputs_dir, "metodika.json")

def _rule_11_validate(data: dict) -> dict:
    """Проверка способа расчёта и источника данных."""
    indicators = data.get('indicators', [])
    errors = []
    for ind in indicators:
        name = ind.get('name', '')
        unit = ind.get('unit')
        calc_method = ind.get('calc_method')
        data_source = ind.get('data_source')
        if not unit:
            continue
        missing = []
        if not calc_method or not str(calc_method).strip():
            missing.append('способ расчёта')
        if not data_source or not str(data_source).strip():
            missing.append('источник данных')
        if missing:
            errors.append(f"Показатель '{name}': не заполнено — {', '.join(missing)}")
    if errors:
        return {'rule_index': _RULE_11_INDEX, 'rule_title': _RULE_11_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': _RULE_11_INDEX, 'rule_title': _RULE_11_TITLE, 'status': 'PASS', 'discrepancy': ''}

RULE_11 = _make_validator_module(
    rule_index='11',
    rule_title='Способ расчёта и источник данных',
    load_data_fn=_rule_11_load_data,
    validate_fn=_rule_11_validate,
)


# START_VALIDATOR_DISPATCH
# PURPOSE: Маппинг rule_index → SimpleNamespace валидатора (1–11). Используется run_kartochka_validator_module для выбора валидатора по ID правила.
KARTOCHKA_VALIDATOR_MODULE_DISPATCH: Dict[str, SimpleNamespace] = {
    '1': RULE_1,
    '2': RULE_2,
    '3': RULE_3,
    '4': RULE_4,
    '5': RULE_5,
    '6': RULE_6,
    '7': RULE_7,
    '8': RULE_8,
    '9': RULE_9,
    '10': RULE_10,
    '11': RULE_11,
}
# END_VALIDATOR_DISPATCH


# START_VALIDATOR_RUNNER
# PURPOSE: Универсальный runner одного валидатора карточки проекта. Сигнатура validate(...) варьируется — inspect подбирает аргументы.
def run_kartochka_validator_module(module_ns: SimpleNamespace, parser_outputs_dir: Path, output_file: Path) -> Dict[str, Any]:
    """
    Назначение:
        Выполняет стандартный пайплайн: load_rule → load_data → validate.
        Сигнатура validate(...) специфична для каждого правила — inspect
        позволяет собрать аргументы автоматически.

    Вход:
        module_ns: SimpleNamespace валидатора (из KARTOCHKA_VALIDATOR_MODULE_DISPATCH).
        parser_outputs_dir: Директория с JSON-ами парсера kartochka_proekta.
        output_file: Путь для записи JSON с результатом.

    Выход:
        dict {rule_index, rule_title, status=PASS/FAIL, discrepancy}. Также
        пишет output_file.
    """
    rule = module_ns.load_rule(module_ns.RULE_INDEX)
    loaded_data = module_ns.load_data(parser_outputs_dir)
    api_key = os.environ.get("OPENAI_API_KEY", "dummy")
    call_args = _build_validate_args(module_ns.validate, loaded_data, rule, api_key, module_ns.RULE_INDEX)
    result = module_ns.validate(*call_args)
    _write_json_file(output_file, result)
    return result
# END_VALIDATOR_RUNNER
