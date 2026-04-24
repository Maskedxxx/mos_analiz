# START_MODULE_CONTRACT
# PURPOSE: Валидаторы документа «Карточка проекта». 11 проверок правил 1–11 по распарсенным JSON-ам из `src/doc_type_parsers/kartochka_proekta.py`.
# INPUTS: parser_outputs_dir с JSON-ами (kartochka_main.json, metodika.json, dropdown.json); rule из validation_rules.json (путь через os.environ['VALIDATION_RULES_PATH']); LLM-сервис через src.llm.client.call_llm (для правил 3 и 6 — семантические проверки).
# OUTPUTS: dict {rule_index, rule_title, status=PASS/FAIL, discrepancy}. Публичные символы: `KARTOCHKA_VALIDATOR_MODULE_DISPATCH`, `run_kartochka_validator_module`.
# KEYWORDS: validators, kartochka-proekta, llm, json-verdict.
# LINKS: src/llm/client.py (call_llm), config/llm.py (LLM_CONFIG), src/doc_type_parsers/kartochka_proekta.py (источник данных), doc_configs/kartochka_proekta/validation_rules.json.
# RATIONALE:
#   Интерфейс валидатора отличается от KPSC: `validate(*data_args, rule, api_key)`
#   возвращает верди кт одним вызовом (вместо extract → build_prompt → call_llm →
#   save_result). Часть правил — non-LLM (filename, daты, числа), часть — LLM
#   (семантика названий/обоснований). Раннер `run_kartochka_validator_module`
#   использует inspect для подбора аргументов под сигнатуру `validate`.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import argparse
import json
import os
import re
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict, List

from config.llm import LLM_CONFIG
from src.llm.client import call_llm
# END_IMPORTS

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_10_UNITS_METODIKA
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_10_units_metodika.py.
kartochka_proekta_validate_10_units_metodika_RULE_INDEX = '10'

kartochka_proekta_validate_10_units_metodika_RULE_TITLE = 'Единицы измерения (методика расчета)'

def kartochka_proekta_validate_10_units_metodika_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_10_units_metodika_load_data(parser_outputs_dir: Path) -> tuple:
    """Загрузка kartochka_main.json и metodika.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        kartochka = json.load(f)
    with open(parser_outputs_dir / 'metodika.json', 'r', encoding='utf-8') as f:
        metodika = json.load(f)
    return (kartochka, metodika)

def kartochka_proekta_validate_10_units_metodika__normalize_name(name: str) -> str:
    """Нормализация названия показателя для сопоставления."""
    import re
    name = re.sub('\\(.*?\\)', '', name)
    name = name.replace(':', '').strip().lower()
    return name

def kartochka_proekta_validate_10_units_metodika_validate(kartochka: dict, metodika: dict) -> dict:
    """Проверка совпадения единиц измерения между методикой и карточкой."""
    k_indicators = kartochka.get('indicators', [])
    m_indicators = metodika.get('indicators', [])
    errors = []
    k_units = {}
    for ind in k_indicators:
        name = ind.get('name', '')
        if name:
            k_units[kartochka_proekta_validate_10_units_metodika__normalize_name(name)] = ind.get('unit', '')
    for m_ind in m_indicators:
        m_name = m_ind.get('name', '')
        m_unit = m_ind.get('unit')
        if not m_unit:
            m_norm = kartochka_proekta_validate_10_units_metodika__normalize_name(m_name)
            if m_norm in k_units and k_units[m_norm]:
                errors.append(f"Показатель '{m_name}': единица в методике не заполнена, в карточке — '{k_units[m_norm]}'")
            continue
        m_norm = kartochka_proekta_validate_10_units_metodika__normalize_name(m_name)
        k_unit = k_units.get(m_norm)
        if k_unit is None:
            continue
        if m_unit.lower().strip() != k_unit.lower().strip():
            errors.append(f"Показатель '{m_name}': единица в методике '{m_unit}' ≠ единица в карточке '{k_unit}'")
    if errors:
        return {'rule_index': kartochka_proekta_validate_10_units_metodika_RULE_INDEX, 'rule_title': kartochka_proekta_validate_10_units_metodika_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_10_units_metodika_RULE_INDEX, 'rule_title': kartochka_proekta_validate_10_units_metodika_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_10_units_metodika_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_10_units_metodika_RULE_INDEX}: {kartochka_proekta_validate_10_units_metodika_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_10_units_metodika_load_rule(kartochka_proekta_validate_10_units_metodika_RULE_INDEX)
    kartochka, metodika = kartochka_proekta_validate_10_units_metodika_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_10_units_metodika_validate(kartochka, metodika)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_10_units_metodika_RULE_INDEX}] {kartochka_proekta_validate_10_units_metodika_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_10_units_metodika_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_10_units_metodika_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_10_units_metodika_RULE_TITLE, load_rule=kartochka_proekta_validate_10_units_metodika_load_rule, load_data=kartochka_proekta_validate_10_units_metodika_load_data, _normalize_name=kartochka_proekta_validate_10_units_metodika__normalize_name, validate=kartochka_proekta_validate_10_units_metodika_validate, main=kartochka_proekta_validate_10_units_metodika_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_10_UNITS_METODIKA

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_11_CALC_METHOD
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_11_calc_method.py.
kartochka_proekta_validate_11_calc_method_RULE_INDEX = '11'

kartochka_proekta_validate_11_calc_method_RULE_TITLE = 'Способ расчёта и источник данных'

def kartochka_proekta_validate_11_calc_method_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_11_calc_method_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка metodika.json."""
    with open(parser_outputs_dir / 'metodika.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_11_calc_method_validate(data: dict) -> dict:
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
        return {'rule_index': kartochka_proekta_validate_11_calc_method_RULE_INDEX, 'rule_title': kartochka_proekta_validate_11_calc_method_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_11_calc_method_RULE_INDEX, 'rule_title': kartochka_proekta_validate_11_calc_method_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_11_calc_method_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_11_calc_method_RULE_INDEX}: {kartochka_proekta_validate_11_calc_method_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_11_calc_method_load_rule(kartochka_proekta_validate_11_calc_method_RULE_INDEX)
    data = kartochka_proekta_validate_11_calc_method_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_11_calc_method_validate(data)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_11_calc_method_RULE_INDEX}] {kartochka_proekta_validate_11_calc_method_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_11_calc_method_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_11_calc_method_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_11_calc_method_RULE_TITLE, load_rule=kartochka_proekta_validate_11_calc_method_load_rule, load_data=kartochka_proekta_validate_11_calc_method_load_data, validate=kartochka_proekta_validate_11_calc_method_validate, main=kartochka_proekta_validate_11_calc_method_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_11_CALC_METHOD

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_1_FILENAME
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_1_filename.py.
kartochka_proekta_validate_1_filename_RULE_INDEX = '1'

kartochka_proekta_validate_1_filename_RULE_TITLE = 'Название файла'

def kartochka_proekta_validate_1_filename_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_1_filename_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_1_filename_validate(data: dict, rule: dict) -> dict:
    """Проверка названия файла.

    Проверяет наличие ключевых слов «Карточка» и «проект» в имени файла.
    Наименование предприятия НЕ проверяется — оно динамическое.
    """
    filename = data.get('meta', {}).get('workbook', '')
    filename_lower = filename.lower()
    keywords = ['карточка', 'проект']
    missing = [kw for kw in keywords if kw not in filename_lower]
    if missing:
        return {'rule_index': kartochka_proekta_validate_1_filename_RULE_INDEX, 'rule_title': kartochka_proekta_validate_1_filename_RULE_TITLE, 'status': 'FAIL', 'discrepancy': f"Имя файла '{filename}' не содержит ключевые слова: {missing}"}
    return {'rule_index': kartochka_proekta_validate_1_filename_RULE_INDEX, 'rule_title': kartochka_proekta_validate_1_filename_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_1_filename_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_1_filename_RULE_INDEX}: {kartochka_proekta_validate_1_filename_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    rule = kartochka_proekta_validate_1_filename_load_rule(kartochka_proekta_validate_1_filename_RULE_INDEX)
    data = kartochka_proekta_validate_1_filename_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_1_filename_validate(data, rule)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_1_filename_RULE_INDEX}] {kartochka_proekta_validate_1_filename_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_1_filename_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_1_filename_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_1_filename_RULE_TITLE, load_rule=kartochka_proekta_validate_1_filename_load_rule, load_data=kartochka_proekta_validate_1_filename_load_data, validate=kartochka_proekta_validate_1_filename_validate, main=kartochka_proekta_validate_1_filename_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_1_FILENAME

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_2_ORG_NAME
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_2_org_name.py.
kartochka_proekta_validate_2_org_name_RULE_INDEX = '2'

kartochka_proekta_validate_2_org_name_RULE_TITLE = 'Вид организации и название предприятия'

def kartochka_proekta_validate_2_org_name_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_2_org_name_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_2_org_name_validate(data: dict, rule: dict) -> dict:
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
        return {'rule_index': kartochka_proekta_validate_2_org_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_2_org_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_2_org_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_2_org_name_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_2_org_name_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_2_org_name_RULE_INDEX}: {kartochka_proekta_validate_2_org_name_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    rule = kartochka_proekta_validate_2_org_name_load_rule(kartochka_proekta_validate_2_org_name_RULE_INDEX)
    data = kartochka_proekta_validate_2_org_name_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_2_org_name_validate(data, rule)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_2_org_name_RULE_INDEX}] {kartochka_proekta_validate_2_org_name_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_2_org_name_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_2_org_name_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_2_org_name_RULE_TITLE, load_rule=kartochka_proekta_validate_2_org_name_load_rule, load_data=kartochka_proekta_validate_2_org_name_load_data, validate=kartochka_proekta_validate_2_org_name_validate, main=kartochka_proekta_validate_2_org_name_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_2_ORG_NAME

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_3_FLOW_NAME
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_3_flow_name.py.
kartochka_proekta_validate_3_flow_name_RULE_INDEX = '3'

kartochka_proekta_validate_3_flow_name_RULE_TITLE = 'Название потока'

def kartochka_proekta_validate_3_flow_name_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_3_flow_name_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_3_flow_name_check_semantic(project_name: str, rule: dict, api_key: str) -> dict:
    """LLM-проверка осмысленности названия проекта/потока (через единый клиент)."""
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    prompt = f'Ты эксперт по проверке документов «Карточка проекта» в рамках бережливого производства.\n\nТРЕБОВАНИЕ:\n{rule['requirement_expert']}\n\nКРИТЕРИИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nНазвание проекта/потока: "{project_name}"\n\nЗАДАНИЕ:\nПроверь, является ли название проекта осмысленным текстом, описывающим реальный проект или поток.\nНе является осмысленным: placeholder ("Название проекта"), набор символов ("ааааа"), слишком общий текст ("тест").\nЯвляется осмысленным: конкретное описание проекта ("Оптимизация производства приборов учёта").\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{kartochka_proekta_validate_3_flow_name_RULE_INDEX}",\n  "rule_title": "{kartochka_proekta_validate_3_flow_name_RULE_TITLE}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n'
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

def kartochka_proekta_validate_3_flow_name_validate(data: dict, rule: dict, api_key: str) -> dict:
    """Проверка названия проекта/потока без сетевой зависимости."""
    project_name = data.get('header', {}).get('project_name', '')
    if not project_name:
        return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': 'Поле project_name (B4) пустое — нет названия проекта/потока'}
    project_name = project_name.strip()
    if len(project_name) <= 5:
        return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта слишком короткое ({len(project_name)} симв.): '{project_name}'"}
    normalized = ' '.join(project_name.lower().split())
    placeholder_values = {'название проекта', 'название потока', 'проект', 'поток', 'тест', 'test', 'aaaaa', 'aaaa', 'qwerty'}
    letters_only = ''.join((ch for ch in normalized if ch.isalpha()))
    if normalized in placeholder_values:
        return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта выглядит как placeholder: '{project_name}'"}
    if not letters_only:
        return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта не содержит осмысленного текста: '{project_name}'"}
    unique_letters = set(letters_only)
    if len(unique_letters) <= 2:
        return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта похоже на набор повторяющихся символов: '{project_name}'"}
    return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_3_flow_name_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_3_flow_name_RULE_INDEX}: {kartochka_proekta_validate_3_flow_name_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY')
    if not api_key:
        raise ValueError('OPENAI_API_KEY не найден в переменных окружения')
    rule = kartochka_proekta_validate_3_flow_name_load_rule(kartochka_proekta_validate_3_flow_name_RULE_INDEX)
    data = kartochka_proekta_validate_3_flow_name_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_3_flow_name_validate(data, rule, api_key)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_3_flow_name_RULE_INDEX}] {kartochka_proekta_validate_3_flow_name_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_3_flow_name_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_3_flow_name_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_3_flow_name_RULE_TITLE, load_rule=kartochka_proekta_validate_3_flow_name_load_rule, load_data=kartochka_proekta_validate_3_flow_name_load_data, check_semantic=kartochka_proekta_validate_3_flow_name_check_semantic, validate=kartochka_proekta_validate_3_flow_name_validate, main=kartochka_proekta_validate_3_flow_name_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_3_FLOW_NAME

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_4_SIGNEE
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_4_signee.py.
kartochka_proekta_validate_4_signee_RULE_INDEX = '4'

kartochka_proekta_validate_4_signee_RULE_TITLE = 'Должность, ФИО подписанта, дата и подпись'

def kartochka_proekta_validate_4_signee_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_4_signee_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_4_signee_validate(data: dict) -> dict:
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
        return {'rule_index': kartochka_proekta_validate_4_signee_RULE_INDEX, 'rule_title': kartochka_proekta_validate_4_signee_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_4_signee_RULE_INDEX, 'rule_title': kartochka_proekta_validate_4_signee_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_4_signee_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_4_signee_RULE_INDEX}: {kartochka_proekta_validate_4_signee_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_4_signee_load_rule(kartochka_proekta_validate_4_signee_RULE_INDEX)
    data = kartochka_proekta_validate_4_signee_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_4_signee_validate(data)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_4_signee_RULE_INDEX}] {kartochka_proekta_validate_4_signee_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_4_signee_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_4_signee_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_4_signee_RULE_TITLE, load_rule=kartochka_proekta_validate_4_signee_load_rule, load_data=kartochka_proekta_validate_4_signee_load_data, validate=kartochka_proekta_validate_4_signee_validate, main=kartochka_proekta_validate_4_signee_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_4_SIGNEE

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_5_REQUIRED_FIELDS
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_5_required_fields.py.
kartochka_proekta_validate_5_required_fields_RULE_INDEX = '5'

kartochka_proekta_validate_5_required_fields_RULE_TITLE = 'Обязательные поля секции 1'

kartochka_proekta_validate_5_required_fields_REQUIRED_FIELDS = {'clients': 'Клиенты процесса', 'perimeter': 'Периметр проекта', 'owner': 'Владелец процесса', 'boundaries': 'Границы процесса', 'leader': 'Руководитель проекта', 'team': 'Команда проекта'}

kartochka_proekta_validate_5_required_fields_FIO_FIELDS = ['owner', 'leader', 'team']

def kartochka_proekta_validate_5_required_fields_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_5_required_fields_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_5_required_fields__check_fio_format(text: str) -> bool:
    """Проверяет наличие паттерна 'ФИО - должность' (разделитель — тире)."""
    return bool(re.search('.+\\s*[-–—]\\s*.+', text))

def kartochka_proekta_validate_5_required_fields_validate(data: dict) -> dict:
    """Проверка обязательных полей секции 1."""
    section1 = data.get('section1', {})
    errors = []
    empty_fields = []
    for field_key, field_name in kartochka_proekta_validate_5_required_fields_REQUIRED_FIELDS.items():
        val = section1.get(field_key, '')
        if not val or not val.strip():
            empty_fields.append(field_name)
    if empty_fields:
        errors.append(f'Пустые поля: {', '.join(empty_fields)}')
    bad_format = []
    for field_key in kartochka_proekta_validate_5_required_fields_FIO_FIELDS:
        val = section1.get(field_key, '')
        if not val:
            continue
        if field_key == 'team':
            members = [m.strip() for m in val.split(',') if m.strip()]
            for member in members:
                if not kartochka_proekta_validate_5_required_fields__check_fio_format(member):
                    bad_format.append(f"team: '{member[:50]}'")
                    break
        elif not kartochka_proekta_validate_5_required_fields__check_fio_format(val):
            bad_format.append(f"{kartochka_proekta_validate_5_required_fields_REQUIRED_FIELDS[field_key]}: '{val[:50]}'")
    if bad_format:
        errors.append(f'Неверный формат ФИО-должность: {'; '.join(bad_format)}')
    if errors:
        return {'rule_index': kartochka_proekta_validate_5_required_fields_RULE_INDEX, 'rule_title': kartochka_proekta_validate_5_required_fields_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_5_required_fields_RULE_INDEX, 'rule_title': kartochka_proekta_validate_5_required_fields_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_5_required_fields_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_5_required_fields_RULE_INDEX}: {kartochka_proekta_validate_5_required_fields_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_5_required_fields_load_rule(kartochka_proekta_validate_5_required_fields_RULE_INDEX)
    data = kartochka_proekta_validate_5_required_fields_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_5_required_fields_validate(data)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_5_required_fields_RULE_INDEX}] {kartochka_proekta_validate_5_required_fields_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_5_required_fields_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_5_required_fields_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_5_required_fields_RULE_TITLE, REQUIRED_FIELDS=kartochka_proekta_validate_5_required_fields_REQUIRED_FIELDS, FIO_FIELDS=kartochka_proekta_validate_5_required_fields_FIO_FIELDS, load_rule=kartochka_proekta_validate_5_required_fields_load_rule, load_data=kartochka_proekta_validate_5_required_fields_load_data, _check_fio_format=kartochka_proekta_validate_5_required_fields__check_fio_format, validate=kartochka_proekta_validate_5_required_fields_validate, main=kartochka_proekta_validate_5_required_fields_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_5_REQUIRED_FIELDS

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_6_JUSTIFICATION
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_6_justification.py.
kartochka_proekta_validate_6_justification_RULE_INDEX = '6'

kartochka_proekta_validate_6_justification_RULE_TITLE = 'Обоснование и ключевой риск'

def kartochka_proekta_validate_6_justification_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_6_justification_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_6_justification_check_semantic(key_risk: str, justification: str, rule: dict, api_key: str) -> dict:
    """LLM-проверка осмысленности обоснования и ключевого риска (через единый клиент)."""
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    prompt = f'Ты эксперт по проверке документов «Карточка проекта» в рамках бережливого производства.\n\nТРЕБОВАНИЕ:\n{rule['requirement_expert']}\n\nКРИТЕРИИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nКлючевой риск (M11): "{key_risk}"\nОбоснование выбора потока (M13): "{justification}"\n\nЗАДАНИЕ:\nПроверь, содержат ли оба поля осмысленный текст:\n- Ключевой риск — должен описывать конкретный риск проекта (например: "Срыв сроков", "Потеря клиентов").\n  НЕ осмысленный: "-", "нет", "риск", набор символов.\n- Обоснование — должно содержать аргументацию выбора потока (например: "Наличие ожидания в потоке, несвоевременная подготовка").\n  НЕ осмысленный: "-", "обоснование", "тест", набор символов.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{kartochka_proekta_validate_6_justification_RULE_INDEX}",\n  "rule_title": "{kartochka_proekta_validate_6_justification_RULE_TITLE}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n'
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

def kartochka_proekta_validate_6_justification_validate(data: dict, rule: dict, api_key: str) -> dict:
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
        return {'rule_index': kartochka_proekta_validate_6_justification_RULE_INDEX, 'rule_title': kartochka_proekta_validate_6_justification_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return kartochka_proekta_validate_6_justification_check_semantic(key_risk, justification, rule, api_key)

def kartochka_proekta_validate_6_justification_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_6_justification_RULE_INDEX}: {kartochka_proekta_validate_6_justification_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY')
    if not api_key:
        raise ValueError('OPENAI_API_KEY не найден в переменных окружения')
    rule = kartochka_proekta_validate_6_justification_load_rule(kartochka_proekta_validate_6_justification_RULE_INDEX)
    data = kartochka_proekta_validate_6_justification_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_6_justification_validate(data, rule, api_key)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_6_justification_RULE_INDEX}] {kartochka_proekta_validate_6_justification_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_6_justification_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_6_justification_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_6_justification_RULE_TITLE, load_rule=kartochka_proekta_validate_6_justification_load_rule, load_data=kartochka_proekta_validate_6_justification_load_data, check_semantic=kartochka_proekta_validate_6_justification_check_semantic, validate=kartochka_proekta_validate_6_justification_validate, main=kartochka_proekta_validate_6_justification_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_6_JUSTIFICATION

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_7_EVENT_DATES
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_7_event_dates.py.
kartochka_proekta_validate_7_event_dates_RULE_INDEX = '7'

kartochka_proekta_validate_7_event_dates_RULE_TITLE = 'Даты мероприятий'

def kartochka_proekta_validate_7_event_dates_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_7_event_dates_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_7_event_dates__is_stage_event(name: str) -> bool:
    """Проверяет, является ли событие этапным (начинается с '1.', '2.', '3.', '4.')."""
    return bool(re.match('^\\d+\\.', name.strip()))

def kartochka_proekta_validate_7_event_dates_validate(data: dict) -> dict:
    """Проверка дат мероприятий."""
    events = data.get('events', [])
    errors = []
    if not events:
        return {'rule_index': kartochka_proekta_validate_7_event_dates_RULE_INDEX, 'rule_title': kartochka_proekta_validate_7_event_dates_RULE_TITLE, 'status': 'FAIL', 'discrepancy': 'Список мероприятий (events) пуст'}
    for i, event in enumerate(events):
        name = event.get('name', f'Событие #{i + 1}')
        start_date = event.get('start_date')
        end_date = event.get('end_date')
        if not start_date:
            errors.append(f"'{name[:40]}' — нет даты начала")
        if kartochka_proekta_validate_7_event_dates__is_stage_event(name) and (not end_date):
            errors.append(f"'{name[:40]}' — этапное событие без даты окончания")
    if errors:
        return {'rule_index': kartochka_proekta_validate_7_event_dates_RULE_INDEX, 'rule_title': kartochka_proekta_validate_7_event_dates_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_7_event_dates_RULE_INDEX, 'rule_title': kartochka_proekta_validate_7_event_dates_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_7_event_dates_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_7_event_dates_RULE_INDEX}: {kartochka_proekta_validate_7_event_dates_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_7_event_dates_load_rule(kartochka_proekta_validate_7_event_dates_RULE_INDEX)
    data = kartochka_proekta_validate_7_event_dates_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_7_event_dates_validate(data)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_7_event_dates_RULE_INDEX}] {kartochka_proekta_validate_7_event_dates_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_7_event_dates_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_7_event_dates_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_7_event_dates_RULE_TITLE, load_rule=kartochka_proekta_validate_7_event_dates_load_rule, load_data=kartochka_proekta_validate_7_event_dates_load_data, _is_stage_event=kartochka_proekta_validate_7_event_dates__is_stage_event, validate=kartochka_proekta_validate_7_event_dates_validate, main=kartochka_proekta_validate_7_event_dates_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_7_EVENT_DATES

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_8_INDICATORS
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_8_indicators.py.
kartochka_proekta_validate_8_indicators_RULE_INDEX = '8'

kartochka_proekta_validate_8_indicators_RULE_TITLE = 'Показатели и даты'

def kartochka_proekta_validate_8_indicators_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_8_indicators_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_8_indicators_validate(data: dict) -> dict:
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
        return {'rule_index': kartochka_proekta_validate_8_indicators_RULE_INDEX, 'rule_title': kartochka_proekta_validate_8_indicators_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_8_indicators_RULE_INDEX, 'rule_title': kartochka_proekta_validate_8_indicators_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_8_indicators_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_8_indicators_RULE_INDEX}: {kartochka_proekta_validate_8_indicators_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_8_indicators_load_rule(kartochka_proekta_validate_8_indicators_RULE_INDEX)
    data = kartochka_proekta_validate_8_indicators_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_8_indicators_validate(data)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_8_indicators_RULE_INDEX}] {kartochka_proekta_validate_8_indicators_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_8_indicators_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_8_indicators_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_8_indicators_RULE_TITLE, load_rule=kartochka_proekta_validate_8_indicators_load_rule, load_data=kartochka_proekta_validate_8_indicators_load_data, validate=kartochka_proekta_validate_8_indicators_validate, main=kartochka_proekta_validate_8_indicators_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_8_INDICATORS

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_9_UNITS_KARTOCHKA
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_9_units_kartochka.py.
kartochka_proekta_validate_9_units_kartochka_RULE_INDEX = '9'

kartochka_proekta_validate_9_units_kartochka_RULE_TITLE = 'Единицы измерения (карточка)'

kartochka_proekta_validate_9_units_kartochka_CATEGORY_KEYWORDS = {'Время протекания процесса': ['время', 'протекан'], 'Выработка': ['выработк'], 'Незавершенное производство': ['запас', 'нзп', 'незавершен']}

kartochka_proekta_validate_9_units_kartochka_DEFAULT_CATEGORY = 'Дополнительный показатель'

def kartochka_proekta_validate_9_units_kartochka_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_9_units_kartochka_load_data(parser_outputs_dir: Path) -> tuple:
    """Загрузка kartochka_main.json и dropdown_units.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        kartochka = json.load(f)
    with open(parser_outputs_dir / 'dropdown_units.json', 'r', encoding='utf-8') as f:
        dropdown = json.load(f)
    return (kartochka, dropdown)

def kartochka_proekta_validate_9_units_kartochka__detect_category(indicator_name: str) -> str:
    """Определяет категорию показателя по ключевым словам в названии."""
    name_lower = indicator_name.lower()
    for category, keywords in kartochka_proekta_validate_9_units_kartochka_CATEGORY_KEYWORDS.items():
        if any((kw in name_lower for kw in keywords)):
            return category
    return kartochka_proekta_validate_9_units_kartochka_DEFAULT_CATEGORY

def kartochka_proekta_validate_9_units_kartochka_validate(kartochka: dict, dropdown: dict) -> dict:
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
        return {'rule_index': kartochka_proekta_validate_9_units_kartochka_RULE_INDEX, 'rule_title': kartochka_proekta_validate_9_units_kartochka_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_9_units_kartochka_RULE_INDEX, 'rule_title': kartochka_proekta_validate_9_units_kartochka_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_9_units_kartochka_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_9_units_kartochka_RULE_INDEX}: {kartochka_proekta_validate_9_units_kartochka_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_9_units_kartochka_load_rule(kartochka_proekta_validate_9_units_kartochka_RULE_INDEX)
    kartochka, dropdown = kartochka_proekta_validate_9_units_kartochka_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_9_units_kartochka_validate(kartochka, dropdown)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_9_units_kartochka_RULE_INDEX}] {kartochka_proekta_validate_9_units_kartochka_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_9_units_kartochka_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_9_units_kartochka_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_9_units_kartochka_RULE_TITLE, CATEGORY_KEYWORDS=kartochka_proekta_validate_9_units_kartochka_CATEGORY_KEYWORDS, DEFAULT_CATEGORY=kartochka_proekta_validate_9_units_kartochka_DEFAULT_CATEGORY, load_rule=kartochka_proekta_validate_9_units_kartochka_load_rule, load_data=kartochka_proekta_validate_9_units_kartochka_load_data, _detect_category=kartochka_proekta_validate_9_units_kartochka__detect_category, validate=kartochka_proekta_validate_9_units_kartochka_validate, main=kartochka_proekta_validate_9_units_kartochka_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_9_UNITS_KARTOCHKA

# START_VALIDATOR_DISPATCH
# PURPOSE: Маппинг rule_index → SimpleNamespace валидатора (1–11). Используется раннером карточки проекта.
KARTOCHKA_VALIDATOR_MODULE_DISPATCH = {
    "1": kartochka_proekta_validate_1_filename_module,
    "2": kartochka_proekta_validate_2_org_name_module,
    "3": kartochka_proekta_validate_3_flow_name_module,
    "4": kartochka_proekta_validate_4_signee_module,
    "5": kartochka_proekta_validate_5_required_fields_module,
    "6": kartochka_proekta_validate_6_justification_module,
    "7": kartochka_proekta_validate_7_event_dates_module,
    "8": kartochka_proekta_validate_8_indicators_module,
    "9": kartochka_proekta_validate_9_units_kartochka_module,
    "10": kartochka_proekta_validate_10_units_metodika_module,
    "11": kartochka_proekta_validate_11_calc_method_module,
}
# END_VALIDATOR_DISPATCH


# START_VALIDATOR_RUNNER
def run_kartochka_validator_module(module_ns: SimpleNamespace, parser_outputs_dir: Path, output_file: Path) -> Dict[str, Any]:
    """
    Назначение:
        Универсальный раннер одного валидатора карточки проекта. Сигнатура
        module_ns.validate(...) варьируется (разные правила принимают разный
        набор данных + опционально rule/api_key). Раннер по inspect'у подбирает
        аргументы.

    Вход:
        module_ns: SimpleNamespace валидатора (из KARTOCHKA_VALIDATOR_MODULE_DISPATCH).
        parser_outputs_dir: Директория с JSON-ами парсера kartochka_proekta.
        output_file: Путь для записи JSON с результатом.

    Выход:
        dict {rule_index, rule_title, status=PASS/FAIL, discrepancy}. Также
        пишет output_file.
    """
    import inspect

    rule = module_ns.load_rule(module_ns.RULE_INDEX)
    loaded_data = module_ns.load_data(parser_outputs_dir)
    pending_data = list(loaded_data) if isinstance(loaded_data, tuple) else [loaded_data]
    call_args: List[Any] = []
    api_key = os.environ.get("OPENAI_API_KEY", "dummy")

    for param in inspect.signature(module_ns.validate).parameters.values():
        if param.name == "rule":
            call_args.append(rule)
        elif param.name == "api_key":
            call_args.append(api_key)
        else:
            if not pending_data:
                raise ValueError(f"Недостаточно данных для валидатора {module_ns.RULE_INDEX}")
            call_args.append(pending_data.pop(0))

    result = module_ns.validate(*call_args)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return result
# END_VALIDATOR_RUNNER
