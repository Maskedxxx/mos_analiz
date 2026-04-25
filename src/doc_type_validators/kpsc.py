# START_MODULE_CONTRACT
# PURPOSE: Валидаторы документа КПСЦ (Карта Потока Создания Ценности). 25 проверок правил 1.1–8.3 по распарсенным JSON-ам из `src/doc_type_parsers/kpsc.py`.
# INPUTS: parser_outputs_dir с JSON-ами (kpsc_header_v2.json, kpsc_table1_v2.json, ...); rule из validation_rules.json (путь через os.environ['VALIDATION_RULES_PATH']); LLM-сервис через src.llm.client.call_llm.
# OUTPUTS: dict {rule_index, rule_title, target_document, status=PASS/FAIL, discrepancy}. Публичные символы: `KPSC_VALIDATOR_MODULE_DISPATCH`, `run_kpsc_validator_module`.
# KEYWORDS: validators, kpsc, llm, json-verdict.
# LINKS: src/llm/client.py (call_llm), config/llm.py (LLM_CONFIG), src/doc_type_parsers/kpsc.py (источник данных), doc_configs/kpsc/validation_rules.json.
# RATIONALE:
#   Каждый валидатор — микро-пайплайн: load_rule → load_data → extract_relevant_data →
#   build_prompt → call_llm → save_result. Шаги load_rule/call_llm/save_result
#   одинаковы у всех 25 правил и вынесены в общие helpers. Уникальное у каждого —
#   только load_data (какой JSON читать), extract_relevant_data (какие поля
#   достать) и build_prompt (текст промпта с фактическими данными).
#   Каждое правило регистрируется в `KPSC_VALIDATOR_MODULE_DISPATCH` как
#   `SimpleNamespace`, совместимый с `run_kpsc_validator_module`.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict, List, Tuple

import pandas as pd
from openpyxl.utils import get_column_letter

from config.llm import LLM_CONFIG
from src.doc_type_parsers.kpsc import (
    parse_kpsc_header,
    parse_kpsc_table1,
    parse_legend,
    parse_loss_digitization,
    parse_pa1_chart,
    parse_pa1_table,
    parse_pokazateli,
    parse_spaghetti_problems,
    parse_spaghetti_sheet,
)
from src.llm.client import call_llm, resolve_runtime_llm_model
# END_IMPORTS


# START_PATHS
# PURPOSE: parents[2] = repo root.
_REPO_ROOT = Path(__file__).resolve().parents[2]
_DOC_CONFIGS_DIR = _REPO_ROOT / "doc_configs"
_LOGS_RESULT_DIR = _REPO_ROOT / "logs_result"
# END_PATHS


# START_PARSER_DISPATCH
# PURPOSE: Маппинг "имя парсера → callable". Используется `_run_kpsc_parsers`.
KPSC_PARSER_MODULE_DISPATCH = {
    "parse_kpsc_header": parse_kpsc_header,
    "parse_kpsc_table1": parse_kpsc_table1,
    "parse_legend": parse_legend,
    "parse_loss_digitization": parse_loss_digitization,
    "parse_pa1_chart": parse_pa1_chart,
    "parse_pa1_table": parse_pa1_table,
    "parse_pokazateli": parse_pokazateli,
    "parse_spaghetti_sheet": parse_spaghetti_sheet,
    "parse_spaghetti_problems": parse_spaghetti_problems,
}
# END_PARSER_DISPATCH


# START_CONSTANTS
# PURPOSE: Общие константы КПСЦ-валидаторов. Target-документ один для всех 25 правил (КПСЦ — один Excel-файл); system-prompt тоже один.
TARGET_DOC = "КПСЦ и Спагетти ТС_Предприятие.xlsx"
SYSTEM_PROMPT = "Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON."
# END_CONSTANTS


# START_COMMON_HELPERS
# PURPOSE: load_rule/call_llm/save_result — одинаковы у всех 25 валидаторов, лежат здесь в одной копии. В SimpleNamespace каждого правила эти же функции переиспользуются.
def _read_json_file(json_file: Path) -> Any:
    """
    Назначение:
        Читает JSON-файл в UTF-8 для валидаторов КПСЦ.

    Вход:
        json_file: Путь к JSON-файлу.

    Выход:
        Распарсенное содержимое JSON-файла.
    """
    with open(json_file, "r", encoding="utf-8") as f:
        return json.load(f)


def _load_required_json(parser_outputs_dir: Path, filename: str) -> Any:
    """
    Назначение:
        Читает обязательный JSON-файл из директории результатов парсера.

    Вход:
        parser_outputs_dir: Директория с JSON-ами парсера КПСЦ.
        filename: Имя JSON-файла.

    Выход:
        Распарсенное содержимое JSON-файла.
    """
    return _read_json_file(parser_outputs_dir / filename)


def _load_optional_json(parser_outputs_dir: Path, filename: str, *, exists_key: str = "file_exists") -> dict:
    """
    Назначение:
        Читает опциональный JSON-файл и возвращает стандартный статус наличия.

    Вход:
        parser_outputs_dir: Директория с JSON-ами парсера КПСЦ.
        filename: Имя JSON-файла.
        exists_key: Имя ключа со статусом наличия файла.

    Выход:
        dict вида `{exists_key: bool, data: ...}`.
    """
    json_file = parser_outputs_dir / filename
    if not json_file.exists():
        return {exists_key: False, "data": None}
    return {exists_key: True, "data": _read_json_file(json_file)}


def _load_required_jsons(parser_outputs_dir: Path, filenames_by_key: dict) -> dict:
    """
    Назначение:
        Читает несколько обязательных JSON-файлов в словарь по заданным ключам.

    Вход:
        parser_outputs_dir: Директория с JSON-ами парсера КПСЦ.
        filenames_by_key: Маппинг `{ключ_данных: имя_json_файла}`.

    Выход:
        dict `{ключ_данных: распарсенный_json}`.
    """
    return {key: _load_required_json(parser_outputs_dir, filename) for key, filename in filenames_by_key.items()}


def _load_existing_jsons(parser_outputs_dir: Path, filenames_by_key: dict) -> tuple[dict, dict]:
    """
    Назначение:
        Читает набор JSON-файлов, если они существуют, и отдельно возвращает
        статус наличия каждого файла.

    Вход:
        parser_outputs_dir: Директория с JSON-ами парсера КПСЦ.
        filenames_by_key: Маппинг `{ключ_данных: имя_json_файла}`.

    Выход:
        tuple `(files_exist, data_content)`.
    """
    files_exist = {}
    data_content = {}
    for key, filename in filenames_by_key.items():
        json_file = parser_outputs_dir / filename
        files_exist[filename] = json_file.exists()
        if json_file.exists():
            data_content[key] = _read_json_file(json_file)
    return files_exist, data_content


def _load_rule(rule_index: str) -> dict:
    """
    Назначение:
        Читает правило проверки из `validation_rules.json` по `rule_index`.

    Вход:
        rule_index: Индекс правила, например `'1.1'`.

    Выход:
        dict с полями правила (requirement_expert, validation_criteria, ...).

    Логика:
        1. Путь к rules.json берётся из `os.environ['VALIDATION_RULES_PATH']`
           (ставит KPSC-раннер перед запуском валидаторов).
        2. Fallback на `Path(__file__).parent.parent / 'validation_rules.json'` —
           legacy-путь для standalone-запуска, сейчас не используется рантаймом.
    """
    rules_file = Path(os.environ.get("VALIDATION_RULES_PATH", str(Path(__file__).parent.parent / "validation_rules.json")))
    data = _read_json_file(rules_file)
    for rule in data["rules"]:
        if rule["rule_index"] == rule_index:
            return rule
    raise ValueError(f"Правило {rule_index} не найдено в validation_rules.json")


def _call_validator_llm(prompt: str, api_key: str = "") -> dict:
    """
    Назначение:
        Единый вызов LLM для всех 25 валидаторов. Использует `call_llm` из
        `src.llm.client`, параметры LLM берутся из `LLM_CONFIG`.

    Вход:
        prompt: Текст пользовательского промпта (формируется в `build_prompt` правила).
        api_key: Игнорируется (Spark-vLLM не проверяет ключ, `src.llm.client` сам
            подставляет `api_key='none'`). Оставлен в сигнатуре для совместимости
            с legacy-раннером, который передаёт api_key из env.

    Выход:
        dict, распарсенный из JSON-ответа LLM ({rule_index, status, discrepancy, ...}).
    """
    del api_key
    result_text = call_llm(
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={"type": "json_object"},
    )
    return json.loads(result_text)


def _save_result(result: dict, output_file: Path) -> None:
    """
    Назначение:
        Сохраняет результат валидатора в JSON-файл.

    Вход:
        result: dict с результатом проверки.
        output_file: Путь к выходному файлу.
    """
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"✓ Результат сохранен: {output_file}")


def _build_validator_prompt(rule: dict, fact_block: str, task_note: str = "") -> str:
    """
    Назначение:
        Формирует общий текст LLM-подсказки для КПСЦ-валидатора.

    Вход:
        rule: Правило из validation_rules.json.
        fact_block: Уникальный для правила блок фактических данных.
        task_note: Дополнительная инструкция внутри блока `ЗАДАНИЕ`.

    Выход:
        Готовая текстовая подсказка для LLM в том же формате, который
        использовали отдельные build_prompt-функции.
    """
    task_text = "Проверь соответствие фактических данных требованию эксперта и критериям проверки."
    if task_note:
        task_text = f"{task_text}\n{task_note}"
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\n{fact_block}\n\nЗАДАНИЕ:\n{task_text}\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt


def _make_validator_module(
    *,
    rule_index: str,
    rule_title: str,
    load_data,
    extract_relevant_data,
    build_prompt,
) -> SimpleNamespace:
    """
    Назначение:
        Собирает SimpleNamespace-обёртку валидатора КПСЦ с единым набором
        общих функций и уникальными шагами конкретного правила.

    Вход:
        rule_index: Индекс правила из validation_rules.json.
        rule_title: Название правила для совместимости с legacy-модулями.
        load_data: Функция загрузки JSON-данных парсера.
        extract_relevant_data: Функция извлечения релевантных полей.
        build_prompt: Функция формирования текстовой подсказки для LLM.

    Выход:
        SimpleNamespace с тем же интерфейсом, который ожидает
        run_kpsc_validator_module.
    """
    return SimpleNamespace(
        RULE_INDEX=rule_index,
        RULE_TITLE=rule_title,
        TARGET_DOC=TARGET_DOC,
        load_rule=_load_rule,
        load_data=load_data,
        extract_relevant_data=extract_relevant_data,
        build_prompt=build_prompt,
        call_llm=_call_validator_llm,
        save_result=_save_result,
    )
# END_COMMON_HELPERS


# ==============================================================================
# RULE 1.1 — Наличие текста 'КПСЦ' в заголовке
# ==============================================================================
def _rule_1_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_required_json(parser_outputs_dir, 'kpsc_header_v2.json')

def _rule_1_1_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'title': data['fields'].get('title', '')}

def _rule_1_1_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Заголовок документа: "{extracted_data['title']}"'
    return _build_validator_prompt(rule, fact_block)

RULE_1_1 = _make_validator_module(
    rule_index='1.1',
    rule_title="Наличие текста 'КПСЦ' в заголовке",
    load_data=_rule_1_1_load_data,
    extract_relevant_data=_rule_1_1_extract_relevant_data,
    build_prompt=_rule_1_1_build_prompt,
)


# ==============================================================================
# RULE 1.2 — Наличие названия предприятия в формате ООО "наименование"
# ==============================================================================
def _rule_1_2_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_required_json(parser_outputs_dir, 'kpsc_header_v2.json')

def _rule_1_2_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'title': data['fields'].get('title', ''), 'organization': data['fields'].get('organization', ''), 'workbook': data.get('meta', {}).get('workbook', '')}

def _rule_1_2_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Заголовок документа: "{extracted_data['title']}"\nНазвание организации: "{extracted_data.get('organization', '')}"'
    return _build_validator_prompt(rule, fact_block)

RULE_1_2 = _make_validator_module(
    rule_index='1.2',
    rule_title='Наличие названия предприятия в формате ООО "наименование"',
    load_data=_rule_1_2_load_data,
    extract_relevant_data=_rule_1_2_extract_relevant_data,
    build_prompt=_rule_1_2_build_prompt,
)


# ==============================================================================
# RULE 1.3 — Наличие названия потока в формате "имя потока"
# ==============================================================================
def _rule_1_3_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_required_json(parser_outputs_dir, 'kpsc_header_v2.json')

def _rule_1_3_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'flow_name': data['fields'].get('flow_name', '')}

def _rule_1_3_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Название потока (flow_name): "{extracted_data['flow_name']}"'
    return _build_validator_prompt(rule, fact_block)

RULE_1_3 = _make_validator_module(
    rule_index='1.3',
    rule_title='Наличие названия потока в формате "имя потока"',
    load_data=_rule_1_3_load_data,
    extract_relevant_data=_rule_1_3_extract_relevant_data,
    build_prompt=_rule_1_3_build_prompt,
)


# ==============================================================================
# RULE 1.4 — Заполнение поля "Ответственный за поток" (ФИО)
# ==============================================================================
def _rule_1_4_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_required_json(parser_outputs_dir, 'kpsc_header_v2.json')

def _rule_1_4_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    responsible = data['fields'].get('responsible', '') or ''
    words = responsible.strip().split()
    return {'responsible': responsible, 'word_count': len(words), 'char_length': len(responsible.strip())}

def _rule_1_4_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Ответственный за поток (responsible): "{extracted_data['responsible']}"'
    return _build_validator_prompt(rule, fact_block)

RULE_1_4 = _make_validator_module(
    rule_index='1.4',
    rule_title='Заполнение поля "Ответственный за поток" (ФИО)',
    load_data=_rule_1_4_load_data,
    extract_relevant_data=_rule_1_4_extract_relevant_data,
    build_prompt=_rule_1_4_build_prompt,
)


# ==============================================================================
# RULE 1.5 — Заполнение даты разработки
# ==============================================================================
def _rule_1_5_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_required_json(parser_outputs_dir, 'kpsc_header_v2.json')

def _rule_1_5_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'date_developed': data['fields'].get('date_developed', '')}

def _rule_1_5_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Дата разработки (date_developed): "{extracted_data['date_developed']}"'
    return _build_validator_prompt(rule, fact_block)

RULE_1_5 = _make_validator_module(
    rule_index='1.5',
    rule_title='Заполнение даты разработки',
    load_data=_rule_1_5_load_data,
    extract_relevant_data=_rule_1_5_extract_relevant_data,
    build_prompt=_rule_1_5_build_prompt,
)


# ==============================================================================
# RULE 1.6 — Заполнение даты реализации
# ==============================================================================
def _rule_1_6_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_required_json(parser_outputs_dir, 'kpsc_header_v2.json')

def _rule_1_6_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'date_implementation': data['fields'].get('date_implementation', '')}

def _rule_1_6_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Дата реализации (date_implementation): "{extracted_data['date_implementation']}"'
    return _build_validator_prompt(rule, fact_block)

RULE_1_6 = _make_validator_module(
    rule_index='1.6',
    rule_title='Заполнение даты реализации',
    load_data=_rule_1_6_load_data,
    extract_relevant_data=_rule_1_6_extract_relevant_data,
    build_prompt=_rule_1_6_build_prompt,
)


# ==============================================================================
# RULE 1.7 — Заполнение поля "Кто составил" (ФИО)
# ==============================================================================
def _rule_1_7_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_required_json(parser_outputs_dir, 'kpsc_header_v2.json')

def _rule_1_7_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'compiled_by': data['fields'].get('compiled_by', '')}

def _rule_1_7_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Кто составил (compiled_by): "{extracted_data['compiled_by']}"'
    return _build_validator_prompt(rule, fact_block)

RULE_1_7 = _make_validator_module(
    rule_index='1.7',
    rule_title='Заполнение поля "Кто составил" (ФИО)',
    load_data=_rule_1_7_load_data,
    extract_relevant_data=_rule_1_7_extract_relevant_data,
    build_prompt=_rule_1_7_build_prompt,
)


# ==============================================================================
# RULE 2.1 — Количество проблем в таблице "Оцифровка потерь"
# ==============================================================================
def _rule_2_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'ocifrovka_poteri_v2.json')

def _rule_2_1_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'problems_count': 0, 'problem_numbers': []}
    rows = data['data'].get('rows', [])
    problem_numbers = []
    for i, row in enumerate(rows):
        if i == 0:
            continue
        cells = row.get('cells', [])
        for cell in cells:
            if cell.get('col') in [1, 2]:
                problem_num_value = cell.get('value')
                if problem_num_value and str(problem_num_value).strip():
                    try:
                        problem_num = int(str(problem_num_value).strip())
                        problem_numbers.append(problem_num)
                    except:
                        pass
                    break
    return {'file_exists': True, 'problems_count': len(problem_numbers), 'problem_numbers': problem_numbers, 'min_number': min(problem_numbers) if problem_numbers else 0, 'max_number': max(problem_numbers) if problem_numbers else 0}

def _rule_2_1_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Файл существует: {extracted_data['file_exists']}\nКоличество проблем: {extracted_data['problems_count']}\nНомера проблем: {extracted_data['problem_numbers']}\nДиапазон: с {extracted_data.get('min_number', 0)} по {extracted_data.get('max_number', 0)}'
    return _build_validator_prompt(
        rule,
        fact_block,
        "Это информационная проверка - нужно подтвердить, что найдены проблемы и извлечены номера.",
    )

RULE_2_1 = _make_validator_module(
    rule_index='2.1',
    rule_title='Количество проблем в таблице "Оцифровка потерь"',
    load_data=_rule_2_1_load_data,
    extract_relevant_data=_rule_2_1_extract_relevant_data,
    build_prompt=_rule_2_1_build_prompt,
)


# ==============================================================================
# RULE 2.2 — Последовательность номеров проблем (пропуски в нумерации)
# ==============================================================================
def _rule_2_2_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'ocifrovka_poteri_v2.json')

def _rule_2_2_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'problem_numbers': [], 'missing_numbers': []}
    rows = data['data'].get('rows', [])
    problem_numbers = []
    for i, row in enumerate(rows):
        if i == 0:
            continue
        cells = row.get('cells', [])
        for cell in cells:
            if cell.get('col') == 2:
                problem_num_value = cell.get('value')
                if problem_num_value and str(problem_num_value).strip():
                    try:
                        problem_num = int(str(problem_num_value).strip())
                        problem_numbers.append(problem_num)
                    except:
                        pass
                break
    if problem_numbers:
        min_num = min(problem_numbers)
        max_num = max(problem_numbers)
        expected_numbers = set(range(min_num, max_num + 1))
        actual_numbers = set(problem_numbers)
        missing_numbers = sorted(expected_numbers - actual_numbers)
    else:
        missing_numbers = []
    return {'file_exists': True, 'problem_numbers': sorted(problem_numbers), 'missing_numbers': missing_numbers, 'has_missing': len(missing_numbers) > 0}

def _rule_2_2_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Найденные номера проблем: {extracted_data['problem_numbers']}\nПропущенные номера: {extracted_data['missing_numbers']}\nЕсть пропуски: {extracted_data['has_missing']}'
    return _build_validator_prompt(rule, fact_block)

RULE_2_2 = _make_validator_module(
    rule_index='2.2',
    rule_title='Последовательность номеров проблем (пропуски в нумерации)',
    load_data=_rule_2_2_load_data,
    extract_relevant_data=_rule_2_2_extract_relevant_data,
    build_prompt=_rule_2_2_build_prompt,
)


# ==============================================================================
# RULE 2.3 — Наличие описания для каждого номера проблемы
# ==============================================================================
def _rule_2_3_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'ocifrovka_poteri_v2.json')

def _rule_2_3_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'problems_without_description': []}
    rows = data['data'].get('rows', [])
    problems_without_description = []
    for i, row in enumerate(rows):
        if i == 0:
            continue
        cells = row.get('cells', [])
        problem_num_value = None
        description = None
        for cell in cells:
            if cell.get('col') == 2:
                problem_num_value = cell.get('value', '')
            if cell.get('col') == 4:
                description = cell.get('value', '')
        if problem_num_value and str(problem_num_value).strip():
            if not description or not str(description).strip():
                try:
                    problem_num = int(str(problem_num_value).strip())
                    problems_without_description.append(problem_num)
                except:
                    pass
    return {'file_exists': True, 'problems_without_description': problems_without_description, 'has_missing_descriptions': len(problems_without_description) > 0, 'missing_count': len(problems_without_description)}

def _rule_2_3_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Проблемы без описания: {extracted_data['problems_without_description']}\nЕсть проблемы без описания: {extracted_data['has_missing_descriptions']}\nКоличество проблем без описания: {extracted_data['missing_count']}'
    return _build_validator_prompt(rule, fact_block)

RULE_2_3 = _make_validator_module(
    rule_index='2.3',
    rule_title='Наличие описания для каждого номера проблемы',
    load_data=_rule_2_3_load_data,
    extract_relevant_data=_rule_2_3_extract_relevant_data,
    build_prompt=_rule_2_3_build_prompt,
)


# ==============================================================================
# RULE 4.1 — Все ячейки строки ВПП заполнены
# ==============================================================================
def _rule_4_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'kpsc_table1_v2.json')

def _rule_4_1_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'empty_cells': []}
    rows = data['data'].get('rows', [])
    vpp_row_idx = None
    header_row_idx = None
    for i, row in enumerate(rows):
        cells = row.get('cells', [])
        for cell in cells:
            if cell.get('col') in [2, 3]:
                cell_value = str(cell.get('value', '')).upper()
                if 'ВПП' in cell_value:
                    vpp_row_idx = i
                    for j in range(max(0, i - 10), i):
                        row_cells = rows[j].get('cells', [])
                        for rc in row_cells:
                            if rc.get('col') in [2, 3]:
                                if str(rc.get('value', '')).lower().strip() == 'показатель':
                                    header_row_idx = j
                                    break
                        if header_row_idx is not None:
                            break
                    break
        if vpp_row_idx is not None:
            break
    if vpp_row_idx is None:
        return {'file_exists': True, 'empty_cells': [], 'error': 'Не найдена строка ВПП'}
    vpp_cells = rows[vpp_row_idx].get('cells', [])
    vpp_dict = {cell.get('col'): cell.get('value') for cell in vpp_cells}
    header_cells = rows[header_row_idx].get('cells', []) if header_row_idx is not None else []
    header_dict = {cell.get('col'): cell.get('value', '') for cell in header_cells}
    empty_cells = []
    all_cols = set(vpp_dict.keys()) | set(header_dict.keys())
    for col_num in sorted(all_cols):
        if col_num < 5:
            continue
        raw_header = header_dict.get(col_num)
        header_val = str(raw_header).strip() if raw_header is not None else ''
        if not header_val:
            continue
        if 'итого' in header_val.lower():
            continue
        vpp_value = vpp_dict.get(col_num)
        if vpp_value is None or str(vpp_value).strip() == '':
            empty_cells.append({'column': col_num, 'operation_name': header_val if header_val else f'Колонка {col_num}'})
    return {'file_exists': True, 'empty_cells': empty_cells, 'has_empty': len(empty_cells) > 0, 'empty_count': len(empty_cells)}

def _rule_4_1_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Пустые ячейки в строке ВПП:\n{json.dumps(extracted_data.get('empty_cells', []), ensure_ascii=False, indent=2)}\n\nЕсть пустые ячейки: {extracted_data.get('has_empty', False)}\nКоличество пустых: {extracted_data.get('empty_count', 0)}'
    return _build_validator_prompt(rule, fact_block)

RULE_4_1 = _make_validator_module(
    rule_index='4.1',
    rule_title='Все ячейки строки ВПП заполнены',
    load_data=_rule_4_1_load_data,
    extract_relevant_data=_rule_4_1_extract_relevant_data,
    build_prompt=_rule_4_1_build_prompt,
)


# ==============================================================================
# RULE 4.2 — Сумма значений строки ВПП совпадает с итоговым значением
# ==============================================================================
def _rule_4_2_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'kpsc_table1_v2.json')

def _rule_4_2_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'vpp_values': [], 'calculated_sum': 0, 'itogo_value': 0}
    rows = data['data'].get('rows', [])
    vpp_row_idx = None
    header_row_idx = None
    for i, row in enumerate(rows):
        cells = row.get('cells', [])
        for cell in cells:
            if cell.get('col') in [2, 3]:
                cell_value = str(cell.get('value', '')).upper()
                if 'ВПП' in cell_value:
                    vpp_row_idx = i
                    for j in range(max(0, i - 20), i):
                        row_cells = rows[j].get('cells', [])
                        for rc in row_cells:
                            if rc.get('col') in [2, 3]:
                                if str(rc.get('value', '')).lower().strip() == 'показатель':
                                    header_row_idx = j
                                    break
                        if header_row_idx is not None:
                            break
                    break
        if vpp_row_idx is not None:
            break
    if vpp_row_idx is None:
        return {'file_exists': True, 'error': 'Не найдена строка ВПП'}
    vpp_cells = rows[vpp_row_idx].get('cells', [])
    vpp_dict = {}
    vpp_cells_full = {}
    for cell in vpp_cells:
        col = cell.get('col')
        vpp_dict[col] = cell.get('value')
        vpp_cells_full[col] = cell
    header_cells = rows[header_row_idx].get('cells', []) if header_row_idx is not None else []
    header_dict = {cell.get('col'): str(cell.get('value', '')).strip() for cell in header_cells}
    vpp_values = []
    itogo_value = None
    all_cols = set(vpp_dict.keys()) | set(header_dict.keys())
    for col_num in sorted(all_cols):
        if col_num < 5:
            continue
        header_val = header_dict.get(col_num, '')
        if col_num in vpp_cells_full:
            cell_info = vpp_cells_full[col_num]
            is_merge_anchor = cell_info.get('merge_anchor', False)
            merge_range = cell_info.get('merge_range')
            if merge_range is not None and (not is_merge_anchor):
                continue
        cell_value = vpp_dict.get(col_num)
        if 'итого' in header_val.lower():
            try:
                itogo_value = float(str(cell_value).replace(',', '.'))
            except:
                itogo_value = None
            break
        if cell_value is not None and str(cell_value).strip():
            try:
                num_value = float(str(cell_value).replace(',', '.'))
                vpp_values.append({'column': col_num, 'operation': header_val if header_val else f'Колонка {col_num}', 'value': num_value})
            except:
                pass
    calculated_sum = sum((v['value'] for v in vpp_values))
    difference = None
    if itogo_value is not None:
        difference = abs(calculated_sum - itogo_value)
    return {'file_exists': True, 'vpp_values': vpp_values, 'calculated_sum': calculated_sum, 'itogo_value': itogo_value, 'difference': difference, 'matches': difference <= 0.01 if difference is not None else False}

def _rule_4_2_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Значения ВПП по операциям:\n{json.dumps(extracted_data.get('vpp_values', []), ensure_ascii=False, indent=2)}\n\nВычисленная сумма: {extracted_data.get('calculated_sum', 0)}\nЗначение ИТОГО: {extracted_data.get('itogo_value', 0)}\nРасхождение: {extracted_data.get('difference', 0)}\nСуммы совпадают (±0.01): {extracted_data.get('matches', False)}'
    return _build_validator_prompt(rule, fact_block)

RULE_4_2 = _make_validator_module(
    rule_index='4.2',
    rule_title='Сумма значений строки ВПП совпадает с итоговым значением',
    load_data=_rule_4_2_load_data,
    extract_relevant_data=_rule_4_2_extract_relevant_data,
    build_prompt=_rule_4_2_build_prompt,
)


# ==============================================================================
# RULE 5.1 — Заполнение всех ячеек в столбце "Единицы измерения"
# ==============================================================================
def _rule_5_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'kpsc_table1_v2.json')

def _rule_5_1_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'empty_units': []}
    rows = data['data'].get('rows', [])
    empty_units = []
    units_col = 4
    for i, row in enumerate(rows):
        if i == 0:
            continue
        cells = row.get('cells', [])
        row_num = row.get('row')
        cells_dict = {cell.get('col'): cell.get('value') for cell in cells}
        unit_value = cells_dict.get(units_col)
        if unit_value is None or str(unit_value).strip() == '':
            indicator_name = cells_dict.get(2, '')
            if not indicator_name or str(indicator_name).strip() == '':
                indicator_name = cells_dict.get(3, '')
            has_numeric_data = False
            for col_idx, val in cells_dict.items():
                if col_idx >= 5 and val is not None:
                    try:
                        float(str(val).replace(',', '.').strip())
                        has_numeric_data = True
                        break
                    except (ValueError, TypeError):
                        pass
            if not has_numeric_data:
                continue
            empty_units.append({'row': row_num, 'indicator': str(indicator_name).strip()})
    return {'file_exists': True, 'empty_units': empty_units, 'has_empty_units': len(empty_units) > 0, 'empty_count': len(empty_units)}

def _rule_5_1_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Строки с пустыми единицами измерения: {extracted_data['empty_units']}\nЕсть пустые ячейки: {extracted_data['has_empty_units']}\nКоличество пустых: {extracted_data['empty_count']}'
    return _build_validator_prompt(rule, fact_block)

RULE_5_1 = _make_validator_module(
    rule_index='5.1',
    rule_title='Заполнение всех ячеек в столбце "Единицы измерения"',
    load_data=_rule_5_1_load_data,
    extract_relevant_data=_rule_5_1_extract_relevant_data,
    build_prompt=_rule_5_1_build_prompt,
)


# ==============================================================================
# RULE 6.1 — Наличие строки "Перемещения" в блоке "Расчет ВПП"
# ==============================================================================
def _rule_6_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'kpsc_table1_v2.json')

def _rule_6_1_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_transport_row': False}
    rows = data['data'].get('rows', [])
    has_transport_row = False
    transport_row_num = None
    for row in rows:
        cells = row.get('cells', [])
        row_num = row.get('row')
        for cell in cells:
            if cell.get('col') in [1, 2, 3]:
                cell_value = str(cell.get('value', '')).lower()
                if 'перемещени' in cell_value:
                    has_transport_row = True
                    transport_row_num = row_num
                    break
        if has_transport_row:
            break
    return {'file_exists': True, 'has_transport_row': has_transport_row, 'transport_row_index': transport_row_num}

def _rule_6_1_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Строка "Перемещения" найдена: {extracted_data['has_transport_row']}\nИндекс строки: {extracted_data.get('transport_row_index', 'не найдена')}'
    return _build_validator_prompt(rule, fact_block)

RULE_6_1 = _make_validator_module(
    rule_index='6.1',
    rule_title='Наличие строки "Перемещения" в блоке "Расчет ВПП"',
    load_data=_rule_6_1_load_data,
    extract_relevant_data=_rule_6_1_extract_relevant_data,
    build_prompt=_rule_6_1_build_prompt,
)


# ==============================================================================
# RULE 7.1 — Наличие и заполнение листа "КПСЦ"
# ==============================================================================
def _rule_7_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    files_exist, data_content = _load_existing_jsons(
        parser_outputs_dir,
        {
            'header': 'kpsc_header_v2.json',
            'table1': 'kpsc_table1_v2.json',
        },
    )
    return {'files_exist': files_exist, 'data_content': data_content}

def _rule_7_1_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    files_exist = data['files_exist']
    data_content = data['data_content']
    has_data = False
    if 'header' in data_content:
        fields = data_content['header'].get('fields', {})
        if any((v for v in fields.values() if v)):
            has_data = True
    if 'table1' in data_content and (not has_data):
        rows = data_content['table1'].get('rows', [])
        if rows:
            has_data = True
    return {'files_status': str(files_exist), 'has_data': has_data, 'files_count': sum(files_exist.values())}

def _rule_7_1_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Статус файлов парсинга: {extracted_data['files_status']}\nКоличество найденных файлов: {extracted_data['files_count']}\nЕсть ли данные в файлах: {extracted_data['has_data']}'
    return _build_validator_prompt(rule, fact_block)

RULE_7_1 = _make_validator_module(
    rule_index='7.1',
    rule_title='Наличие и заполнение листа "КПСЦ"',
    load_data=_rule_7_1_load_data,
    extract_relevant_data=_rule_7_1_extract_relevant_data,
    build_prompt=_rule_7_1_build_prompt,
)


# ==============================================================================
# RULE 7.2 — Наличие и заполнение листа "Условные обозначения"
# ==============================================================================
def _rule_7_2_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'legend_v2.json')

def _rule_7_2_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_entries': False, 'entries_count': 0}
    entries = data['data'].get('entries', [])
    has_text = any((entry.get('text') for entry in entries))
    return {'file_exists': True, 'has_entries': len(entries) > 0, 'entries_count': len(entries), 'has_text': has_text}

def _rule_7_2_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Файл существует: {extracted_data['file_exists']}\nЕсть записи (entries): {extracted_data['has_entries']}\nКоличество записей: {extracted_data['entries_count']}\nЕсть заполненный текст: {extracted_data.get('has_text', False)}'
    return _build_validator_prompt(rule, fact_block)

RULE_7_2 = _make_validator_module(
    rule_index='7.2',
    rule_title='Наличие и заполнение листа "Условные обозначения"',
    load_data=_rule_7_2_load_data,
    extract_relevant_data=_rule_7_2_extract_relevant_data,
    build_prompt=_rule_7_2_build_prompt,
)


# ==============================================================================
# RULE 7.3 — Наличие и заполнение листа "Показатели"
# ==============================================================================
def _rule_7_3_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'pokazateli_v3.json')

def _rule_7_3_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    has_data = any((any((cell.get('value') for cell in row.get('cells', []))) for row in rows))
    return {'file_exists': True, 'has_rows': len(rows) > 0, 'rows_count': len(rows), 'has_data': has_data}

def _rule_7_3_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Файл существует: {extracted_data['file_exists']}\nЕсть строки (rows): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}\nЕсть данные в ячейках: {extracted_data.get('has_data', False)}'
    return _build_validator_prompt(rule, fact_block)

RULE_7_3 = _make_validator_module(
    rule_index='7.3',
    rule_title='Наличие и заполнение листа "Показатели"',
    load_data=_rule_7_3_load_data,
    extract_relevant_data=_rule_7_3_extract_relevant_data,
    build_prompt=_rule_7_3_build_prompt,
)


# ==============================================================================
# RULE 7.4 — Наличие и заполнение листа "Оцифровка потерь КПСЦ"
# ==============================================================================
def _rule_7_4_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'ocifrovka_poteri_v2.json')

def _rule_7_4_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    return {'file_exists': True, 'has_rows': len(rows) >= 2, 'rows_count': len(rows)}

def _rule_7_4_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Файл существует: {extracted_data['file_exists']}\nЕсть строки (минимум 2): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}'
    return _build_validator_prompt(rule, fact_block)

RULE_7_4 = _make_validator_module(
    rule_index='7.4',
    rule_title='Наличие и заполнение листа "Оцифровка потерь КПСЦ"',
    load_data=_rule_7_4_load_data,
    extract_relevant_data=_rule_7_4_extract_relevant_data,
    build_prompt=_rule_7_4_build_prompt,
)


# ==============================================================================
# RULE 7.5 — Наличие и заполнение листа "ПА-1"
# ==============================================================================
def _rule_7_5_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    files_exist, data_content = _load_existing_jsons(
        parser_outputs_dir,
        {
            'table': 'pa1_table_v1.json',
            'chart': 'pa1_chart_v3.json',
        },
    )
    return {'files_exist': files_exist, 'data_content': data_content}

def _rule_7_5_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    files_exist = data['files_exist']
    data_content = data['data_content']
    has_data = len(data_content) > 0
    return {'files_status': str(files_exist), 'files_count': sum(files_exist.values()), 'has_data': has_data}

def _rule_7_5_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Статус файлов: {extracted_data['files_status']}\nКоличество найденных файлов: {extracted_data['files_count']}\nЕсть данные: {extracted_data['has_data']}'
    return _build_validator_prompt(rule, fact_block)

RULE_7_5 = _make_validator_module(
    rule_index='7.5',
    rule_title='Наличие и заполнение листа "ПА-1"',
    load_data=_rule_7_5_load_data,
    extract_relevant_data=_rule_7_5_extract_relevant_data,
    build_prompt=_rule_7_5_build_prompt,
)


# ==============================================================================
# RULE 7.6 — Наличие и заполнение листа "Диаграмма Спагетти" или "ДС"
# ==============================================================================
def _rule_7_6_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'spaghetti_sheet_v2.json')

def _rule_7_6_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    return {'file_exists': True, 'has_rows': len(rows) > 0, 'rows_count': len(rows)}

def _rule_7_6_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Файл существует: {extracted_data['file_exists']}\nЕсть строки (rows): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}'
    return _build_validator_prompt(rule, fact_block)

RULE_7_6 = _make_validator_module(
    rule_index='7.6',
    rule_title='Наличие и заполнение листа "Диаграмма Спагетти" или "ДС"',
    load_data=_rule_7_6_load_data,
    extract_relevant_data=_rule_7_6_extract_relevant_data,
    build_prompt=_rule_7_6_build_prompt,
)


# ==============================================================================
# RULE 7.7 — Наличие и заполнение листа "Перечень проблем по Диаграмме Спагетти"
# ==============================================================================
def _rule_7_7_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'spaghetti_problems_v1.json')

def _rule_7_7_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    return {'file_exists': True, 'has_rows': len(rows) >= 2, 'rows_count': len(rows)}

def _rule_7_7_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Файл существует: {extracted_data['file_exists']}\nЕсть строки (минимум 2): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}'
    return _build_validator_prompt(rule, fact_block)

RULE_7_7 = _make_validator_module(
    rule_index='7.7',
    rule_title='Наличие и заполнение листа "Перечень проблем по Диаграмме Спагетти"',
    load_data=_rule_7_7_load_data,
    extract_relevant_data=_rule_7_7_extract_relevant_data,
    build_prompt=_rule_7_7_build_prompt,
)


# ==============================================================================
# RULE 7.8 — Наличие и заполнение листа "Расчет такта" / "Расчет времени такта"
# ==============================================================================
def _rule_7_8_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    return _load_optional_json(parser_outputs_dir, 'kpsc_header_v2.json', exists_key='header_exists')

def _rule_7_8_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['header_exists']:
        return {'header_exists': False, 'takt_time': None}
    takt_time = data['data'].get('fields', {}).get('takt_time', '')
    return {'header_exists': True, 'takt_time': takt_time, 'has_takt_time': bool(takt_time)}

def _rule_7_8_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Файл заголовка существует: {extracted_data['header_exists']}\nПоле takt_time: "{extracted_data.get('takt_time', '')}"\nПоле заполнено: {extracted_data.get('has_takt_time', False)}'
    return _build_validator_prompt(rule, fact_block)

RULE_7_8 = _make_validator_module(
    rule_index='7.8',
    rule_title='Наличие и заполнение листа "Расчет такта" / "Расчет времени такта"',
    load_data=_rule_7_8_load_data,
    extract_relevant_data=_rule_7_8_extract_relevant_data,
    build_prompt=_rule_7_8_build_prompt,
)


# ==============================================================================
# RULE 8.1 — Соответствие единиц измерения между листами "Показатели" и "КПСЦ"
# ==============================================================================
def _rule_8_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    pokazateli_file = parser_outputs_dir / 'pokazateli_v3.json'
    kpsc_table_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not pokazateli_file.exists() or not kpsc_table_file.exists():
        return {'pokazateli_exists': pokazateli_file.exists(), 'kpsc_exists': kpsc_table_file.exists(), 'pokazateli_data': None, 'kpsc_data': None}
    loaded = _load_required_jsons(parser_outputs_dir, {'pokazateli_data': 'pokazateli_v3.json', 'kpsc_data': 'kpsc_table1_v2.json'})
    pokazateli_data = loaded['pokazateli_data']
    kpsc_data = loaded['kpsc_data']
    return {'pokazateli_exists': True, 'kpsc_exists': True, 'pokazateli_data': pokazateli_data, 'kpsc_data': kpsc_data}

def _rule_8_1_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['pokazateli_exists'] or not data['kpsc_exists']:
        return {'files_exist': False, 'mismatches': []}
    pokazateli_rows = data['pokazateli_data'].get('rows', [])
    pokazateli_units = {}
    for row in pokazateli_rows:
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(3, '')
        if not pokazatel_name or pokazatel_name == 'Показатель':
            continue
        unit = cells_dict.get(4, '')
        if unit:
            pokazateli_units[str(pokazatel_name).strip().lower()] = str(unit).strip()
    kpsc_rows = data['kpsc_data'].get('rows', [])
    kpsc_units = {}
    for i, row in enumerate(kpsc_rows):
        if i == 0:
            continue
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(2, '') or cells_dict.get(3, '')
        if not pokazatel_name:
            continue
        unit = cells_dict.get(4, '') or cells_dict.get(27, '')
        if unit and pokazatel_name:
            kpsc_units[str(pokazatel_name).strip().lower()] = str(unit).strip()
    mismatches = []
    for pokazatel, unit_pokazateli in pokazateli_units.items():
        if pokazatel in kpsc_units:
            unit_kpsc = kpsc_units[pokazatel]
            if unit_pokazateli.lower() != unit_kpsc.lower():
                mismatches.append({'pokazatel': pokazatel, 'unit_pokazateli': unit_pokazateli, 'unit_kpsc': unit_kpsc})
    return {'files_exist': True, 'pokazateli_count': len(pokazateli_units), 'kpsc_count': len(kpsc_units), 'pokazateli_units': pokazateli_units, 'kpsc_units': kpsc_units, 'mismatches': mismatches, 'has_mismatches': len(mismatches) > 0}

def _rule_8_1_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Файлы существуют: {extracted_data['files_exist']}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "ПОКАЗАТЕЛИ" (всего {extracted_data.get('pokazateli_count', 0)}):\n{json.dumps(extracted_data.get('pokazateli_units', {}), ensure_ascii=False, indent=2)}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "КПСЦ" БЛОК "РАСЧЕТ ВПП" (всего {extracted_data.get('kpsc_count', 0)}):\n{json.dumps(extracted_data.get('kpsc_units', {}), ensure_ascii=False, indent=2)}\n\nНЕСООТВЕТСТВИЯ ЕДИНИЦ ИЗМЕРЕНИЯ:\n{json.dumps(extracted_data.get('mismatches', []), ensure_ascii=False, indent=2)}\n\nЕсть несоответствия: {extracted_data.get('has_mismatches', False)}'
    return _build_validator_prompt(rule, fact_block)

RULE_8_1 = _make_validator_module(
    rule_index='8.1',
    rule_title='Соответствие единиц измерения между листами "Показатели" и "КПСЦ"',
    load_data=_rule_8_1_load_data,
    extract_relevant_data=_rule_8_1_extract_relevant_data,
    build_prompt=_rule_8_1_build_prompt,
)


# ==============================================================================
# RULE 8.2 — Соответствие значений из "Показатели" с колонкой ИТОГО в "КПСЦ"
# ==============================================================================
def _rule_8_2_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    pokazateli_file = parser_outputs_dir / 'pokazateli_v3.json'
    kpsc_table_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not pokazateli_file.exists() or not kpsc_table_file.exists():
        return {'pokazateli_exists': pokazateli_file.exists(), 'kpsc_exists': kpsc_table_file.exists(), 'pokazateli_data': None, 'kpsc_data': None}
    loaded = _load_required_jsons(parser_outputs_dir, {'pokazateli_data': 'pokazateli_v3.json', 'kpsc_data': 'kpsc_table1_v2.json'})
    pokazateli_data = loaded['pokazateli_data']
    kpsc_data = loaded['kpsc_data']
    return {'pokazateli_exists': True, 'kpsc_exists': True, 'pokazateli_data': pokazateli_data, 'kpsc_data': kpsc_data}

def _rule_8_2_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['pokazateli_exists'] or not data['kpsc_exists']:
        return {'files_exist': False, 'mismatches': []}
    pokazateli_rows = data['pokazateli_data'].get('rows', [])
    pokazateli_values = {}
    for row in pokazateli_rows:
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(3, '')
        if not pokazatel_name or pokazatel_name == 'Показатель':
            continue
        value_e = cells_dict.get(5)
        value_g = cells_dict.get(7)
        value = value_e if value_e is not None else value_g
        if value is not None:
            try:
                pokazateli_values[str(pokazatel_name).strip().lower()] = float(value)
            except (ValueError, TypeError):
                pass
    kpsc_rows = data['kpsc_data'].get('rows', [])
    kpsc_itogo_values = {}
    itogo_col = None
    if kpsc_rows:
        header_cells = {cell.get('col'): cell.get('value') for cell in kpsc_rows[0].get('cells', [])}
        for col_idx, val in header_cells.items():
            if isinstance(val, str) and 'итого' in val.lower():
                itogo_col = col_idx
                break
    for i, row in enumerate(kpsc_rows):
        if i == 0:
            continue
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(2, '') or cells_dict.get(3, '')
        if not pokazatel_name:
            continue
        itogo_value = cells_dict.get(itogo_col) if itogo_col else None
        if itogo_value is not None and pokazatel_name:
            try:
                kpsc_itogo_values[str(pokazatel_name).strip().lower()] = float(itogo_value)
            except (ValueError, TypeError):
                pass
    mismatches = []
    tolerance = 0.01
    for pokazatel, value_pokazateli in pokazateli_values.items():
        if pokazatel in kpsc_itogo_values:
            value_kpsc = kpsc_itogo_values[pokazatel]
            if abs(value_pokazateli - value_kpsc) > tolerance:
                mismatches.append({'pokazatel': pokazatel, 'value_pokazateli': value_pokazateli, 'value_kpsc_itogo': value_kpsc, 'difference': abs(value_pokazateli - value_kpsc)})
    return {'files_exist': True, 'pokazateli_count': len(pokazateli_values), 'kpsc_count': len(kpsc_itogo_values), 'pokazateli_values': pokazateli_values, 'kpsc_itogo_values': kpsc_itogo_values, 'mismatches': mismatches, 'has_mismatches': len(mismatches) > 0}

def _rule_8_2_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Файлы существуют: {extracted_data['files_exist']}\n\nЗНАЧЕНИЯ ИЗ ЛИСТА "ПОКАЗАТЕЛИ" (всего {extracted_data.get('pokazateli_count', 0)}):\n{json.dumps(extracted_data.get('pokazateli_values', {}), ensure_ascii=False, indent=2)}\n\nЗНАЧЕНИЯ ИТОГО ИЗ ЛИСТА "КПСЦ" БЛОК "РАСЧЕТ ВПП" (всего {extracted_data.get('kpsc_count', 0)}):\n{json.dumps(extracted_data.get('kpsc_itogo_values', {}), ensure_ascii=False, indent=2)}\n\nНЕСООТВЕТСТВИЯ ЗНАЧЕНИЙ (погрешность > 0.01):\n{json.dumps(extracted_data.get('mismatches', []), ensure_ascii=False, indent=2)}\n\nЕсть несоответствия: {extracted_data.get('has_mismatches', False)}'
    return _build_validator_prompt(rule, fact_block)

RULE_8_2 = _make_validator_module(
    rule_index='8.2',
    rule_title='Соответствие значений из "Показатели" с колонкой ИТОГО в "КПСЦ"',
    load_data=_rule_8_2_load_data,
    extract_relevant_data=_rule_8_2_extract_relevant_data,
    build_prompt=_rule_8_2_build_prompt,
)


# ==============================================================================
# RULE 8.3 — Соответствие названий показателей между листами "Показатели" и "КПСЦ"
# ==============================================================================
def _rule_8_3_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    pokazateli_file = parser_outputs_dir / 'pokazateli_v3.json'
    kpsc_table_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not pokazateli_file.exists() or not kpsc_table_file.exists():
        return {'pokazateli_exists': pokazateli_file.exists(), 'kpsc_exists': kpsc_table_file.exists(), 'pokazateli_data': None, 'kpsc_data': None}
    loaded = _load_required_jsons(parser_outputs_dir, {'pokazateli_data': 'pokazateli_v3.json', 'kpsc_data': 'kpsc_table1_v2.json'})
    pokazateli_data = loaded['pokazateli_data']
    kpsc_data = loaded['kpsc_data']
    return {'pokazateli_exists': True, 'kpsc_exists': True, 'pokazateli_data': pokazateli_data, 'kpsc_data': kpsc_data}

def _rule_8_3_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['pokazateli_exists'] or not data['kpsc_exists']:
        return {'files_exist': False, 'missing_indicators': []}
    pokazateli_rows = data['pokazateli_data'].get('rows', [])
    pokazateli_indicators = []
    for row in pokazateli_rows:
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(3, '')
        if not pokazatel_name or pokazatel_name == 'Показатель':
            continue
        pokazateli_indicators.append(str(pokazatel_name).strip())
    kpsc_rows = data['kpsc_data'].get('rows', [])
    kpsc_indicators = []
    kpsc_indicators_normalized = set()
    for i, row in enumerate(kpsc_rows):
        if i == 0:
            continue
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(2, '') or cells_dict.get(3, '')
        if pokazatel_name:
            original_name = str(pokazatel_name).strip()
            kpsc_indicators.append(original_name)
            kpsc_indicators_normalized.add(original_name.lower())
    ALIASES = {'время протекания процесса': ['впп'], 'выработка': ['выпуск продукции'], 'объем выпускаемой продукции': ['выпуск продукции']}
    missing_indicators = []
    for indicator in pokazateli_indicators:
        ind_lower = indicator.lower()
        found = ind_lower in kpsc_indicators_normalized or any((ind_lower in kpsc_name for kpsc_name in kpsc_indicators_normalized))
        if not found:
            for alias in ALIASES.get(ind_lower, []):
                if any((alias in kpsc_name for kpsc_name in kpsc_indicators_normalized)):
                    found = True
                    break
        if not found:
            missing_indicators.append(indicator)
    return {'files_exist': True, 'pokazateli_count': len(pokazateli_indicators), 'kpsc_count': len(kpsc_indicators), 'pokazateli_indicators': pokazateli_indicators, 'kpsc_indicators': kpsc_indicators, 'missing_indicators': missing_indicators, 'has_missing': len(missing_indicators) > 0}

def _rule_8_3_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    fact_block = f'Файлы существуют: {extracted_data['files_exist']}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "ПОКАЗАТЕЛИ" (всего {extracted_data.get('pokazateli_count', 0)}):\n{json.dumps(extracted_data.get('pokazateli_indicators', []), ensure_ascii=False, indent=2)}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "КПСЦ" БЛОК "РАСЧЕТ ВПП" (всего {extracted_data.get('kpsc_count', 0)}):\n{json.dumps(extracted_data.get('kpsc_indicators', []), ensure_ascii=False, indent=2)}\n\nОТСУТСТВУЮЩИЕ ПОКАЗАТЕЛИ (есть в "Показатели", но нет в "КПСЦ"):\n{json.dumps(extracted_data.get('missing_indicators', []), ensure_ascii=False, indent=2)}\n\nЕсть отсутствующие показатели: {extracted_data.get('has_missing', False)}'
    return _build_validator_prompt(rule, fact_block)

RULE_8_3 = _make_validator_module(
    rule_index='8.3',
    rule_title='Соответствие названий показателей между листами "Показатели" и "КПСЦ"',
    load_data=_rule_8_3_load_data,
    extract_relevant_data=_rule_8_3_extract_relevant_data,
    build_prompt=_rule_8_3_build_prompt,
)


# START_VALIDATOR_DISPATCH
# PURPOSE: Маппинг rule_index (строка из validation_rules.json) → SimpleNamespace-обёртка валидатора. Используется KPSC-раннером для выбора валидатора по ID правила.
# KEYWORDS: dispatch, rule-index.
KPSC_VALIDATOR_MODULE_DISPATCH: Dict[str, SimpleNamespace] = {
    '1.1': RULE_1_1,
    '1.2': RULE_1_2,
    '1.3': RULE_1_3,
    '1.4': RULE_1_4,
    '1.5': RULE_1_5,
    '1.6': RULE_1_6,
    '1.7': RULE_1_7,
    '2.1': RULE_2_1,
    '2.2': RULE_2_2,
    '2.3': RULE_2_3,
    '4.1': RULE_4_1,
    '4.2': RULE_4_2,
    '5.1': RULE_5_1,
    '6.1': RULE_6_1,
    '7.1': RULE_7_1,
    '7.2': RULE_7_2,
    '7.3': RULE_7_3,
    '7.4': RULE_7_4,
    '7.5': RULE_7_5,
    '7.6': RULE_7_6,
    '7.7': RULE_7_7,
    '7.8': RULE_7_8,
    '8.1': RULE_8_1,
    '8.2': RULE_8_2,
    '8.3': RULE_8_3,
}
# END_VALIDATOR_DISPATCH


# START_VALIDATOR_RUNNER
# PURPOSE: Универсальный runner одного валидатора — использует module_ns interface (RULE_INDEX, load_rule, load_data, extract_relevant_data, build_prompt, call_llm, save_result). Один на все 25 KPSC-валидаторов.
def run_kpsc_validator_module(module_ns: SimpleNamespace, parser_outputs_dir: Path, output_file: Path) -> Dict[str, Any]:
    """
    Назначение:
        Выполняет стандартный 6-шаговый пайплайн валидатора: загрузка правила,
        загрузка данных парсера, извлечение релевантных полей, построение промпта,
        вызов LLM, сохранение результата.

    Вход:
        module_ns: SimpleNamespace валидатора (из `KPSC_VALIDATOR_MODULE_DISPATCH`).
        parser_outputs_dir: Директория с JSON-ами парсера.
        output_file: Путь для записи JSON с результатом.

    Выход:
        dict с ключами rule_index, rule_title, target_document, status, discrepancy.
    """
    api_key = os.environ.get("OPENAI_API_KEY", "dummy")
    rule = module_ns.load_rule(module_ns.RULE_INDEX)
    data = module_ns.load_data(parser_outputs_dir)
    extracted = module_ns.extract_relevant_data(data)
    prompt = module_ns.build_prompt(extracted, rule)
    result = module_ns.call_llm(prompt, api_key)
    module_ns.save_result(result, output_file)
    return result
# END_VALIDATOR_RUNNER



# ==============================================================================
#                            KPSC SPECIAL RUNNER
# Публичный entrypoint `run_kpsc_special` для регистрации в `SPECIAL_ENGINE_RUNNERS`.
# Оркестрация: load_config → run_parsers (9 шт) → load_rules → run_validators_parallel → excel_report.
# ==============================================================================

# START_KPSC_HELPERS
def _load_kpsc_config() -> Dict[str, Any]:
    """Читает doc_configs/kpsc/config.json (model, max_workers, llm_base_url)."""
    with open(_DOC_CONFIGS_DIR / "kpsc" / "config.json", "r", encoding="utf-8") as f:
        return json.load(f)


def _load_kpsc_rules(rules_path: Path) -> List[Dict[str, Any]]:
    """Загружает все правила из validation_rules.json."""
    with open(rules_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data["rules"]


def _kpsc_core_sheet_missing(parser_results: Dict[str, Dict[str, Any]]) -> bool:
    """
    Назначение:
        True если в xlsx нет основного листа КПСЦ — тогда смысла гонять валидаторы
        нет, можно сразу выдать FAIL по всем правилам без вызова LLM.

    Логика:
        Проверяет что оба парсера `parse_kpsc_header` и `parse_kpsc_table1`
        упали с сообщением «Лист КПСЦ не найден».
    """
    required = ("parse_kpsc_header", "parse_kpsc_table1")
    hits = 0
    for parser_name in required:
        parser_state = parser_results.get(parser_name, {})
        error_text = str(parser_state.get("error", ""))
        if parser_state.get("status") == "error" and "Лист КПСЦ не найден" in error_text:
            hits += 1
    return hits == len(required)


def _kpsc_fast_fail_results(
    rules: List[Dict[str, Any]],
    output_dir: Path,
    discrepancy: str,
) -> List[Tuple[Dict[str, Any], Dict[str, Any], float]]:
    """Генерирует FAIL-результаты для всех правил без вызова LLM (используется при отсутствии листа КПСЦ)."""
    output_dir.mkdir(parents=True, exist_ok=True)
    results: List[Tuple[Dict[str, Any], Dict[str, Any], float]] = []
    for rule in rules:
        result_data = {
            "rule_index": rule["rule_index"],
            "rule_title": rule.get("rule_title", ""),
            "status": "FAIL",
            "discrepancy": discrepancy,
        }
        output_file = output_dir / f"validate_{rule['rule_index'].replace('.', '_')}.json"
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(result_data, f, ensure_ascii=False, indent=2)
        results.append((rule, result_data, 0.0))
    return results


def _run_kpsc_parsers(xlsx_path: Path, output_dir: Path) -> Dict[str, Any]:
    """Прогоняет все 9 парсеров КПСЦ. Каждый парсер сам пишет JSON в output_dir."""
    output_dir.mkdir(parents=True, exist_ok=True)
    results: Dict[str, Any] = {}
    for short_name, parser_fn in KPSC_PARSER_MODULE_DISPATCH.items():
        try:
            parser_fn(xlsx_path, output_dir)
            results[short_name] = {"status": "ok"}
            print(f"  [OK] {short_name}")
        except Exception as e:
            results[short_name] = {"status": "error", "error": str(e)}
            print(f"  [ERR] {short_name}: {e}")
    return results


def _run_kpsc_single_validator(
    rule: Dict[str, Any],
    parser_outputs_dir: Path,
    output_dir: Path,
    rules_path: Path,
) -> Tuple[Dict[str, Any], Dict[str, Any], float]:
    """Запускает один валидатор по rule_index. Возвращает (rule, result, duration)."""
    rule_index = rule["rule_index"]
    start = time.time()
    module_ns = KPSC_VALIDATOR_MODULE_DISPATCH.get(rule_index)
    if module_ns is None:
        return rule, {
            "rule_index": rule_index,
            "status": "MISSING",
            "discrepancy": f"Валидатор для правила {rule_index} не зарегистрирован",
        }, 0.0
    os.environ["VALIDATION_RULES_PATH"] = str(rules_path)
    output_file = output_dir / f"validate_{rule_index.replace('.', '_')}.json"
    try:
        result_data = run_kpsc_validator_module(module_ns, parser_outputs_dir, output_file)
        return rule, result_data, time.time() - start
    except FileNotFoundError as e:
        result_data = {
            "rule_index": rule_index,
            "status": "FAIL",
            "discrepancy": f"Недостаточно данных парсинга для правила {rule_index}: {e}",
        }
        output_file.parent.mkdir(parents=True, exist_ok=True)
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(result_data, f, ensure_ascii=False, indent=2)
        return rule, result_data, time.time() - start
    except Exception as e:
        return rule, {
            "rule_index": rule_index,
            "status": "ERROR",
            "discrepancy": f"Исключение: {e}",
        }, time.time() - start


def _run_kpsc_validators_parallel(
    rules: List[Dict[str, Any]],
    parser_outputs_dir: Path,
    output_dir: Path,
    rules_path: Path,
    max_workers: int = 5,
) -> List[Tuple[Dict[str, Any], Dict[str, Any], float]]:
    """Параллельный прогон всех валидаторов через ThreadPoolExecutor."""
    output_dir.mkdir(parents=True, exist_ok=True)
    results: List[Tuple[Dict[str, Any], Dict[str, Any], float]] = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(
                _run_kpsc_single_validator, rule, parser_outputs_dir, output_dir, rules_path,
            ): rule
            for rule in rules
        }
        for future in as_completed(futures):
            rule = futures[future]
            try:
                rule_data, result_data, duration = future.result()
                results.append((rule_data, result_data, duration))
                status = result_data.get("status", "?")
                emoji = "+" if status == "PASS" else "-" if status == "FAIL" else "!"
                print(f"  [{emoji}] [{len(results)}/{len(rules)}] {result_data.get('rule_index', '?')}: {status} ({duration:.1f}s)")
            except Exception as e:
                results.append((rule, {
                    "rule_index": rule["rule_index"],
                    "status": "ERROR",
                    "discrepancy": f"Future exception: {e}",
                }, 0.0))
    return results


def _create_kpsc_excel_report(
    results: List[Tuple[Dict[str, Any], Dict[str, Any], float]],
    report_path: Path,
) -> pd.DataFrame:
    """Собирает Excel-отчёт по результатам KPSC-валидации."""
    rows = []
    for rule, result_data, duration in results:
        rows.append({
            "rule_index": rule["rule_index"],
            "section": rule.get("section", ""),
            "rule_title": rule.get("rule_title", ""),
            "status": result_data.get("status", "UNKNOWN"),
            "discrepancy": result_data.get("discrepancy", ""),
            "duration_sec": round(duration, 2),
        })
    df = pd.DataFrame(rows)
    df = df.sort_values("rule_index")
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Validation Results", index=False)
        worksheet = writer.sheets["Validation Results"]
        for idx, col in enumerate(df.columns):
            max_length = max(df[col].astype(str).apply(len).max(), len(col))
            worksheet.column_dimensions[get_column_letter(idx + 1)].width = min(max_length + 2, 50)
    return df
# END_KPSC_HELPERS


# START_KPSC_RUNNER
def run_kpsc_special(args):
    """
    Назначение:
        Публичный entrypoint для KPSC. Регистрируется в `SPECIAL_ENGINE_RUNNERS`
        в main.py. Оркестрирует: 9 парсеров → загрузка правил → параллельный
        прогон 25 валидаторов через `KPSC_VALIDATOR_MODULE_DISPATCH` → Excel-отчёт.

    Вход:
        args: argparse-Namespace со полями {target, model?, temperature?,
              parse_only?, rule_filter?, session_dir?}.

    Выход:
        AuditResult.

    Логика:
        1. Загружает config.json + готовит env (VALIDATION_RULES_PATH, LLM_*).
        2. Запускает 9 парсеров → JSON-файлы в parser_outputs/.
        3. Если основного листа КПСЦ нет — fast-fail без LLM.
        4. Иначе параллельно прогоняет валидаторы (max_workers из config).
        5. Пишет Excel-отчёт + AuditResult с violations.
    """
    from main import AuditResult  # late import: main.py импортирует этот модуль

    start_time = time.time()
    target_path = Path(args.target)
    config = _load_kpsc_config()
    max_workers = config.get("max_workers", 5)
    rules_path = _DOC_CONFIGS_DIR / "kpsc" / "validation_rules.json"
    session_dir = (
        Path(args.session_dir) if args.session_dir
        else _LOGS_RESULT_DIR / "kpsc" / f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    )
    session_dir.mkdir(parents=True, exist_ok=True)
    parser_outputs_dir = session_dir / "parser_outputs"
    validation_outputs_dir = session_dir / "validation_outputs"

    os.environ["VALIDATION_RULES_PATH"] = str(rules_path)
    os.environ["LLM_BASE_URL"] = config.get("llm_base_url", LLM_CONFIG.base_url)
    os.environ["LLM_MODEL"] = resolve_runtime_llm_model(config.get("model", LLM_CONFIG.default_model))
    os.environ.setdefault("OPENAI_API_KEY", "dummy")

    parser_results = _run_kpsc_parsers(target_path, parser_outputs_dir)
    with open(session_dir / "parser_summary.json", "w", encoding="utf-8") as f:
        json.dump(parser_results, f, ensure_ascii=False, indent=2)

    if getattr(args, "parse_only", False):
        return AuditResult(
            doc_type="kpsc", session_dir=session_dir,
            target_path=str(target_path), duration_sec=time.time() - start_time,
        )

    rules = _load_kpsc_rules(rules_path)
    if args.rule_filter:
        rules = [r for r in rules if r["rule_index"] == str(args.rule_filter)]
        if not rules:
            raise ValueError(f"Правило {args.rule_filter} не найдено в kpsc")

    if _kpsc_core_sheet_missing(parser_results):
        print("  [WARN] Лист 'КПСЦ' отсутствует. Формируем FAIL-результаты без вызова LLM.")
        results = _kpsc_fast_fail_results(
            rules=rules,
            output_dir=validation_outputs_dir,
            discrepancy="Файл не похож на КПСЦ: отсутствует основной лист 'КПСЦ', поэтому правила КПСЦ завершены как FAIL без LLM-вызова.",
        )
    else:
        results = _run_kpsc_validators_parallel(
            rules=rules,
            parser_outputs_dir=parser_outputs_dir,
            output_dir=validation_outputs_dir,
            rules_path=rules_path,
            max_workers=max_workers,
        )
    df = _create_kpsc_excel_report(results, session_dir / "validation_report.xlsx")

    violations = []
    for rule, result_data, _dur in results:
        if result_data.get("status") == "FAIL":
            violations.append({
                "rule_index": result_data.get("rule_index", rule.get("rule_index", "?")),
                "rule_title": result_data.get("rule_title", rule.get("rule_title", "")),
                "section": rule.get("section", ""),
                "discrepancy": result_data.get("discrepancy", ""),
            })
    return AuditResult(
        violations=violations, doc_type="kpsc", session_dir=session_dir,
        duration_sec=time.time() - start_time, rules_checked=len(df),
        target_path=str(target_path),
    )
# END_KPSC_RUNNER
