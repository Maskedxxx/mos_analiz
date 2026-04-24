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
from pathlib import Path
from types import SimpleNamespace
from typing import Any, Dict

from config.llm import LLM_CONFIG
from src.llm.client import call_llm
# END_IMPORTS


# START_CONSTANTS
# PURPOSE: Общие константы КПСЦ-валидаторов. Target-документ один для всех 25 правил (КПСЦ — один Excel-файл); system-prompt тоже один.
TARGET_DOC = "КПСЦ и Спагетти ТС_Предприятие.xlsx"
SYSTEM_PROMPT = "Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON."
# END_CONSTANTS


# START_COMMON_HELPERS
# PURPOSE: load_rule/call_llm/save_result — одинаковы у всех 25 валидаторов, лежат здесь в одной копии. В SimpleNamespace каждого правила эти же функции переиспользуются.
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
    with open(rules_file, "r", encoding="utf-8") as f:
        data = json.load(f)
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
# END_COMMON_HELPERS


# ==============================================================================
# RULE 1.1 — Наличие текста 'КПСЦ' в заголовке
# ==============================================================================
def _rule_1_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def _rule_1_1_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'title': data['fields'].get('title', '')}

def _rule_1_1_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nЗаголовок документа: "{extracted_data['title']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_1_1 = SimpleNamespace(
    RULE_INDEX='1.1',
    RULE_TITLE="Наличие текста 'КПСЦ' в заголовке",
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_1_1_load_data,
    extract_relevant_data=_rule_1_1_extract_relevant_data,
    build_prompt=_rule_1_1_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 1.2 — Наличие названия предприятия в формате ООО "наименование"
# ==============================================================================
def _rule_1_2_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def _rule_1_2_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'title': data['fields'].get('title', ''), 'organization': data['fields'].get('organization', ''), 'workbook': data.get('meta', {}).get('workbook', '')}

def _rule_1_2_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nЗаголовок документа: "{extracted_data['title']}"\nНазвание организации: "{extracted_data.get('organization', '')}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_1_2 = SimpleNamespace(
    RULE_INDEX='1.2',
    RULE_TITLE='Наличие названия предприятия в формате ООО "наименование"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_1_2_load_data,
    extract_relevant_data=_rule_1_2_extract_relevant_data,
    build_prompt=_rule_1_2_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 1.3 — Наличие названия потока в формате "имя потока"
# ==============================================================================
def _rule_1_3_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def _rule_1_3_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'flow_name': data['fields'].get('flow_name', '')}

def _rule_1_3_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nНазвание потока (flow_name): "{extracted_data['flow_name']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_1_3 = SimpleNamespace(
    RULE_INDEX='1.3',
    RULE_TITLE='Наличие названия потока в формате "имя потока"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_1_3_load_data,
    extract_relevant_data=_rule_1_3_extract_relevant_data,
    build_prompt=_rule_1_3_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 1.4 — Заполнение поля "Ответственный за поток" (ФИО)
# ==============================================================================
def _rule_1_4_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def _rule_1_4_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    responsible = data['fields'].get('responsible', '') or ''
    words = responsible.strip().split()
    return {'responsible': responsible, 'word_count': len(words), 'char_length': len(responsible.strip())}

def _rule_1_4_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nОтветственный за поток (responsible): "{extracted_data['responsible']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_1_4 = SimpleNamespace(
    RULE_INDEX='1.4',
    RULE_TITLE='Заполнение поля "Ответственный за поток" (ФИО)',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_1_4_load_data,
    extract_relevant_data=_rule_1_4_extract_relevant_data,
    build_prompt=_rule_1_4_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 1.5 — Заполнение даты разработки
# ==============================================================================
def _rule_1_5_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def _rule_1_5_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'date_developed': data['fields'].get('date_developed', '')}

def _rule_1_5_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nДата разработки (date_developed): "{extracted_data['date_developed']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_1_5 = SimpleNamespace(
    RULE_INDEX='1.5',
    RULE_TITLE='Заполнение даты разработки',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_1_5_load_data,
    extract_relevant_data=_rule_1_5_extract_relevant_data,
    build_prompt=_rule_1_5_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 1.6 — Заполнение даты реализации
# ==============================================================================
def _rule_1_6_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def _rule_1_6_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'date_implementation': data['fields'].get('date_implementation', '')}

def _rule_1_6_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nДата реализации (date_implementation): "{extracted_data['date_implementation']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_1_6 = SimpleNamespace(
    RULE_INDEX='1.6',
    RULE_TITLE='Заполнение даты реализации',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_1_6_load_data,
    extract_relevant_data=_rule_1_6_extract_relevant_data,
    build_prompt=_rule_1_6_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 1.7 — Заполнение поля "Кто составил" (ФИО)
# ==============================================================================
def _rule_1_7_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def _rule_1_7_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'compiled_by': data['fields'].get('compiled_by', '')}

def _rule_1_7_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nКто составил (compiled_by): "{extracted_data['compiled_by']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_1_7 = SimpleNamespace(
    RULE_INDEX='1.7',
    RULE_TITLE='Заполнение поля "Кто составил" (ФИО)',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_1_7_load_data,
    extract_relevant_data=_rule_1_7_extract_relevant_data,
    build_prompt=_rule_1_7_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 2.1 — Количество проблем в таблице "Оцифровка потерь"
# ==============================================================================
def _rule_2_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    ocifrovka_file = parser_outputs_dir / 'ocifrovka_poteri_v2.json'
    if not ocifrovka_file.exists():
        return {'file_exists': False, 'data': None}
    with open(ocifrovka_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

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
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nКоличество проблем: {extracted_data['problems_count']}\nНомера проблем: {extracted_data['problem_numbers']}\nДиапазон: с {extracted_data.get('min_number', 0)} по {extracted_data.get('max_number', 0)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\nЭто информационная проверка - нужно подтвердить, что найдены проблемы и извлечены номера.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_2_1 = SimpleNamespace(
    RULE_INDEX='2.1',
    RULE_TITLE='Количество проблем в таблице "Оцифровка потерь"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_2_1_load_data,
    extract_relevant_data=_rule_2_1_extract_relevant_data,
    build_prompt=_rule_2_1_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 2.2 — Последовательность номеров проблем (пропуски в нумерации)
# ==============================================================================
def _rule_2_2_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    ocifrovka_file = parser_outputs_dir / 'ocifrovka_poteri_v2.json'
    if not ocifrovka_file.exists():
        return {'file_exists': False, 'data': None}
    with open(ocifrovka_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

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
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nНайденные номера проблем: {extracted_data['problem_numbers']}\nПропущенные номера: {extracted_data['missing_numbers']}\nЕсть пропуски: {extracted_data['has_missing']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_2_2 = SimpleNamespace(
    RULE_INDEX='2.2',
    RULE_TITLE='Последовательность номеров проблем (пропуски в нумерации)',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_2_2_load_data,
    extract_relevant_data=_rule_2_2_extract_relevant_data,
    build_prompt=_rule_2_2_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 2.3 — Наличие описания для каждого номера проблемы
# ==============================================================================
def _rule_2_3_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    ocifrovka_file = parser_outputs_dir / 'ocifrovka_poteri_v2.json'
    if not ocifrovka_file.exists():
        return {'file_exists': False, 'data': None}
    with open(ocifrovka_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

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
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nПроблемы без описания: {extracted_data['problems_without_description']}\nЕсть проблемы без описания: {extracted_data['has_missing_descriptions']}\nКоличество проблем без описания: {extracted_data['missing_count']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_2_3 = SimpleNamespace(
    RULE_INDEX='2.3',
    RULE_TITLE='Наличие описания для каждого номера проблемы',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_2_3_load_data,
    extract_relevant_data=_rule_2_3_extract_relevant_data,
    build_prompt=_rule_2_3_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 4.1 — Все ячейки строки ВПП заполнены
# ==============================================================================
def _rule_4_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    table1_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not table1_file.exists():
        return {'file_exists': False, 'data': None}
    with open(table1_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

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
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nПустые ячейки в строке ВПП:\n{json.dumps(extracted_data.get('empty_cells', []), ensure_ascii=False, indent=2)}\n\nЕсть пустые ячейки: {extracted_data.get('has_empty', False)}\nКоличество пустых: {extracted_data.get('empty_count', 0)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_4_1 = SimpleNamespace(
    RULE_INDEX='4.1',
    RULE_TITLE='Все ячейки строки ВПП заполнены',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_4_1_load_data,
    extract_relevant_data=_rule_4_1_extract_relevant_data,
    build_prompt=_rule_4_1_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 4.2 — Сумма значений строки ВПП совпадает с итоговым значением
# ==============================================================================
def _rule_4_2_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    table1_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not table1_file.exists():
        return {'file_exists': False, 'data': None}
    with open(table1_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

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
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nЗначения ВПП по операциям:\n{json.dumps(extracted_data.get('vpp_values', []), ensure_ascii=False, indent=2)}\n\nВычисленная сумма: {extracted_data.get('calculated_sum', 0)}\nЗначение ИТОГО: {extracted_data.get('itogo_value', 0)}\nРасхождение: {extracted_data.get('difference', 0)}\nСуммы совпадают (±0.01): {extracted_data.get('matches', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_4_2 = SimpleNamespace(
    RULE_INDEX='4.2',
    RULE_TITLE='Сумма значений строки ВПП совпадает с итоговым значением',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_4_2_load_data,
    extract_relevant_data=_rule_4_2_extract_relevant_data,
    build_prompt=_rule_4_2_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 5.1 — Заполнение всех ячеек в столбце "Единицы измерения"
# ==============================================================================
def _rule_5_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    table1_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not table1_file.exists():
        return {'file_exists': False, 'data': None}
    with open(table1_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

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
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nСтроки с пустыми единицами измерения: {extracted_data['empty_units']}\nЕсть пустые ячейки: {extracted_data['has_empty_units']}\nКоличество пустых: {extracted_data['empty_count']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_5_1 = SimpleNamespace(
    RULE_INDEX='5.1',
    RULE_TITLE='Заполнение всех ячеек в столбце "Единицы измерения"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_5_1_load_data,
    extract_relevant_data=_rule_5_1_extract_relevant_data,
    build_prompt=_rule_5_1_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 6.1 — Наличие строки "Перемещения" в блоке "Расчет ВПП"
# ==============================================================================
def _rule_6_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    table1_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not table1_file.exists():
        return {'file_exists': False, 'data': None}
    with open(table1_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

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
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nСтрока "Перемещения" найдена: {extracted_data['has_transport_row']}\nИндекс строки: {extracted_data.get('transport_row_index', 'не найдена')}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_6_1 = SimpleNamespace(
    RULE_INDEX='6.1',
    RULE_TITLE='Наличие строки "Перемещения" в блоке "Расчет ВПП"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_6_1_load_data,
    extract_relevant_data=_rule_6_1_extract_relevant_data,
    build_prompt=_rule_6_1_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 7.1 — Наличие и заполнение листа "КПСЦ"
# ==============================================================================
def _rule_7_1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    table1_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    files_exist = {'kpsc_header_v2.json': header_file.exists(), 'kpsc_table1_v2.json': table1_file.exists()}
    data_content = {}
    if header_file.exists():
        with open(header_file, 'r', encoding='utf-8') as f:
            data_content['header'] = json.load(f)
    if table1_file.exists():
        with open(table1_file, 'r', encoding='utf-8') as f:
            data_content['table1'] = json.load(f)
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
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nСтатус файлов парсинга: {extracted_data['files_status']}\nКоличество найденных файлов: {extracted_data['files_count']}\nЕсть ли данные в файлах: {extracted_data['has_data']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_7_1 = SimpleNamespace(
    RULE_INDEX='7.1',
    RULE_TITLE='Наличие и заполнение листа "КПСЦ"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_7_1_load_data,
    extract_relevant_data=_rule_7_1_extract_relevant_data,
    build_prompt=_rule_7_1_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 7.2 — Наличие и заполнение листа "Условные обозначения"
# ==============================================================================
def _rule_7_2_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    legend_file = parser_outputs_dir / 'legend_v2.json'
    if not legend_file.exists():
        return {'file_exists': False, 'data': None}
    with open(legend_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'file_exists': True, 'data': data}

def _rule_7_2_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_entries': False, 'entries_count': 0}
    entries = data['data'].get('entries', [])
    has_text = any((entry.get('text') for entry in entries))
    return {'file_exists': True, 'has_entries': len(entries) > 0, 'entries_count': len(entries), 'has_text': has_text}

def _rule_7_2_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nЕсть записи (entries): {extracted_data['has_entries']}\nКоличество записей: {extracted_data['entries_count']}\nЕсть заполненный текст: {extracted_data.get('has_text', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_7_2 = SimpleNamespace(
    RULE_INDEX='7.2',
    RULE_TITLE='Наличие и заполнение листа "Условные обозначения"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_7_2_load_data,
    extract_relevant_data=_rule_7_2_extract_relevant_data,
    build_prompt=_rule_7_2_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 7.3 — Наличие и заполнение листа "Показатели"
# ==============================================================================
def _rule_7_3_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    pokazateli_file = parser_outputs_dir / 'pokazateli_v3.json'
    if not pokazateli_file.exists():
        return {'file_exists': False, 'data': None}
    with open(pokazateli_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'file_exists': True, 'data': data}

def _rule_7_3_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    has_data = any((any((cell.get('value') for cell in row.get('cells', []))) for row in rows))
    return {'file_exists': True, 'has_rows': len(rows) > 0, 'rows_count': len(rows), 'has_data': has_data}

def _rule_7_3_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nЕсть строки (rows): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}\nЕсть данные в ячейках: {extracted_data.get('has_data', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_7_3 = SimpleNamespace(
    RULE_INDEX='7.3',
    RULE_TITLE='Наличие и заполнение листа "Показатели"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_7_3_load_data,
    extract_relevant_data=_rule_7_3_extract_relevant_data,
    build_prompt=_rule_7_3_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 7.4 — Наличие и заполнение листа "Оцифровка потерь КПСЦ"
# ==============================================================================
def _rule_7_4_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    ocifrovka_file = parser_outputs_dir / 'ocifrovka_poteri_v2.json'
    if not ocifrovka_file.exists():
        return {'file_exists': False, 'data': None}
    with open(ocifrovka_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'file_exists': True, 'data': data}

def _rule_7_4_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    return {'file_exists': True, 'has_rows': len(rows) >= 2, 'rows_count': len(rows)}

def _rule_7_4_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nЕсть строки (минимум 2): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_7_4 = SimpleNamespace(
    RULE_INDEX='7.4',
    RULE_TITLE='Наличие и заполнение листа "Оцифровка потерь КПСЦ"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_7_4_load_data,
    extract_relevant_data=_rule_7_4_extract_relevant_data,
    build_prompt=_rule_7_4_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 7.5 — Наличие и заполнение листа "ПА-1"
# ==============================================================================
def _rule_7_5_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    pa1_table_file = parser_outputs_dir / 'pa1_table_v1.json'
    pa1_chart_file = parser_outputs_dir / 'pa1_chart_v3.json'
    files_exist = {'pa1_table_v1.json': pa1_table_file.exists(), 'pa1_chart_v3.json': pa1_chart_file.exists()}
    data_content = {}
    if pa1_table_file.exists():
        with open(pa1_table_file, 'r', encoding='utf-8') as f:
            data_content['table'] = json.load(f)
    if pa1_chart_file.exists():
        with open(pa1_chart_file, 'r', encoding='utf-8') as f:
            data_content['chart'] = json.load(f)
    return {'files_exist': files_exist, 'data_content': data_content}

def _rule_7_5_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    files_exist = data['files_exist']
    data_content = data['data_content']
    has_data = len(data_content) > 0
    return {'files_status': str(files_exist), 'files_count': sum(files_exist.values()), 'has_data': has_data}

def _rule_7_5_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nСтатус файлов: {extracted_data['files_status']}\nКоличество найденных файлов: {extracted_data['files_count']}\nЕсть данные: {extracted_data['has_data']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_7_5 = SimpleNamespace(
    RULE_INDEX='7.5',
    RULE_TITLE='Наличие и заполнение листа "ПА-1"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_7_5_load_data,
    extract_relevant_data=_rule_7_5_extract_relevant_data,
    build_prompt=_rule_7_5_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 7.6 — Наличие и заполнение листа "Диаграмма Спагетти" или "ДС"
# ==============================================================================
def _rule_7_6_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    spaghetti_file = parser_outputs_dir / 'spaghetti_sheet_v2.json'
    if not spaghetti_file.exists():
        return {'file_exists': False, 'data': None}
    with open(spaghetti_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'file_exists': True, 'data': data}

def _rule_7_6_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    return {'file_exists': True, 'has_rows': len(rows) > 0, 'rows_count': len(rows)}

def _rule_7_6_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nЕсть строки (rows): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_7_6 = SimpleNamespace(
    RULE_INDEX='7.6',
    RULE_TITLE='Наличие и заполнение листа "Диаграмма Спагетти" или "ДС"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_7_6_load_data,
    extract_relevant_data=_rule_7_6_extract_relevant_data,
    build_prompt=_rule_7_6_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 7.7 — Наличие и заполнение листа "Перечень проблем по Диаграмме Спагетти"
# ==============================================================================
def _rule_7_7_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    spaghetti_problems_file = parser_outputs_dir / 'spaghetti_problems_v1.json'
    if not spaghetti_problems_file.exists():
        return {'file_exists': False, 'data': None}
    with open(spaghetti_problems_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'file_exists': True, 'data': data}

def _rule_7_7_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    return {'file_exists': True, 'has_rows': len(rows) >= 2, 'rows_count': len(rows)}

def _rule_7_7_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nЕсть строки (минимум 2): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_7_7 = SimpleNamespace(
    RULE_INDEX='7.7',
    RULE_TITLE='Наличие и заполнение листа "Перечень проблем по Диаграмме Спагетти"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_7_7_load_data,
    extract_relevant_data=_rule_7_7_extract_relevant_data,
    build_prompt=_rule_7_7_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
)


# ==============================================================================
# RULE 7.8 — Наличие и заполнение листа "Расчет такта" / "Расчет времени такта"
# ==============================================================================
def _rule_7_8_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    if not header_file.exists():
        return {'header_exists': False, 'data': None}
    with open(header_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'header_exists': True, 'data': data}

def _rule_7_8_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['header_exists']:
        return {'header_exists': False, 'takt_time': None}
    takt_time = data['data'].get('fields', {}).get('takt_time', '')
    return {'header_exists': True, 'takt_time': takt_time, 'has_takt_time': bool(takt_time)}

def _rule_7_8_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл заголовка существует: {extracted_data['header_exists']}\nПоле takt_time: "{extracted_data.get('takt_time', '')}"\nПоле заполнено: {extracted_data.get('has_takt_time', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_7_8 = SimpleNamespace(
    RULE_INDEX='7.8',
    RULE_TITLE='Наличие и заполнение листа "Расчет такта" / "Расчет времени такта"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_7_8_load_data,
    extract_relevant_data=_rule_7_8_extract_relevant_data,
    build_prompt=_rule_7_8_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
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
    with open(pokazateli_file, 'r', encoding='utf-8') as f:
        pokazateli_data = json.load(f)
    with open(kpsc_table_file, 'r', encoding='utf-8') as f:
        kpsc_data = json.load(f)
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
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайлы существуют: {extracted_data['files_exist']}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "ПОКАЗАТЕЛИ" (всего {extracted_data.get('pokazateli_count', 0)}):\n{json.dumps(extracted_data.get('pokazateli_units', {}), ensure_ascii=False, indent=2)}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "КПСЦ" БЛОК "РАСЧЕТ ВПП" (всего {extracted_data.get('kpsc_count', 0)}):\n{json.dumps(extracted_data.get('kpsc_units', {}), ensure_ascii=False, indent=2)}\n\nНЕСООТВЕТСТВИЯ ЕДИНИЦ ИЗМЕРЕНИЯ:\n{json.dumps(extracted_data.get('mismatches', []), ensure_ascii=False, indent=2)}\n\nЕсть несоответствия: {extracted_data.get('has_mismatches', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_8_1 = SimpleNamespace(
    RULE_INDEX='8.1',
    RULE_TITLE='Соответствие единиц измерения между листами "Показатели" и "КПСЦ"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_8_1_load_data,
    extract_relevant_data=_rule_8_1_extract_relevant_data,
    build_prompt=_rule_8_1_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
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
    with open(pokazateli_file, 'r', encoding='utf-8') as f:
        pokazateli_data = json.load(f)
    with open(kpsc_table_file, 'r', encoding='utf-8') as f:
        kpsc_data = json.load(f)
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
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайлы существуют: {extracted_data['files_exist']}\n\nЗНАЧЕНИЯ ИЗ ЛИСТА "ПОКАЗАТЕЛИ" (всего {extracted_data.get('pokazateli_count', 0)}):\n{json.dumps(extracted_data.get('pokazateli_values', {}), ensure_ascii=False, indent=2)}\n\nЗНАЧЕНИЯ ИТОГО ИЗ ЛИСТА "КПСЦ" БЛОК "РАСЧЕТ ВПП" (всего {extracted_data.get('kpsc_count', 0)}):\n{json.dumps(extracted_data.get('kpsc_itogo_values', {}), ensure_ascii=False, indent=2)}\n\nНЕСООТВЕТСТВИЯ ЗНАЧЕНИЙ (погрешность > 0.01):\n{json.dumps(extracted_data.get('mismatches', []), ensure_ascii=False, indent=2)}\n\nЕсть несоответствия: {extracted_data.get('has_mismatches', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_8_2 = SimpleNamespace(
    RULE_INDEX='8.2',
    RULE_TITLE='Соответствие значений из "Показатели" с колонкой ИТОГО в "КПСЦ"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_8_2_load_data,
    extract_relevant_data=_rule_8_2_extract_relevant_data,
    build_prompt=_rule_8_2_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
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
    with open(pokazateli_file, 'r', encoding='utf-8') as f:
        pokazateli_data = json.load(f)
    with open(kpsc_table_file, 'r', encoding='utf-8') as f:
        kpsc_data = json.load(f)
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
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайлы существуют: {extracted_data['files_exist']}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "ПОКАЗАТЕЛИ" (всего {extracted_data.get('pokazateli_count', 0)}):\n{json.dumps(extracted_data.get('pokazateli_indicators', []), ensure_ascii=False, indent=2)}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "КПСЦ" БЛОК "РАСЧЕТ ВПП" (всего {extracted_data.get('kpsc_count', 0)}):\n{json.dumps(extracted_data.get('kpsc_indicators', []), ensure_ascii=False, indent=2)}\n\nОТСУТСТВУЮЩИЕ ПОКАЗАТЕЛИ (есть в "Показатели", но нет в "КПСЦ"):\n{json.dumps(extracted_data.get('missing_indicators', []), ensure_ascii=False, indent=2)}\n\nЕсть отсутствующие показатели: {extracted_data.get('has_missing', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

RULE_8_3 = SimpleNamespace(
    RULE_INDEX='8.3',
    RULE_TITLE='Соответствие названий показателей между листами "Показатели" и "КПСЦ"',
    TARGET_DOC=TARGET_DOC,
    load_rule=_load_rule,
    load_data=_rule_8_3_load_data,
    extract_relevant_data=_rule_8_3_extract_relevant_data,
    build_prompt=_rule_8_3_build_prompt,
    call_llm=_call_validator_llm,
    save_result=_save_result,
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
