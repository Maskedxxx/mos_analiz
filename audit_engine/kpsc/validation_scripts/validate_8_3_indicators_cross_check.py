#!/usr/bin/env python3
"""
Валидатор 8.3: Соответствие названий показателей между листами "Показатели" и "КПСЦ"
Источник: Колонка J, лист Показатели, ячейка J4
"""
import os
import json
import argparse
from pathlib import Path
from openai import OpenAI
from datetime import datetime

# Константы
RULE_INDEX = "8.3"
RULE_TITLE = "Соответствие названий показателей между листами \"Показатели\" и \"КПСЦ\""
TARGET_DOC = "КПСЦ и Спагетти ТС_Предприятие.xlsx"

def load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get("VALIDATION_RULES_PATH", str(Path(__file__).parent.parent / "validation_rules.json")))
    with open(rules_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    for rule in data["rules"]:
        if rule["rule_index"] == rule_index:
            return rule

    raise ValueError(f"Правило {rule_index} не найдено в validation_rules.json")

def load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    pokazateli_file = parser_outputs_dir / "pokazateli_v3.json"
    kpsc_table_file = parser_outputs_dir / "kpsc_table1_v2.json"

    if not pokazateli_file.exists() or not kpsc_table_file.exists():
        return {
            "pokazateli_exists": pokazateli_file.exists(),
            "kpsc_exists": kpsc_table_file.exists(),
            "pokazateli_data": None,
            "kpsc_data": None
        }

    with open(pokazateli_file, "r", encoding="utf-8") as f:
        pokazateli_data = json.load(f)

    with open(kpsc_table_file, "r", encoding="utf-8") as f:
        kpsc_data = json.load(f)

    return {
        "pokazateli_exists": True,
        "kpsc_exists": True,
        "pokazateli_data": pokazateli_data,
        "kpsc_data": kpsc_data
    }

def extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data["pokazateli_exists"] or not data["kpsc_exists"]:
        return {
            "files_exist": False,
            "missing_indicators": []
        }

    # Извлекаем названия показателей из листа "Показатели"
    pokazateli_rows = data["pokazateli_data"].get("rows", [])
    pokazateli_indicators = []

    for row in pokazateli_rows:
        cells_dict = {cell.get("col"): cell.get("value") for cell in row.get("cells", [])}

        # Пропускаем заголовок
        pokazatel_name = cells_dict.get(3, "")  # Колонка C - название показателя
        if not pokazatel_name or pokazatel_name == "Показатель":
            continue

        pokazateli_indicators.append(str(pokazatel_name).strip())

    # Извлекаем названия показателей из блока "Расчет ВПП" на листе КПСЦ
    kpsc_rows = data["kpsc_data"].get("rows", [])
    kpsc_indicators = []
    kpsc_indicators_normalized = set()

    for i, row in enumerate(kpsc_rows):
        if i == 0:  # Пропускаем заголовок
            continue

        cells_dict = {cell.get("col"): cell.get("value") for cell in row.get("cells", [])}

        # Колонка B или C - название показателя
        pokazatel_name = cells_dict.get(2, "") or cells_dict.get(3, "")
        if pokazatel_name:
            original_name = str(pokazatel_name).strip()
            kpsc_indicators.append(original_name)
            kpsc_indicators_normalized.add(original_name.lower())

    # Проверяем наличие показателей из "Показатели" в "КПСЦ"
    missing_indicators = []

    for indicator in pokazateli_indicators:
        if indicator.lower() not in kpsc_indicators_normalized:
            missing_indicators.append(indicator)

    return {
        "files_exist": True,
        "pokazateli_count": len(pokazateli_indicators),
        "kpsc_count": len(kpsc_indicators),
        "pokazateli_indicators": pokazateli_indicators,
        "kpsc_indicators": kpsc_indicators,
        "missing_indicators": missing_indicators,
        "has_missing": len(missing_indicators) > 0
    }

def build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f"""
Ты эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).

ТРЕБОВАНИЕ ЭКСПЕРТА:
{rule['requirement_expert']}

ТЕХНИЧЕСКОЕ ОПИСАНИЕ:
{rule['technical_description']}

КРИТЕРИИ ПРОВЕРКИ:
- Что проверять: {rule['validation_criteria']['what_to_check']}
- Условие успеха: {rule['validation_criteria']['success_condition']}
- Условие ошибки: {rule['validation_criteria']['error_condition']}

ФАКТИЧЕСКИЕ ДАННЫЕ:
Файлы существуют: {extracted_data['files_exist']}

ПОКАЗАТЕЛИ ИЗ ЛИСТА "ПОКАЗАТЕЛИ" (всего {extracted_data.get('pokazateli_count', 0)}):
{json.dumps(extracted_data.get('pokazateli_indicators', []), ensure_ascii=False, indent=2)}

ПОКАЗАТЕЛИ ИЗ ЛИСТА "КПСЦ" БЛОК "РАСЧЕТ ВПП" (всего {extracted_data.get('kpsc_count', 0)}):
{json.dumps(extracted_data.get('kpsc_indicators', []), ensure_ascii=False, indent=2)}

ОТСУТСТВУЮЩИЕ ПОКАЗАТЕЛИ (есть в "Показатели", но нет в "КПСЦ"):
{json.dumps(extracted_data.get('missing_indicators', []), ensure_ascii=False, indent=2)}

Есть отсутствующие показатели: {extracted_data.get('has_missing', False)}

ЗАДАНИЕ:
Проверь соответствие фактических данных требованию эксперта и критериям проверки.

ФОРМАТ ОТВЕТА (строго JSON):
{{
  "rule_index": "{rule['rule_index']}",
  "rule_title": "{rule['rule_title']}",
  "target_document": "{TARGET_DOC}",
  "status": "PASS или FAIL",
  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"
}}

Верни только JSON, без дополнительного текста.
"""
    return prompt

def call_llm(prompt: str, api_key: str) -> dict:
    """Вызов OpenAI API"""
    client = OpenAI(api_key=api_key)

    response = client.chat.completions.create(
        model="gpt-4.1-mini-2025-04-14",
        messages=[
            {"role": "system", "content": "Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON."},
            {"role": "user", "content": prompt}
        ],
        temperature=0,
        response_format={"type": "json_object"}
    )

    result_text = response.choices[0].message.content
    return json.loads(result_text)

def save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"✓ Результат сохранен: {output_file}")

def log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = "=" * 80
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"\n{separator}\n")
        f.write(f"[STEP {step_num}/6] {step_title}\n")
        f.write(f"Timestamp: {timestamp}\n")
        f.write(f"{separator}\n")
        f.write(f"{content}\n")

def main():
    parser = argparse.ArgumentParser(description=f"Валидатор {RULE_INDEX}: {RULE_TITLE}")
    parser.add_argument(
        "--parser-outputs",
        type=Path,
        required=True,
        help="Путь к папке с результатами парсинга (parser_outputs)"
    )
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Путь для сохранения результата валидации (JSON файл)"
    )
    parser.add_argument("--verbose", action="store_true", help="Включить детальное логирование")
    parser.add_argument("--log-file", type=Path, help="Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)")
    args = parser.parse_args()

    # Получение API ключа
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("Не найден OPENAI_API_KEY в переменных окружения")


    # Настройка логирования
    log_file = args.log_file
    if args.verbose and not log_file:
        log_dir = Path(__file__).parent.parent / "validation_logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f"validate_{RULE_INDEX.replace('.', '_')}.log"

    if args.verbose and log_file:
        with open(log_file, "w", encoding="utf-8") as f:
            f.write(f"ВАЛИДАТОР {RULE_INDEX}: {RULE_TITLE}\n")
            f.write(f"Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Целевой документ: {TARGET_DOC}\n")

    # Шаги валидации
    print(f"[1/6] Загрузка правила проверки {RULE_INDEX}")
    rule = load_rule(RULE_INDEX)
    if args.verbose:
        log_step(log_file, 1, "Загрузка правила проверки", f"Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}")

    print(f"[2/6] Загрузка данных из {args.parser_outputs}")
    data = load_data(args.parser_outputs)
    if args.verbose:
        log_step(log_file, 2, "Загрузка данных", f"Source files loaded")

    print("[3/6] Извлечение релевантных данных")
    extracted = extract_relevant_data(data)
    if args.verbose:
        log_step(log_file, 3, "Извлечение релевантных данных", f"Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}")

    print("[4/6] Формирование промпта с правилом из ТЗ")
    prompt = build_prompt(extracted, rule)
    if args.verbose:
        log_step(log_file, 4, "Сформированный промпт для LLM", f"FULL PROMPT:\n{prompt}")

    print("[5/6] Вызов LLM (gpt-4.1-mini-2025-04-14)")
    result = call_llm(prompt, api_key)
    if args.verbose:
        log_step(log_file, 5, "Ответ от LLM", f"LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}")

    print("[6/6] Сохранение результата")
    save_result(result, args.output)

    # Вывод результата
    status_emoji = "✅" if result["status"] == "PASS" else "❌"
    print(f"\n{status_emoji} Проверка {RULE_INDEX}: {result['status']}")
    if result["discrepancy"]:
        print(f"   Проблема: {result['discrepancy']}")
    if args.verbose:
        log_step(log_file, 6, "Финальный результат", f"Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}")
        print(f"\n📄 Лог сохранен: {log_file}")

if __name__ == "__main__":
    main()
