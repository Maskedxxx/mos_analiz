#!/usr/bin/env python3
"""
Валидатор 4.1: Все ячейки строки ВПП заполнены
Источник ТЗ: ТЗ_Валидация_КПСЦ_v2_упрощенная.md, Раздел 4, Проверка 4.1
"""
import os
import json
import argparse
from pathlib import Path
from openai import OpenAI
from datetime import datetime

# Константы
RULE_INDEX = "4.1"
RULE_TITLE = "Все ячейки строки ВПП заполнены"
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
    table1_file = parser_outputs_dir / "kpsc_table1_v2.json"

    if not table1_file.exists():
        return {"file_exists": False, "data": None}

    with open(table1_file, "r", encoding="utf-8") as f:
        return {"file_exists": True, "data": json.load(f)}

def extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data["file_exists"]:
        return {
            "file_exists": False,
            "empty_cells": []
        }

    rows = data["data"].get("rows", [])

    # Находим строку с "ВПП"
    vpp_row_idx = None
    header_row_idx = None

    for i, row in enumerate(rows):
        cells = row.get("cells", [])
        # Ищем "ВПП" в колонках B или C (col 2-3)
        for cell in cells:
            if cell.get("col") in [2, 3]:
                cell_value = str(cell.get("value", "")).upper()
                if "ВПП" in cell_value:
                    vpp_row_idx = i
                    # Ищем заголовок "Показатель" - обычно несколько строк выше
                    for j in range(max(0, i-10), i):
                        row_cells = rows[j].get("cells", [])
                        for rc in row_cells:
                            if rc.get("col") in [2, 3]:
                                if str(rc.get("value", "")).lower().strip() == "показатель":
                                    header_row_idx = j
                                    break
                        if header_row_idx is not None:
                            break
                    break
        if vpp_row_idx is not None:
            break

    if vpp_row_idx is None:
        return {
            "file_exists": True,
            "empty_cells": [],
            "error": "Не найдена строка ВПП"
        }

    # Извлекаем заголовки операций и значения ВПП
    # Используем словари для поиска по номеру колонки
    vpp_cells = rows[vpp_row_idx].get("cells", [])
    vpp_dict = {cell.get("col"): cell.get("value") for cell in vpp_cells}

    header_cells = rows[header_row_idx].get("cells", []) if header_row_idx is not None else []
    header_dict = {cell.get("col"): cell.get("value", "") for cell in header_cells}

    empty_cells = []

    # Проверяем все колонки начиная с E (col=5) до последней колонки в данных
    all_cols = set(vpp_dict.keys()) | set(header_dict.keys())

    for col_num in sorted(all_cols):
        if col_num < 5:  # Пропускаем колонки A-D
            continue

        # Получаем название операции из заголовка
        raw_header = header_dict.get(col_num)
        header_val = str(raw_header).strip() if raw_header is not None else ""

        # Пропускаем колонки без заголовка (пустые/служебные колонки за пределами данных)
        if not header_val:
            continue

        # Пропускаем колонку "ИТОГО"
        if "итого" in header_val.lower():
            continue

        # Проверяем, заполнена ли ячейка ВПП
        vpp_value = vpp_dict.get(col_num)
        if vpp_value is None or str(vpp_value).strip() == "":
            empty_cells.append({
                "column": col_num,
                "operation_name": header_val if header_val else f"Колонка {col_num}"
            })

    return {
        "file_exists": True,
        "empty_cells": empty_cells,
        "has_empty": len(empty_cells) > 0,
        "empty_count": len(empty_cells)
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
Пустые ячейки в строке ВПП:
{json.dumps(extracted_data.get('empty_cells', []), ensure_ascii=False, indent=2)}

Есть пустые ячейки: {extracted_data.get('has_empty', False)}
Количество пустых: {extracted_data.get('empty_count', 0)}

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
    base_url = os.environ.get("LLM_BASE_URL", "http://localhost:8001/v1/")
    client = OpenAI(api_key=api_key, base_url=base_url)

    response = client.chat.completions.create(
        model=os.environ.get("LLM_MODEL", "openai/gpt-oss-120b"),
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
    api_key = os.environ.get("OPENAI_API_KEY", "dummy")


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
        log_step(log_file, 2, "Загрузка данных", f"Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}")

    print("[3/6] Извлечение релевантных данных")
    extracted = extract_relevant_data(data)
    if args.verbose:
        log_step(log_file, 3, "Извлечение релевантных данных", f"Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}")

    print("[4/6] Формирование промпта с правилом из ТЗ")
    prompt = build_prompt(extracted, rule)
    if args.verbose:
        log_step(log_file, 4, "Сформированный промпт для LLM", f"FULL PROMPT:\n{prompt}")

    print("[5/6] Вызов LLM")
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
