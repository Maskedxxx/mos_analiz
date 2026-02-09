#!/usr/bin/env python3
"""
Валидатор 4.2: Сумма значений строки ВПП совпадает с итоговым значением
Источник ТЗ: ТЗ_Валидация_КПСЦ_v2_упрощенная.md, Раздел 4, Проверка 4.2
"""
import os
import json
import argparse
from pathlib import Path
from openai import OpenAI
from datetime import datetime

# Константы
RULE_INDEX = "4.2"
RULE_TITLE = "Сумма значений строки ВПП совпадает с итоговым значением"
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
            "vpp_values": [],
            "calculated_sum": 0,
            "itogo_value": 0
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
            "error": "Не найдена строка ВПП"
        }

    # Извлекаем значения ВПП - используем словари для поиска по номеру колонки
    # ВАЖНО: учитываем только merge_anchor=True или merge_range=null (чтобы не считать merge-ячейки несколько раз)
    vpp_cells = rows[vpp_row_idx].get("cells", [])
    vpp_dict = {}
    vpp_cells_full = {}  # Сохраняем полную информацию о ячейках

    for cell in vpp_cells:
        col = cell.get("col")
        vpp_dict[col] = cell.get("value")
        vpp_cells_full[col] = cell

    header_cells = rows[header_row_idx].get("cells", []) if header_row_idx is not None else []
    header_dict = {cell.get("col"): str(cell.get("value", "")).strip() for cell in header_cells}

    vpp_values = []
    itogo_value = None

    # Проверяем все колонки начиная с E (col=5)
    all_cols = set(vpp_dict.keys()) | set(header_dict.keys())

    for col_num in sorted(all_cols):
        if col_num < 5:  # Пропускаем колонки A-D
            continue

        header_val = header_dict.get(col_num, "")

        # Проверяем merge-ячейки: берем только merge_anchor=True или не-merge ячейки
        if col_num in vpp_cells_full:
            cell_info = vpp_cells_full[col_num]
            is_merge_anchor = cell_info.get("merge_anchor", False)
            merge_range = cell_info.get("merge_range")

            # Пропускаем не-anchor ячейки из merge-диапазона
            if merge_range is not None and not is_merge_anchor:
                continue

        cell_value = vpp_dict.get(col_num)

        # Если это колонка ИТОГО
        if "итого" in header_val.lower():
            try:
                itogo_value = float(str(cell_value).replace(",", "."))
            except:
                itogo_value = None
            break

        # Пытаемся преобразовать в число
        if cell_value is not None and str(cell_value).strip():
            try:
                num_value = float(str(cell_value).replace(",", "."))
                vpp_values.append({
                    "column": col_num,
                    "operation": header_val if header_val else f"Колонка {col_num}",
                    "value": num_value
                })
            except:
                pass  # Пропускаем нечисловые значения

    # Вычисляем сумму
    calculated_sum = sum(v["value"] for v in vpp_values)

    # Проверяем разницу (допуск ±0.01)
    difference = None
    if itogo_value is not None:
        difference = abs(calculated_sum - itogo_value)

    return {
        "file_exists": True,
        "vpp_values": vpp_values,
        "calculated_sum": calculated_sum,
        "itogo_value": itogo_value,
        "difference": difference,
        "matches": difference <= 0.01 if difference is not None else False
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
Значения ВПП по операциям:
{json.dumps(extracted_data.get('vpp_values', []), ensure_ascii=False, indent=2)}

Вычисленная сумма: {extracted_data.get('calculated_sum', 0)}
Значение ИТОГО: {extracted_data.get('itogo_value', 0)}
Расхождение: {extracted_data.get('difference', 0)}
Суммы совпадают (±0.01): {extracted_data.get('matches', False)}

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
        log_step(log_file, 2, "Загрузка данных", f"Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}")

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
