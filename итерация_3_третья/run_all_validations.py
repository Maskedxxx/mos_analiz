#!/usr/bin/env python3
"""
Мастер-скрипт для запуска всех валидаторов параллельно и создания итогового отчета
"""
import os
import json
import argparse
import subprocess
from pathlib import Path
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
from typing import Dict, List, Tuple

def load_rules(rules_file: Path) -> List[Dict]:
    """Загрузка всех правил из validation_rules.json"""
    with open(rules_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data["rules"]

def get_validator_script(rule_index: str, scripts_dir: Path) -> Path:
    """Определить путь к валидатору по номеру правила"""
    # Преобразуем rule_index в имя файла: "1.1" -> "validate_1_1_*.py"
    rule_prefix = rule_index.replace(".", "_")

    # Ищем файл валидатора
    pattern = f"validate_{rule_prefix}_*.py"
    matches = list(scripts_dir.glob(pattern))

    if not matches:
        raise FileNotFoundError(f"Валидатор для правила {rule_index} не найден (паттерн: {pattern})")

    if len(matches) > 1:
        raise ValueError(f"Найдено несколько валидаторов для правила {rule_index}: {matches}")

    return matches[0]

def run_validator(
    rule_index: str,
    validator_script: Path,
    parser_outputs: Path,
    output_file: Path,
    verbose: bool = False
) -> Tuple[str, Dict, float]:
    """
    Запуск одного валидатора
    Возвращает: (rule_index, result_data, duration)
    """
    start_time = datetime.now()

    # Формируем команду
    cmd = [
        "python",
        str(validator_script),
        "--parser-outputs", str(parser_outputs),
        "--output", str(output_file)
    ]

    if verbose:
        cmd.append("--verbose")

    try:
        # Запускаем валидатор
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300  # 5 минут таймаут
        )

        # Вычисляем длительность
        duration = (datetime.now() - start_time).total_seconds()

        # Читаем результат
        if output_file.exists():
            with open(output_file, "r", encoding="utf-8") as f:
                result_data = json.load(f)
        else:
            result_data = {
                "rule_index": rule_index,
                "status": "ERROR",
                "discrepancy": f"Output file not created. STDERR: {result.stderr[:200]}"
            }

        return rule_index, result_data, duration

    except subprocess.TimeoutExpired:
        duration = (datetime.now() - start_time).total_seconds()
        return rule_index, {
            "rule_index": rule_index,
            "status": "TIMEOUT",
            "discrepancy": "Validator exceeded 5 minute timeout"
        }, duration

    except Exception as e:
        duration = (datetime.now() - start_time).total_seconds()
        return rule_index, {
            "rule_index": rule_index,
            "status": "ERROR",
            "discrepancy": f"Exception: {str(e)}"
        }, duration

def run_all_validators_parallel(
    rules: List[Dict],
    scripts_dir: Path,
    parser_outputs: Path,
    output_dir: Path,
    max_workers: int = 5,
    verbose: bool = False
) -> List[Tuple[Dict, Dict, float]]:
    """
    Запуск всех валидаторов параллельно
    Возвращает список: [(rule, result_data, duration), ...]
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    results = []

    # Используем ThreadPoolExecutor для параллельного запуска
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        # Подготавливаем задачи
        futures = {}

        for rule in rules:
            rule_index = rule["rule_index"]

            try:
                validator_script = get_validator_script(rule_index, scripts_dir)
                output_file = output_dir / f"validate_{rule_index.replace('.', '_')}.json"

                future = executor.submit(
                    run_validator,
                    rule_index,
                    validator_script,
                    parser_outputs,
                    output_file,
                    verbose
                )
                futures[future] = rule

            except FileNotFoundError as e:
                # Валидатор не найден - записываем ошибку
                results.append((rule, {
                    "rule_index": rule_index,
                    "status": "MISSING",
                    "discrepancy": str(e)
                }, 0.0))

        # Собираем результаты по мере выполнения
        for future in as_completed(futures):
            rule = futures[future]
            try:
                rule_index, result_data, duration = future.result()
                results.append((rule, result_data, duration))

                # Прогресс
                status_emoji = "✅" if result_data.get("status") == "PASS" else "❌"
                print(f"{status_emoji} [{len(results)}/{len(rules)}] {rule_index}: {result_data.get('status')} ({duration:.1f}s)")

            except Exception as e:
                results.append((rule, {
                    "rule_index": rule["rule_index"],
                    "status": "ERROR",
                    "discrepancy": f"Future exception: {str(e)}"
                }, 0.0))

    return results

def create_excel_report(
    results: List[Tuple[Dict, Dict, float]],
    report_file: Path
):
    """Создание Excel-отчета с результатами"""

    # Подготавливаем данные для таблицы
    rows = []

    for rule, result_data, duration in results:
        row = {
            "rule_index": rule["rule_index"],
            "section": rule.get("section", ""),
            "rule_title": rule.get("rule_title", ""),
            "status": result_data.get("status", "UNKNOWN"),
            "discrepancy": result_data.get("discrepancy", ""),
            "duration_sec": round(duration, 2)
        }
        rows.append(row)

    # Создаем DataFrame
    df = pd.DataFrame(rows)

    # Сортируем по rule_index
    df = df.sort_values("rule_index")

    # Сохраняем в Excel
    report_file.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Validation Results', index=False)

        # Автоширина колонок
        worksheet = writer.sheets['Validation Results']
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(col)
            )
            worksheet.column_dimensions[chr(65 + idx)].width = min(max_length + 2, 50)

    print(f"\n📊 Отчет сохранен: {report_file}")

    # Статистика
    total = len(df)
    passed = len(df[df["status"] == "PASS"])
    failed = len(df[df["status"] == "FAIL"])
    errors = len(df[~df["status"].isin(["PASS", "FAIL"])])

    print("\n📈 Статистика:")
    print(f"   Всего правил: {total}")
    print(f"   ✅ PASS: {passed} ({passed/total*100:.1f}%)")
    print(f"   ❌ FAIL: {failed} ({failed/total*100:.1f}%)")
    if errors > 0:
        print(f"   ⚠️  ERRORS: {errors} ({errors/total*100:.1f}%)")

    total_time = df["duration_sec"].sum()
    print(f"\n⏱  Общее время: {total_time:.1f}s")

def main():
    parser = argparse.ArgumentParser(
        description="Запуск всех валидаторов и создание итогового отчета"
    )
    parser.add_argument(
        "--parser-outputs",
        type=Path,
        default=Path("parser_outputs"),
        help="Путь к папке с результатами парсинга"
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("validation_outputs"),
        help="Папка для сохранения результатов валидации"
    )
    parser.add_argument(
        "--report",
        type=Path,
        default=Path("validation_report.xlsx"),
        help="Путь к файлу итогового отчета (Excel)"
    )
    parser.add_argument(
        "--rules",
        type=Path,
        default=Path("validation_rules.json"),
        help="Путь к файлу с правилами валидации"
    )
    parser.add_argument(
        "--scripts-dir",
        type=Path,
        default=Path("validation_scripts"),
        help="Папка с валидаторами"
    )
    parser.add_argument(
        "--parallel",
        type=int,
        default=5,
        help="Количество параллельных процессов"
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Включить детальное логирование валидаторов"
    )

    args = parser.parse_args()

    # Проверка наличия API ключа
    if not os.environ.get("OPENAI_API_KEY"):
        print("❌ Ошибка: не найден OPENAI_API_KEY в переменных окружения")
        return 1

    print("🚀 Запуск всех валидаторов...")
    print(f"   Правила: {args.rules}")
    print(f"   Валидаторы: {args.scripts_dir}")
    print(f"   Параллельность: {args.parallel}")
    print(f"   Verbose: {args.verbose}")
    print()

    # Загружаем правила
    rules = load_rules(args.rules)
    print(f"📋 Загружено правил: {len(rules)}")
    print()

    # Запускаем валидаторы
    start_time = datetime.now()
    results = run_all_validators_parallel(
        rules,
        args.scripts_dir,
        args.parser_outputs,
        args.output_dir,
        max_workers=args.parallel,
        verbose=args.verbose
    )
    total_duration = (datetime.now() - start_time).total_seconds()

    print(f"\n✅ Все валидаторы завершены за {total_duration:.1f}s")

    # Создаем отчет
    create_excel_report(results, args.report)

    return 0

if __name__ == "__main__":
    exit(main())
