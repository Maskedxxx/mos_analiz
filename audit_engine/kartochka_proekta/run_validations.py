#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Оркестратор валидации «Карточка проекта».

Запускает 3 парсера и 11 валидаторов, собирает результаты.
Паттерн аналогичен audit_engine/kpsc/run_validations.py.
"""

import importlib
import json
import os
import subprocess
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd


# === Реестр парсеров ===
PARSER_MODULES = [
    "audit_engine.kartochka_proekta.parser_scripts.parse_kartochka_main",
    "audit_engine.kartochka_proekta.parser_scripts.parse_metodika",
    "audit_engine.kartochka_proekta.parser_scripts.parse_dropdown",
]


def load_rules(rules_path: Path) -> List[Dict]:
    """Загрузка правил из validation_rules.json."""
    with open(rules_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data["rules"]


def run_parsers(xlsx_path: Path, output_dir: Path) -> Dict[str, Any]:
    """
    Запускает все 3 парсера.

    Каждый парсер: import → parse(xlsx_path, output_dir).
    Ошибки отдельных парсеров не останавливают остальные.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    results = {}

    for module_name in PARSER_MODULES:
        short_name = module_name.rsplit(".", 1)[-1]
        try:
            mod = importlib.import_module(module_name)
            mod.parse(xlsx_path, output_dir)
            results[short_name] = {"status": "ok"}
            print(f"  [OK] {short_name}")
        except Exception as e:
            results[short_name] = {"status": "error", "error": str(e)}
            print(f"  [ERR] {short_name}: {e}")

    return results


def run_single_validator(
    rule: Dict,
    parser_outputs_dir: Path,
    output_dir: Path,
    rules_path: Path,
) -> Tuple[Dict, Dict, float]:
    """
    Запускает один валидатор как subprocess.

    Валидатор — скрипт validate_N_*.py в validation_scripts/.
    Принимает --parser-outputs и --output, пишет JSON-результат.
    """
    rule_index = rule["rule_index"]
    start = time.time()

    try:
        # Находим скрипт валидатора
        prefix = rule_index.replace(".", "_")
        scripts_dir = Path(__file__).parent / "validation_scripts"
        pattern = f"validate_{prefix}_*.py"
        matches = list(scripts_dir.glob(pattern))

        if not matches:
            return rule, {
                "rule_index": rule_index,
                "status": "MISSING",
                "discrepancy": f"Валидатор не найден (паттерн: {pattern})"
            }, 0.0

        script_path = matches[0]
        output_file = output_dir / f"validate_{prefix}.json"

        # Запускаем как subprocess — изоляция argparse
        env = os.environ.copy()
        env["VALIDATION_RULES_PATH"] = str(rules_path)

        result = subprocess.run(
            [sys.executable, str(script_path),
             "--parser-outputs", str(parser_outputs_dir),
             "--output", str(output_file)],
            capture_output=True,
            text=True,
            timeout=300,
            env=env,
        )

        duration = time.time() - start

        if output_file.exists():
            with open(output_file, "r", encoding="utf-8") as f:
                result_data = json.load(f)
        else:
            result_data = {
                "rule_index": rule_index,
                "status": "ERROR",
                "discrepancy": f"Выходной файл не создан. STDERR: {result.stderr[:500]}"
            }

        return rule, result_data, duration

    except Exception as e:
        duration = time.time() - start
        return rule, {
            "rule_index": rule_index,
            "status": "ERROR",
            "discrepancy": f"Исключение: {str(e)}"
        }, duration


def run_validators_parallel(
    rules: List[Dict],
    parser_outputs_dir: Path,
    output_dir: Path,
    rules_path: Path,
    max_workers: int = 5,
) -> List[Tuple[Dict, Dict, float]]:
    """Запуск всех валидаторов параллельно через ThreadPoolExecutor."""
    output_dir.mkdir(parents=True, exist_ok=True)
    results = []

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {}
        for rule in rules:
            future = executor.submit(
                run_single_validator,
                rule,
                parser_outputs_dir,
                output_dir,
                rules_path,
            )
            futures[future] = rule

        for future in as_completed(futures):
            rule = futures[future]
            try:
                rule_data, result_data, duration = future.result()
                results.append((rule_data, result_data, duration))

                status = result_data.get("status", "?")
                emoji = "+" if status == "PASS" else "-" if status == "FAIL" else "!"
                idx = result_data.get("rule_index", "?")
                print(f"  [{emoji}] [{len(results)}/{len(rules)}] {idx}: {status} ({duration:.1f}s)")

            except Exception as e:
                results.append((rule, {
                    "rule_index": rule["rule_index"],
                    "status": "ERROR",
                    "discrepancy": f"Future exception: {str(e)}"
                }, 0.0))

    return results


def create_excel_report(
    results: List[Tuple[Dict, Dict, float]],
    report_path: Path,
):
    """Создание Excel-отчёта с результатами валидации."""
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

        # Автоширина столбцов
        worksheet = writer.sheets["Validation Results"]
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(col)
            )
            from openpyxl.utils import get_column_letter
            worksheet.column_dimensions[get_column_letter(idx + 1)].width = min(max_length + 2, 50)

    return df
