#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Валидатор 11: Способ расчёта и источник данных.

Проверяет наличие способа расчёта и источника данных
для каждого показателя в методике расчёта.
Без LLM.
"""

import json
import os
import argparse
from pathlib import Path

RULE_INDEX = "11"
RULE_TITLE = "Способ расчёта и источник данных"


def load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get(
        "VALIDATION_RULES_PATH",
        str(Path(__file__).resolve().parents[2] / "doc_configs" / "kartochka_proekta" / "validation_rules.json")
    ))
    with open(rules_file, "r", encoding="utf-8") as f:
        data = json.load(f)
    for rule in data["rules"]:
        if rule["rule_index"] == rule_index:
            return rule
    raise ValueError(f"Правило {rule_index} не найдено")


def load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка metodika.json."""
    with open(parser_outputs_dir / "metodika.json", "r", encoding="utf-8") as f:
        return json.load(f)


def validate(data: dict) -> dict:
    """Проверка способа расчёта и источника данных."""
    indicators = data.get("indicators", [])
    errors = []

    for ind in indicators:
        name = ind.get("name", "")
        unit = ind.get("unit")
        calc_method = ind.get("calc_method")
        data_source = ind.get("data_source")

        # Проверяем только показатели с заполненной единицей измерения
        if not unit:
            continue

        missing = []
        if not calc_method or not str(calc_method).strip():
            missing.append("способ расчёта")
        if not data_source or not str(data_source).strip():
            missing.append("источник данных")

        if missing:
            errors.append(
                f"Показатель '{name}': не заполнено — {', '.join(missing)}"
            )

    if errors:
        return {
            "rule_index": RULE_INDEX,
            "rule_title": RULE_TITLE,
            "status": "FAIL",
            "discrepancy": "; ".join(errors),
        }

    return {
        "rule_index": RULE_INDEX,
        "rule_title": RULE_TITLE,
        "status": "PASS",
        "discrepancy": "",
    }


def main():
    parser = argparse.ArgumentParser(description=f"Валидатор {RULE_INDEX}: {RULE_TITLE}")
    parser.add_argument("--parser-outputs", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    load_rule(RULE_INDEX)
    data = load_data(args.parser_outputs)
    result = validate(data)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    status_emoji = "✅" if result["status"] == "PASS" else "❌"
    print(f"{status_emoji} [{RULE_INDEX}] {RULE_TITLE}: {result['status']}")


if __name__ == "__main__":
    main()
