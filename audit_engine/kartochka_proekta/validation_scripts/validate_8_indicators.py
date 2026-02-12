#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Валидатор 8: Показатели и даты.

Проверяет наличие показателей (name, unit, base_value, target_value, ideal_value)
и дат периодов (base_date, target_date, ideal_date).
Без LLM.
"""

import json
import os
import argparse
from pathlib import Path

RULE_INDEX = "8"
RULE_TITLE = "Показатели и даты"


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
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / "kartochka_main.json", "r", encoding="utf-8") as f:
        return json.load(f)


def validate(data: dict) -> dict:
    """Проверка показателей и дат."""
    indicators = data.get("indicators", [])
    indicator_dates = data.get("indicator_dates", {})
    errors = []

    # Минимум 1 показатель
    if not indicators:
        errors.append("Нет показателей (массив indicators пуст)")
    else:
        # Проверяем каждый показатель
        required_fields = ["name", "unit", "base_value", "target_value", "ideal_value"]
        for ind in indicators:
            num = ind.get("number", "?")
            missing = []
            for field in required_fields:
                val = ind.get(field)
                if val is None or (isinstance(val, str) and not val.strip()):
                    missing.append(field)
            if missing:
                errors.append(
                    f"Показатель #{num}: пустые поля [{', '.join(missing)}]"
                )

    # Проверка дат периодов
    date_fields = {"base_date": "Дата базы", "target_date": "Дата цели", "ideal_date": "Дата идеала"}
    for field, label in date_fields.items():
        val = indicator_dates.get(field)
        if not val:
            errors.append(f"{label} не заполнена")

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
