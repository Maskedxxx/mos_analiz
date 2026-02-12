#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Валидатор 4: Должность, ФИО подписанта, дата и подпись.

Проверяет наличие и заполненность должности (K4), ФИО (K7), даты (K8).
Без LLM.
"""

import json
import os
import re
import argparse
from pathlib import Path

RULE_INDEX = "4"
RULE_TITLE = "Должность, ФИО подписанта, дата и подпись"


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
    """Проверка подписанта."""
    header = data.get("header", {})
    errors = []

    # Должность
    position = header.get("signee_position", "")
    if not position:
        errors.append("Должность подписанта (K4) не заполнена")

    # ФИО — хотя бы 2 буквенных слова
    name = header.get("signee_name", "")
    if not name:
        errors.append("ФИО подписанта (K7) не заполнено")
    else:
        words = re.findall(r'[А-Яа-яЁёA-Za-z]+\.?', name)
        if len(words) < 2:
            errors.append(f"ФИО подписанта содержит менее 2 слов: '{name}'")

    # Дата — не должна содержать placeholder '___'
    date_val = header.get("signee_date", "")
    if not date_val:
        errors.append("Дата подписания (K8) не заполнена")
    elif "___" in date_val or "20__" in date_val:
        errors.append(
            f"Дата подписания содержит незаполненные placeholder: '{date_val}'"
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
