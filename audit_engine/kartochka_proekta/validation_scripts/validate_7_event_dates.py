#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Валидатор 7: Даты мероприятий.

Проверяет наличие дат начала и окончания мероприятий (секция 4).
Каждое событие должно иметь start_date.
Этапные события (начинающиеся с цифры и точки) должны иметь и end_date.
Без LLM.
"""

import json
import os
import re
import argparse
from pathlib import Path

RULE_INDEX = "7"
RULE_TITLE = "Даты мероприятий"


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


def _is_stage_event(name: str) -> bool:
    """Проверяет, является ли событие этапным (начинается с '1.', '2.', '3.', '4.')."""
    return bool(re.match(r'^\d+\.', name.strip()))


def validate(data: dict) -> dict:
    """Проверка дат мероприятий."""
    events = data.get("events", [])
    errors = []

    if not events:
        return {
            "rule_index": RULE_INDEX,
            "rule_title": RULE_TITLE,
            "status": "FAIL",
            "discrepancy": "Список мероприятий (events) пуст",
        }

    for i, event in enumerate(events):
        name = event.get("name", f"Событие #{i+1}")
        start_date = event.get("start_date")
        end_date = event.get("end_date")

        # Каждое событие должно иметь start_date
        if not start_date:
            errors.append(f"'{name[:40]}' — нет даты начала")

        # Этапные события должны иметь end_date
        if _is_stage_event(name) and not end_date:
            errors.append(f"'{name[:40]}' — этапное событие без даты окончания")

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
