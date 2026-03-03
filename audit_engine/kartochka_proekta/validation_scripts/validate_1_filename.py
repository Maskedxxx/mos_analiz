#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Валидатор 1: Название файла.

Проверяет что имя XLSX-файла содержит ключевые слова «Карточка» и «проект».
Наименование предприятия НЕ проверяется (динамические данные).
Без LLM — детерминированная проверка.
"""

import json
import os
import argparse
from pathlib import Path

RULE_INDEX = "1"
RULE_TITLE = "Название файла"


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


def validate(data: dict, rule: dict) -> dict:
    """Проверка названия файла.

    Проверяет наличие ключевых слов «Карточка» и «проект» в имени файла.
    Наименование предприятия НЕ проверяется — оно динамическое.
    """
    filename = data.get("meta", {}).get("workbook", "")
    filename_lower = filename.lower()

    keywords = ["карточка", "проект"]
    missing = [kw for kw in keywords if kw not in filename_lower]

    if missing:
        return {
            "rule_index": RULE_INDEX,
            "rule_title": RULE_TITLE,
            "status": "FAIL",
            "discrepancy": f"Имя файла '{filename}' не содержит ключевые слова: {missing}",
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

    rule = load_rule(RULE_INDEX)
    data = load_data(args.parser_outputs)
    result = validate(data, rule)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    status_emoji = "✅" if result["status"] == "PASS" else "❌"
    print(f"{status_emoji} [{RULE_INDEX}] {RULE_TITLE}: {result['status']}")


if __name__ == "__main__":
    main()
