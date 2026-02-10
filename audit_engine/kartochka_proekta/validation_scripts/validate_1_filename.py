#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Валидатор 1: Название файла.

Проверяет что имя XLSX-файла содержит '0.2 Карточка проекта'
и наименование предприятия.
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
    """Проверка названия файла."""
    filename = data.get("meta", {}).get("workbook", "")
    org_name = data.get("header", {}).get("org_name", "")

    # Мок: ожидаемый паттерн (Этап 2 — из БД через API)
    expected_prefix = "0.2 Карточка проекта"

    errors = []

    # Проверка 1: файл содержит ожидаемый префикс
    if expected_prefix.lower() not in filename.lower():
        errors.append(
            f"Имя файла '{filename}' не содержит ожидаемый префикс '{expected_prefix}'"
        )

    # Проверка 2: файл содержит наименование предприятия
    # Извлекаем название из org_name (убираем юр. форму и кавычки)
    if org_name:
        # "ООО \"ОБРАЗЕЦ\"" → "ОБРАЗЕЦ"
        import re
        name_match = re.search(r'["\«](.+?)["\»]', org_name)
        if name_match:
            enterprise_name = name_match.group(1).lower()
            if enterprise_name not in filename.lower():
                errors.append(
                    f"Имя файла не содержит наименование предприятия '{name_match.group(1)}'"
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
