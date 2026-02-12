#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Валидатор 2: Вид организации и название предприятия.

Проверяет наличие юридической формы (ООО, ЗАО, АО, ПАО, ОАО, ИП)
и непустого наименования предприятия в ячейке B2.
Без LLM.
"""

import json
import os
import re
import argparse
from pathlib import Path

RULE_INDEX = "2"
RULE_TITLE = "Вид организации и название предприятия"


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
    """Проверка юридической формы и названия предприятия."""
    org_name = data.get("header", {}).get("org_name", "")
    errors = []

    if not org_name:
        errors.append("Поле org_name (B2) пустое — нет названия организации")
    else:
        # Проверка юридической формы
        jur_form_match = re.search(r'\b(ООО|ЗАО|АО|ПАО|ОАО|ИП)\b', org_name)
        if not jur_form_match:
            errors.append(
                f"Юридическая форма (ООО/ЗАО/АО/ПАО/ОАО/ИП) не найдена в '{org_name}'"
            )

        # Проверка наименования — должно быть что-то после юр. формы
        name_match = re.search(r'["\«](.+?)["\»]', org_name)
        if not name_match or not name_match.group(1).strip():
            # Попробуем без кавычек — просто текст после юр. формы
            after_form = re.sub(r'^(ООО|ЗАО|АО|ПАО|ОАО|ИП)\s*', '', org_name).strip()
            after_form = after_form.strip('"«»\'" ')
            if not after_form:
                errors.append("Наименование предприятия пустое после юридической формы")

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
