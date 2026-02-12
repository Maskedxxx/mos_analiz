#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Валидатор 5: Обязательные поля секции 1.

Проверяет 6 полей секции 1 (клиенты, периметр, владелец, границы, руководитель, команда).
Условие 1: все поля не пустые.
Условие 2: owner, leader, team содержат паттерн 'ФИО - должность'.
Без LLM.
"""

import json
import os
import re
import argparse
from pathlib import Path

RULE_INDEX = "5"
RULE_TITLE = "Обязательные поля секции 1"

# Поля секции 1 и их русские названия
REQUIRED_FIELDS = {
    "clients": "Клиенты процесса",
    "perimeter": "Периметр проекта",
    "owner": "Владелец процесса",
    "boundaries": "Границы процесса",
    "leader": "Руководитель проекта",
    "team": "Команда проекта",
}

# Поля, требующие формат "ФИО - должность"
FIO_FIELDS = ["owner", "leader", "team"]


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


def _check_fio_format(text: str) -> bool:
    """Проверяет наличие паттерна 'ФИО - должность' (разделитель — тире)."""
    return bool(re.search(r'.+\s*[-–—]\s*.+', text))


def validate(data: dict) -> dict:
    """Проверка обязательных полей секции 1."""
    section1 = data.get("section1", {})
    errors = []

    # Условие 1: все 6 полей не пустые
    empty_fields = []
    for field_key, field_name in REQUIRED_FIELDS.items():
        val = section1.get(field_key, "")
        if not val or not val.strip():
            empty_fields.append(field_name)

    if empty_fields:
        errors.append(f"Пустые поля: {', '.join(empty_fields)}")

    # Условие 2: формат ФИО-должность в owner, leader, team
    bad_format = []
    for field_key in FIO_FIELDS:
        val = section1.get(field_key, "")
        if not val:
            continue  # Уже отмечено выше как пустое
        # Для team — проверяем каждого участника (разделитель — запятая)
        if field_key == "team":
            members = [m.strip() for m in val.split(",") if m.strip()]
            for member in members:
                if not _check_fio_format(member):
                    bad_format.append(f"team: '{member[:50]}'")
                    break  # Достаточно одного примера
        else:
            if not _check_fio_format(val):
                bad_format.append(f"{REQUIRED_FIELDS[field_key]}: '{val[:50]}'")

    if bad_format:
        errors.append(f"Неверный формат ФИО-должность: {'; '.join(bad_format)}")

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
