#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Валидатор 9: Единицы измерения (карточка).

Кросс-проверка между листами: единицы измерения из показателей карточки
сверяются со справочником допустимых единиц из листа «выпадающий список».
Без LLM.
"""

import json
import os
import argparse
from pathlib import Path

RULE_INDEX = "9"
RULE_TITLE = "Единицы измерения (карточка)"

# Маппинг ключевых слов в названии показателя → категория справочника
CATEGORY_KEYWORDS = {
    "Время протекания процесса": ["время", "протекан"],
    "Выработка": ["выработк"],
    "Незавершенное производство": ["запас", "нзп", "незавершен"],
}
DEFAULT_CATEGORY = "Дополнительный показатель"


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


def load_data(parser_outputs_dir: Path) -> tuple:
    """Загрузка kartochka_main.json и dropdown_units.json."""
    with open(parser_outputs_dir / "kartochka_main.json", "r", encoding="utf-8") as f:
        kartochka = json.load(f)
    with open(parser_outputs_dir / "dropdown_units.json", "r", encoding="utf-8") as f:
        dropdown = json.load(f)
    return kartochka, dropdown


def _detect_category(indicator_name: str) -> str:
    """Определяет категорию показателя по ключевым словам в названии."""
    name_lower = indicator_name.lower()
    for category, keywords in CATEGORY_KEYWORDS.items():
        if any(kw in name_lower for kw in keywords):
            return category
    return DEFAULT_CATEGORY


def validate(kartochka: dict, dropdown: dict) -> dict:
    """Проверка единиц измерения показателей по справочнику."""
    indicators = kartochka.get("indicators", [])
    categories = dropdown.get("categories", {})
    errors = []

    # Собираем все допустимые единицы (все категории) для fallback
    all_units = set()
    for units_list in categories.values():
        for u in units_list:
            all_units.add(u.lower().strip())

    for ind in indicators:
        name = ind.get("name", "")
        unit = ind.get("unit", "")
        num = ind.get("number", "?")

        if not unit:
            errors.append(f"Показатель #{num} '{name}': единица измерения не заполнена")
            continue

        # Определяем категорию
        category = _detect_category(name) if name else DEFAULT_CATEGORY

        # Получаем допустимые единицы для категории
        allowed_units = categories.get(category, [])
        allowed_lower = [u.lower().strip() for u in allowed_units]

        unit_lower = unit.lower().strip()

        # Проверяем принадлежность
        if unit_lower not in allowed_lower:
            # Проверяем по всем категориям (мягкая проверка)
            if unit_lower not in all_units:
                errors.append(
                    f"Показатель #{num} '{name}': единица '{unit}' не найдена в справочнике "
                    f"(категория: {category}, допустимо: {', '.join(allowed_units[:5])}...)"
                )
            else:
                errors.append(
                    f"Показатель #{num} '{name}': единица '{unit}' не соответствует категории "
                    f"'{category}' (допустимо: {', '.join(allowed_units[:5])}...)"
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
    kartochka, dropdown = load_data(args.parser_outputs)
    result = validate(kartochka, dropdown)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    status_emoji = "✅" if result["status"] == "PASS" else "❌"
    print(f"{status_emoji} [{RULE_INDEX}] {RULE_TITLE}: {result['status']}")


if __name__ == "__main__":
    main()
