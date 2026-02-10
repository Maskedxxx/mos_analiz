#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Валидатор 10: Единицы измерения (методика расчета).

Кросс-проверка между листами: единицы измерения из методики расчёта
сверяются с единицами в карточке проекта.
Без LLM.
"""

import json
import os
import argparse
from pathlib import Path

RULE_INDEX = "10"
RULE_TITLE = "Единицы измерения (методика расчета)"


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
    """Загрузка kartochka_main.json и metodika.json."""
    with open(parser_outputs_dir / "kartochka_main.json", "r", encoding="utf-8") as f:
        kartochka = json.load(f)
    with open(parser_outputs_dir / "metodika.json", "r", encoding="utf-8") as f:
        metodika = json.load(f)
    return kartochka, metodika


def _normalize_name(name: str) -> str:
    """Нормализация названия показателя для сопоставления."""
    # Убираем суффиксы типа '(при наличии):', двоеточия, лишние пробелы
    import re
    name = re.sub(r'\(.*?\)', '', name)
    name = name.replace(':', '').strip().lower()
    return name


def validate(kartochka: dict, metodika: dict) -> dict:
    """Проверка совпадения единиц измерения между методикой и карточкой."""
    k_indicators = kartochka.get("indicators", [])
    m_indicators = metodika.get("indicators", [])
    errors = []

    # Строим карту карточки: normalized_name → unit
    k_units = {}
    for ind in k_indicators:
        name = ind.get("name", "")
        if name:
            k_units[_normalize_name(name)] = ind.get("unit", "")

    # Проверяем каждый показатель из методики
    for m_ind in m_indicators:
        m_name = m_ind.get("name", "")
        m_unit = m_ind.get("unit")

        if not m_unit:
            # Единица в методике пустая — ошибка если в карточке есть такой показатель
            m_norm = _normalize_name(m_name)
            if m_norm in k_units and k_units[m_norm]:
                errors.append(
                    f"Показатель '{m_name}': единица в методике не заполнена, "
                    f"в карточке — '{k_units[m_norm]}'"
                )
            continue

        # Ищем соответствие в карточке
        m_norm = _normalize_name(m_name)
        k_unit = k_units.get(m_norm)

        if k_unit is None:
            # Показатель из методики не найден в карточке — пропускаем
            continue

        # Сравниваем единицы (case-insensitive, strip)
        if m_unit.lower().strip() != k_unit.lower().strip():
            errors.append(
                f"Показатель '{m_name}': единица в методике '{m_unit}' ≠ "
                f"единица в карточке '{k_unit}'"
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
    kartochka, metodika = load_data(args.parser_outputs)
    result = validate(kartochka, metodika)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    status_emoji = "✅" if result["status"] == "PASS" else "❌"
    print(f"{status_emoji} [{RULE_INDEX}] {RULE_TITLE}: {result['status']}")


if __name__ == "__main__":
    main()
