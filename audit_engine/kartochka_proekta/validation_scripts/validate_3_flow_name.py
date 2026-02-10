#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Валидатор 3: Название потока.

Двухфазная проверка:
  Фаза 1 (non-LLM): поле не пустое и длина > 5
  Фаза 2 (LLM): текст является осмысленным названием проекта/потока,
                 а не placeholder или набор символов
"""

import json
import os
import argparse
from pathlib import Path

from openai import OpenAI

RULE_INDEX = "3"
RULE_TITLE = "Название потока"


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


def check_semantic(project_name: str, rule: dict, api_key: str) -> dict:
    """LLM-проверка осмысленности названия проекта/потока."""
    client = OpenAI(api_key=api_key)

    prompt = f"""Ты эксперт по проверке документов «Карточка проекта» в рамках бережливого производства.

ТРЕБОВАНИЕ:
{rule['requirement_expert']}

КРИТЕРИИ:
- Что проверять: {rule['validation_criteria']['what_to_check']}
- Условие успеха: {rule['validation_criteria']['success_condition']}
- Условие ошибки: {rule['validation_criteria']['error_condition']}

ФАКТИЧЕСКИЕ ДАННЫЕ:
Название проекта/потока: "{project_name}"

ЗАДАНИЕ:
Проверь, является ли название проекта осмысленным текстом, описывающим реальный проект или поток.
Не является осмысленным: placeholder ("Название проекта"), набор символов ("ааааа"), слишком общий текст ("тест").
Является осмысленным: конкретное описание проекта ("Оптимизация производства приборов учёта").

ФОРМАТ ОТВЕТА (строго JSON):
{{
  "rule_index": "{RULE_INDEX}",
  "rule_title": "{RULE_TITLE}",
  "status": "PASS или FAIL",
  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"
}}
"""

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": "Ты эксперт по валидации документов. Отвечаешь строго в формате JSON."},
            {"role": "user", "content": prompt}
        ],
        temperature=0,
        response_format={"type": "json_object"},
    )

    return json.loads(response.choices[0].message.content)


def validate(data: dict, rule: dict, api_key: str) -> dict:
    """Проверка названия проекта/потока: non-LLM + LLM."""
    project_name = data.get("header", {}).get("project_name", "")

    # Фаза 1: детерминированная проверка
    if not project_name:
        return {
            "rule_index": RULE_INDEX,
            "rule_title": RULE_TITLE,
            "status": "FAIL",
            "discrepancy": "Поле project_name (B4) пустое — нет названия проекта/потока",
        }

    if len(project_name) <= 5:
        return {
            "rule_index": RULE_INDEX,
            "rule_title": RULE_TITLE,
            "status": "FAIL",
            "discrepancy": f"Название проекта слишком короткое ({len(project_name)} симв.): '{project_name}'",
        }

    # Фаза 2: LLM-проверка осмысленности
    return check_semantic(project_name, rule, api_key)


def main():
    parser = argparse.ArgumentParser(description=f"Валидатор {RULE_INDEX}: {RULE_TITLE}")
    parser.add_argument("--parser-outputs", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("OPENAI_API_KEY не найден в переменных окружения")

    rule = load_rule(RULE_INDEX)
    data = load_data(args.parser_outputs)
    result = validate(data, rule, api_key)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    status_emoji = "✅" if result["status"] == "PASS" else "❌"
    print(f"{status_emoji} [{RULE_INDEX}] {RULE_TITLE}: {result['status']}")


if __name__ == "__main__":
    main()
