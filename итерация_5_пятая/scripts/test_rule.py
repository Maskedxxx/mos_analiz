#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для тестирования одного правила.
Парсит документы, формирует промпт, отправляет в LLM, выводит ответ.

Использование:
  python scripts/test_rule.py --rule 2
  python scripts/test_rule.py --rule 7 --use-errors
"""

import argparse
import json
import sys
from pathlib import Path

from openai import OpenAI

# Добавляем путь к модулям
sys.path.insert(0, str(Path(__file__).parent.parent))
from scripts.parser_prikaz_ic_docs import parse_prikaz_ic

# === КОНФИГУРАЦИЯ ===
PROJECT_DIR = Path(__file__).parent.parent
TARGET_DOC = PROJECT_DIR / "documents" / "prikaz_ic_ok.docx"
TARGET_WITH_ERRORS = PROJECT_DIR / "documents" / "prikaz_ic_errors.docx"
TEMPLATE_DOC = PROJECT_DIR / "templates" / "prikaz_ic_template.docx"
RULES_FILE = PROJECT_DIR / "config" / "tz_prikaz_ic.json"
DEFAULT_MODEL = "gpt-4.1-mini"

# === СИСТЕМНЫЙ ПРОМПТ ===
SYSTEM_PROMPT = """Ты — строгий аудитор документов.
Ты проверяешь ОДНО правило за один запрос.

Тебе даётся:
- фрагменты целевого документа (TARGET_*);
- фрагменты шаблона (TEMPLATE_*) — если требуется сверка;
- параметры правила: RULE_INDEX, COMPARE, SCOPE, RULE_TITLE;
- текстовые инструкции правила (RULE_INSTRUCTIONS).

COMPARE:
- template — сверяй чанк целевого с чанком шаблона (игнорируй плейсхолдеры: даты, ФИО, должности, названия компаний).
- target_only — используй только целевой документ.
- cross_check — сверяй два чанка целевого документа между собой.

Формат ответа — ОДИН JSON-объект:
- Если нарушений нет:
  {"status": "ok"}
- Если есть нарушения:
  {
    "status": "fail",
    "rule_index": <номер>,
    "rule_title": "<заголовок>",
    "нарушения": [
      {"Целевой документ": "...", "Различие": "..."}
    ]
  }

Что писать в полях:
- «Целевой документ» — что ФАКТИЧЕСКИ есть в документе (цитата) или «отсутствует»
- «Различие» — что ДОЛЖНО быть или в чём проблема

ВАЖНО:
- Верни ТОЛЬКО JSON, без пояснений.
- Одно правило = один объект ответа.
- Все нарушения по правилу собери в массив "нарушения".
"""


def load_rules() -> list:
    """Загружает все правила из JSON."""
    with open(RULES_FILE, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data.get("правила", [])


def get_rule(rules: list, index: int) -> dict:
    """Находит правило по индексу."""
    for rule in rules:
        if rule["index"] == index:
            return rule
    raise ValueError(f"Правило #{index} не найдено")


def build_context(rule: dict, target_doc: dict, template_doc: dict) -> str:
    """Формирует контекст для правила."""
    parts = []
    scope = rule["scope"]
    scopes = [scope] if isinstance(scope, str) else scope
    compare = rule["compare"]

    for s in scopes:
        # Добавляем чанк целевого документа
        parts.append(f"[TARGET_{s}]")
        parts.append(target_doc.get(s, "(чанк отсутствует)"))
        parts.append(f"[/TARGET_{s}]")

        # Для template — добавляем чанк шаблона
        if compare == "template":
            parts.append(f"[TEMPLATE_{s}]")
            parts.append(template_doc.get(s, "(чанк отсутствует)"))
            parts.append(f"[/TEMPLATE_{s}]")

    return "\n".join(parts)


def build_user_prompt(rule: dict, context: str) -> str:
    """Формирует пользовательский промпт."""
    scope = rule["scope"]
    scope_str = scope if isinstance(scope, str) else ", ".join(scope)
    instructions = "\n".join(f"- {instr}" for instr in rule.get("content", []))

    return f"""RULE_INDEX: {rule["index"]}
COMPARE: {rule["compare"]}
SCOPE: {scope_str}
RULE_TITLE: {rule["title"]}
RULE_INSTRUCTIONS:
{instructions}

CONTEXT:
{context}
"""


def call_llm(system_prompt: str, user_prompt: str, model: str) -> str:
    """Вызывает LLM."""
    client = OpenAI()
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.0
    )
    return response.choices[0].message.content


def main():
    parser = argparse.ArgumentParser(description="Тест одного правила")
    parser.add_argument("--rule", type=int, required=True, help="Номер правила (1-9)")
    parser.add_argument("--use-errors", action="store_true", help="Использовать файл с ошибками")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"Модель (default: {DEFAULT_MODEL})")
    parser.add_argument("--dry-run", action="store_true", help="Только показать промпт, без вызова LLM")
    parser.add_argument("--show-prompt", action="store_true", help="Показать промпт перед вызовом")

    args = parser.parse_args()

    # Выбираем целевой документ
    target_path = TARGET_WITH_ERRORS if args.use_errors else TARGET_DOC
    print(f"📄 Целевой: {Path(target_path).name}")
    print(f"📄 Шаблон: {Path(TEMPLATE_DOC).name}")

    # Загружаем правила
    rules = load_rules()
    rule = get_rule(rules, args.rule)
    print(f"\n📋 Правило #{rule['index']}: {rule['title']}")
    print(f"   compare: {rule['compare']}, scope: {rule['scope']}, llm: {rule.get('llm', True)}")

    # Проверяем что правило использует LLM
    if not rule.get("llm", True):
        print(f"\n⚠️  Правило #{rule['index']} не использует LLM (llm: false)")
        return

    # Парсим документы
    print("\n⏳ Парсинг документов...")
    target_doc = parse_prikaz_ic(target_path)
    template_doc = parse_prikaz_ic(TEMPLATE_DOC)

    # Строим промпт
    context = build_context(rule, target_doc, template_doc)
    user_prompt = build_user_prompt(rule, context)

    # Показываем промпт если нужно
    if args.show_prompt or args.dry_run:
        print("\n" + "=" * 70)
        print("SYSTEM PROMPT:")
        print("=" * 70)
        print(SYSTEM_PROMPT)
        print("\n" + "=" * 70)
        print("USER PROMPT:")
        print("=" * 70)
        print(user_prompt)

    if args.dry_run:
        print("\n⚠️  --dry-run: LLM не вызывается")
        return

    # Вызываем LLM
    print(f"\n🚀 Вызов LLM ({args.model})...")
    response = call_llm(SYSTEM_PROMPT, user_prompt, args.model)

    # Выводим ответ
    print("\n" + "=" * 70)
    print("ОТВЕТ МОДЕЛИ:")
    print("=" * 70)
    print(response)

    # Пробуем распарсить JSON
    print("\n" + "=" * 70)
    print("РЕЗУЛЬТАТ:")
    print("=" * 70)
    try:
        text = response.strip()
        if text.startswith("```"):
            lines = text.split("\n")[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            text = "\n".join(lines)

        parsed = json.loads(text)

        if parsed.get("status") == "ok":
            print("✅ Нарушений не найдено")
        else:
            violations = parsed.get("нарушения", [])
            print(f"❌ Найдено нарушений: {len(violations)}")
            for i, v in enumerate(violations, 1):
                print(f"\n   {i}. Целевой документ: {v.get('Целевой документ', '?')}")
                print(f"      Различие: {v.get('Различие', '?')}")

    except json.JSONDecodeError as e:
        print(f"⚠️  Ошибка парсинга JSON: {e}")


if __name__ == "__main__":
    main()
