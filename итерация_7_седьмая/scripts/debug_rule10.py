#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Диагностический скрипт для правила #10 (Проверка структуры приложений).

Проверяет весь пайплайн и выявляет причину ложного срабатывания:
1. Извлекает чанки через Vision
2. Показывает извлечённый текст приложения_2_регламент
3. Показывает промпт для правила #10
4. Вызывает LLM и показывает ответ
5. Анализирует расхождение

Использование:
    python scripts/debug_rule10.py
"""

import json
import os
import sys
from pathlib import Path

# Настройка путей
SCRIPT_DIR = Path(__file__).parent
PROJECT_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(PROJECT_DIR))
sys.path.insert(0, str(PROJECT_DIR.parent))

# API ключ
# os.environ["OPENAI_API_KEY"] = "YOUR_API_KEY"  # Установите через переменную окружения

from openai import OpenAI


def print_section(title: str, content: str = None):
    """Печатает секцию с заголовком."""
    print(f"\n{'='*80}")
    print(f"  {title}")
    print(f"{'='*80}")
    if content:
        print(content)


def main():
    print_section("ДИАГНОСТИКА ПРАВИЛА #10: Проверка структуры приложений")

    # =========================================================================
    # ШАГ 1: Загрузка Vision-парсера и извлечение чанков
    # =========================================================================
    print_section("ШАГ 1: Извлечение чанков через Vision Pipeline")

    from shared.vision_parser import VisionParser

    target_path = PROJECT_DIR / "documents" / "prikaz_ic_potoka_ok.docx"
    template_path = PROJECT_DIR / "templates" / "prikaz_ic_potoka_template.docx"
    vision_config_path = PROJECT_DIR / "config" / "chunks_vision.json"

    print(f"Целевой документ: {target_path}")
    print(f"Шаблон: {template_path}")
    print(f"Vision конфиг: {vision_config_path}")

    parser = VisionParser(str(vision_config_path))

    print("\nИзвлечение TARGET...")
    target_doc = parser.parse(str(target_path), parallel=False)

    print("Извлечение TEMPLATE...")
    template_doc = parser.parse(str(template_path), parallel=False)

    # =========================================================================
    # ШАГ 2: Анализ извлечённого текста приложение_2_регламент
    # =========================================================================
    print_section("ШАГ 2: Извлечённый текст приложение_2_регламент (TARGET)")

    target_reglament = target_doc.get("приложение_2_регламент", "НЕ НАЙДЕНО")
    print(f"Длина: {len(target_reglament)} символов")
    print("\n--- СОДЕРЖИМОЕ (первые 3000 символов) ---")
    print(target_reglament[:3000])
    if len(target_reglament) > 3000:
        print(f"\n... (ещё {len(target_reglament) - 3000} символов)")

    # Проверяем наличие разделов
    print("\n--- ПРОВЕРКА НАЛИЧИЯ РАЗДЕЛОВ ---")
    sections = {
        "1. Общие положения": "1. Общие положения" in target_reglament or "1.1." in target_reglament,
        "2. Порядок обновления": "2. Порядок обновления" in target_reglament or "2.1." in target_reglament,
        "3. Порядок работы": "3. Порядок работы" in target_reglament or "3.1." in target_reglament,
        "4. Проведение совещаний": "4. Проведение совещаний" in target_reglament or "4.1." in target_reglament,
    }

    for section, found in sections.items():
        status = "✅ НАЙДЕН" if found else "❌ НЕ НАЙДЕН"
        print(f"  {section}: {status}")

    # =========================================================================
    # ШАГ 3: Загрузка правила #10
    # =========================================================================
    print_section("ШАГ 3: Правило #10 из конфига")

    rules_path = PROJECT_DIR / "config" / "tz_prikaz_ic_potoka.json"
    with open(rules_path, 'r', encoding='utf-8') as f:
        rules_data = json.load(f)

    rule10 = None
    for rule in rules_data.get("правила", []):
        if rule["index"] == 10:
            rule10 = rule
            break

    if rule10:
        print(f"Title: {rule10['title']}")
        print(f"Scope: {rule10['scope']}")
        print(f"Compare: {rule10['compare']}")
        print(f"LLM: {rule10['llm']}")
        print("\nContent (инструкции):")
        for i, instr in enumerate(rule10.get("content", []), 1):
            print(f"  {i}. {instr}")
    else:
        print("❌ Правило #10 не найдено!")
        return

    # =========================================================================
    # ШАГ 4: Построение промпта для LLM
    # =========================================================================
    print_section("ШАГ 4: Промпт для LLM (правило #10)")

    # Системный промпт (из run_prikaz_ic_potoka_audit.py)
    system_prompt = """Ты — строгий аудитор документов.
Ты проверяешь ОДНО правило за один запрос.

Тебе даётся:
- фрагменты целевого документа (TARGET_*);
- фрагменты шаблона (TEMPLATE_*) — если требуется сверка;
- параметры правила: RULE_INDEX, COMPARE, SCOPE, RULE_TITLE;
- текстовые инструкции правила (RULE_INSTRUCTIONS).

COMPARE:
- template — сверяй чанк целевого с чанком шаблона (игнорируй плейсхолдеры: даты, ФИО, должности).
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

    # User промпт
    scopes = rule10["scope"] if isinstance(rule10["scope"], list) else [rule10["scope"]]
    instructions_text = "\n".join(f"- {instr}" for instr in rule10.get("content", []))

    context_parts = []
    for scope in scopes:
        target_content = target_doc.get(scope, "")
        context_parts.append(f"[TARGET_{scope}]")
        context_parts.append(target_content)
        context_parts.append(f"[/TARGET_{scope}]")

    context = "\n".join(context_parts)

    user_prompt = f"""RULE_INDEX: {rule10['index']}
COMPARE: {rule10['compare']}
SCOPE: {', '.join(scopes)}
RULE_TITLE: {rule10['title']}
RULE_INSTRUCTIONS:
{instructions_text}

CONTEXT:
{context}
"""

    print("--- SYSTEM PROMPT ---")
    print(system_prompt[:500] + "..." if len(system_prompt) > 500 else system_prompt)

    print("\n--- USER PROMPT (первые 2000 символов) ---")
    print(user_prompt[:2000])
    if len(user_prompt) > 2000:
        print(f"\n... (ещё {len(user_prompt) - 2000} символов)")

    print(f"\n📊 Статистика промпта:")
    print(f"   System prompt: {len(system_prompt)} символов")
    print(f"   User prompt: {len(user_prompt)} символов")
    print(f"   ИТОГО: {len(system_prompt) + len(user_prompt)} символов (~{(len(system_prompt) + len(user_prompt)) // 4} токенов)")

    # =========================================================================
    # ШАГ 5: Вызов LLM
    # =========================================================================
    print_section("ШАГ 5: Ответ LLM")

    client = OpenAI()

    print("Вызываем gpt-4.1-mini...")
    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.0
    )

    raw_response = response.choices[0].message.content
    print("\n--- RAW RESPONSE ---")
    print(raw_response)

    # Парсим ответ
    try:
        # Убираем markdown
        text = raw_response.strip()
        if text.startswith("```"):
            lines = text.split("\n")
            lines = lines[1:] if lines[0].startswith("```") else lines
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            text = "\n".join(lines)

        result = json.loads(text)

        print("\n--- PARSED RESULT ---")
        print(json.dumps(result, ensure_ascii=False, indent=2))

        if result.get("status") == "ok":
            print("\n✅ LLM считает что нарушений НЕТ")
        else:
            violations = result.get("нарушения", [])
            print(f"\n❌ LLM нашёл {len(violations)} нарушений")

    except json.JSONDecodeError as e:
        print(f"\n❌ Ошибка парсинга JSON: {e}")

    # =========================================================================
    # ШАГ 6: Анализ и выводы
    # =========================================================================
    print_section("ШАГ 6: АНАЛИЗ И ВЫВОДЫ")

    print("Проверка: все ли разделы 1-4 есть в извлечённом тексте?")
    all_sections_found = all(sections.values())

    if all_sections_found:
        print("✅ Все разделы 1, 2, 3, 4 ПРИСУТСТВУЮТ в извлечённом тексте")
        print("\n🔍 Возможные причины ложного срабатывания:")
        print("   1. LLM не видит разделы из-за слишком длинного контекста")
        print("   2. Инструкции правила #10 слишком строгие")
        print("   3. Формат извлечённого текста отличается от ожидаемого")
    else:
        missing = [s for s, f in sections.items() if not f]
        print(f"❌ Не найдены разделы: {missing}")
        print("\n🔍 Причина: Vision парсер не извлёк эти разделы")

    print("\n" + "="*80)
    print("  ДИАГНОСТИКА ЗАВЕРШЕНА")
    print("="*80)


if __name__ == "__main__":
    main()
