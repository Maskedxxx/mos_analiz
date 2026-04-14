#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Multi-rule runner — один LLM-вызов на документ со всеми правилами сразу.

Альтернатива классическому per-rule pipeline. Использует:
- sections.json: семантическая карта секций документа (start/end маркеры)
- rules_multi.json: список правил с target_section + check + exclusions
- LLM отвечает массивом JSON: [{rule_index, reasoning, verdict: {status, нарушения}}]

Ключевое отличие от legacy:
- НЕ чанкинг по scope — весь текст в промпте
- НЕ per-rule запросы — один вызов
- НЕТ template-comparison — только target_only семантика
"""
from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from openai import OpenAI

logger = logging.getLogger(__name__)


SYSTEM_PROMPT = """Ты — строгий аудитор документов.

В одном запросе тебе даётся ВЕСЬ документ и СПИСОК ПРАВИЛ для проверки.

Документ логически разделён на именованные СЕКЦИИ. Для каждой секции даны границы (маркеры начала/конца) и описание. Секции НЕ размечены в тексте — ты должен определить их сам по маркерам.

Для каждого правила указана ЦЕЛЕВАЯ СЕКЦИЯ (target_section). Анализируй ТОЛЬКО эту секцию документа, даже если похожий текст встречается в других местах. Если у правила несколько target_sections — проверяй их совокупно (достаточно выполнения в любой из них, если явно не сказано иное).

Формат ответа — СТРОГО JSON-массив, ПО ОДНОМУ объекту на правило, в том же порядке, что и в списке правил. У каждого объекта ДВА обязательных поля: «reasoning» и «verdict».

Пример для ok:
{
  "rule_index": 1,
  "reasoning": "Краткое обоснование в 1-3 предложения.",
  "verdict": {
    "status": "ok"
  }
}

Пример для fail:
{
  "rule_index": 3,
  "reasoning": "Краткое обоснование: где и что не так.",
  "verdict": {
    "status": "fail",
    "нарушения": [
      {"Целевой документ": "<цитата>", "Различие": "<что не так>"}
    ]
  }
}

Поля:
- «rule_index» — номер правила из списка (обязательно)
- «reasoning» — КРАТКОЕ рассуждение (1-3 предложения): какую секцию смотрел, что нашёл/не нашёл, почему такой вердикт. Без пересказа содержимого документа и правила.
- «verdict» — объект с финальным вердиктом (ТОЛЬКО структурированный, никаких рассуждений):
    - «status»: «ok» / «fail» / «error»
    - «нарушения»: массив {«Целевой документ», «Различие»}, только при «fail»
    - «Целевой документ»: цитата из документа или «отсутствует»
    - «Различие»: что не так или чего не хватает
- «error» — status для случая, когда правило технически невозможно проверить (нет нужной секции в документе). Добавь поле «reason».

ВАЖНО:
- Ответ — ТОЛЬКО JSON-массив, без markdown-ограждения и без пояснений до/после массива.
- В поле «verdict» — СТРОГО структурированные данные. Никаких «проверю ещё раз», «пересмотрю». Все такие мысли — в «reasoning».
- В «reasoning» НЕ включай JSON-объекты и не пытайся там формировать ответ — это свободный текст для размышлений.
- Не пропускай правила — в массиве должно быть ровно столько объектов, сколько правил.
- Не смешивай правила между собой.
- Игнорируй OCR-артефакты (пробелы между буквами, склейку строк, дублирование) — оценивай смысл.
- Уважай target_section правила: не ищи нарушения вне указанной секции."""


def build_user_prompt(filename: str, doc_text: str, sections: dict, rules: list) -> str:
    """Собирает user-prompt из FILENAME + DOCUMENT + SECTIONS + RULES."""
    parts = ["## FILENAME", filename, ""]
    parts += ["## DOCUMENT", doc_text, ""]
    parts.append("## SECTIONS")
    for name, meta in sections.items():
        parts.append(f"- **{name}**: {meta['description']}")
        parts.append(f"    start: {meta['start']}")
        parts.append(f"    end: {meta['end']}")
    parts.append("")
    parts.append("## RULES")
    for rule in rules:
        parts.append(f"### RULE {rule['index']} — {rule['title']}")
        if "target_section" in rule:
            parts.append(f"target_section: {rule['target_section']}")
        elif "target_sections" in rule:
            parts.append(f"target_sections: {', '.join(rule['target_sections'])}")
        parts.append(f"check: {rule['check']}")
        if rule.get("exclusions"):
            parts.append(f"exclusions: {rule['exclusions']}")
        parts.append("")
    return "\n".join(parts)


def _collect_doc_text(parsed: Dict[str, Any], include_scopes: Optional[List[str]]) -> str:
    """
    Собирает текст документа из parsed_docs для multi-rule.
    - Если include_scopes задан — берём только указанные поля.
    - Иначе — все непустые scope'ы кроме имя_файла/путь, с дедупликацией.
    """
    if include_scopes:
        text_keys = [k for k in include_scopes if k in parsed]
    else:
        text_keys = [k for k in parsed.keys() if k not in ("имя_файла", "путь")]
    unique_texts = list(dict.fromkeys(
        parsed[k] for k in text_keys if isinstance(parsed.get(k), str)
    ))
    return "\n\n".join(unique_texts) if unique_texts else ""


def _parse_response(response_text: str) -> List[Dict[str, Any]]:
    """
    Парсит JSON-массив вердиктов от LLM, снимая ```json обёртку если есть.

    Fallback: если LLM вернул невалидный JSON, пытаемся извлечь отдельные
    объекты верхнего уровня через regex — LLM иногда склеивает или ломает
    структуру между элементами массива.
    """
    import re

    clean = response_text.strip()
    if clean.startswith("```"):
        parts = clean.split("```", 2)
        if len(parts) >= 2:
            clean = parts[1]
            if clean.lstrip().startswith("json"):
                clean = clean.lstrip()[4:]

    clean = clean.strip()

    # Попытка 1: прямой json.loads
    try:
        parsed = json.loads(clean)
        if isinstance(parsed, list):
            return parsed
        if isinstance(parsed, dict):
            return [parsed]  # Одно правило — завернём в список
    except json.JSONDecodeError:
        pass

    # Попытка 2: выдираем отдельные объекты через сбалансированный поиск
    verdicts = []
    i = 0
    while i < len(clean):
        if clean[i] != '{':
            i += 1
            continue
        depth = 0
        start = i
        in_str = False
        escape = False
        while i < len(clean):
            ch = clean[i]
            if escape:
                escape = False
            elif ch == '\\':
                escape = True
            elif ch == '"':
                in_str = not in_str
            elif not in_str:
                if ch == '{':
                    depth += 1
                elif ch == '}':
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
            i += 1
        obj_text = clean[start:i]
        try:
            verdicts.append(json.loads(obj_text))
        except json.JSONDecodeError:
            logger.warning(f"Пропущен невалидный JSON-объект: {obj_text[:100]}...")

    if not verdicts:
        raise json.JSONDecodeError("Не удалось распарсить ни одного объекта", clean, 0)
    return verdicts


def _verdict_to_violations(verdicts: List[Dict[str, Any]], rules: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Конвертирует массив verdicts из multi-rule формата в формат legacy violations.

    Legacy format (как в старом engine.py):
        [{"правило": "...", "Целевой документ": "...", "Различие": "..."}]

    Multi-rule format:
        [{"rule_index": 1, "verdict": {"status": "fail", "нарушения": [...]}, "reasoning": "..."}]
    """
    # Индекс правил для лукапа
    rules_by_idx = {r["index"]: r for r in rules}

    violations = []
    for v in verdicts:
        idx = v.get("rule_index")
        rule = rules_by_idx.get(idx, {})
        rule_title = rule.get("title", f"Правило {idx}")
        verdict_obj = v.get("verdict", {}) if isinstance(v.get("verdict"), dict) else {}
        status = verdict_obj.get("status", "?")

        if status == "fail":
            for violation in verdict_obj.get("нарушения", []):
                violations.append({
                    "правило": rule_title,
                    "rule_index": idx,
                    "Целевой документ": violation.get("Целевой документ", "отсутствует"),
                    "Различие": violation.get("Различие", "?"),
                })

    return violations


def run_multi_rule_audit(
    *,
    parsed: Dict[str, Any],
    sections: Dict[str, Dict[str, str]],
    rules: List[Dict[str, Any]],
    include_scopes: Optional[List[str]],
    filename: str,
    llm_base_url: str,
    llm_model: str = "Qwen3.5-35B-A3B",
    session_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Запускает multi-rule audit: сборка промпта → LLM → парсинг вердикта.

    Returns:
        {"verdicts": [...], "violations": [...], "usage": {...}, "prompt_chars": int}

    Side effects:
        Если указан session_dir — сохраняет system_prompt.txt, user_prompt.txt,
        response_raw.txt, response_parsed.json.
    """
    doc_text = _collect_doc_text(parsed, include_scopes)
    user_prompt = build_user_prompt(filename, doc_text, sections, rules)

    # Сохраняем промпты для отладки
    if session_dir:
        session_dir = Path(session_dir)
        session_dir.mkdir(parents=True, exist_ok=True)
        (session_dir / "multi_rule_system_prompt.txt").write_text(SYSTEM_PROMPT, encoding="utf-8")
        (session_dir / "multi_rule_user_prompt.txt").write_text(user_prompt, encoding="utf-8")

    # LLM-вызов
    client = OpenAI(base_url=llm_base_url, api_key="dummy")
    response = client.chat.completions.create(
        model=llm_model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.7,
        top_p=0.8,
        presence_penalty=1.5,
        max_tokens=8192,
        seed=42,
        extra_body={
            "chat_template_kwargs": {"enable_thinking": False},
            "top_k": 20,
            "min_p": 0.0,
            "repetition_penalty": 1.0,
        },
    )

    response_text = response.choices[0].message.content or ""

    if session_dir:
        (session_dir / "multi_rule_response_raw.txt").write_text(response_text, encoding="utf-8")

    verdicts = _parse_response(response_text)

    if session_dir:
        (session_dir / "multi_rule_response_parsed.json").write_text(
            json.dumps(verdicts, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    violations = _verdict_to_violations(verdicts, rules)

    return {
        "verdicts": verdicts,
        "violations": violations,
        "usage": {
            "prompt_tokens": response.usage.prompt_tokens,
            "completion_tokens": response.usage.completion_tokens,
            "total_tokens": response.usage.total_tokens,
        },
        "prompt_chars": len(user_prompt),
    }


def load_multi_rule_config(doc_configs_dir: Path, doc_type: str) -> Optional[Dict[str, Any]]:
    """
    Загружает multi-rule конфиг (sections.json + rules_multi.json) для типа документа.

    Returns:
        {"sections": ..., "rules": ..., "include_scopes": ...} или None если multi-rule
        не настроен для этого типа.
    """
    base = Path(doc_configs_dir) / doc_type
    sections_path = base / "sections.json"
    rules_path = base / "rules_multi.json"

    if not sections_path.exists() or not rules_path.exists():
        return None

    sections_data = json.loads(sections_path.read_text(encoding="utf-8"))
    rules_data = json.loads(rules_path.read_text(encoding="utf-8"))

    return {
        "sections": sections_data["sections"],
        "rules": rules_data["rules"],
        "include_scopes": rules_data.get("include_scopes"),
    }
