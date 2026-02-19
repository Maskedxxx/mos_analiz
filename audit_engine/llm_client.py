#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Клиент для взаимодействия с LLM (OpenAI API).

Функции:
- call_llm — отправка запроса к OpenAI Chat API
- parse_json_response — парсинг JSON-ответа от LLM
- sanitize_json_string — исправление битого JSON (переносы строк внутри строк)
"""

import json
import sys
from typing import Any, Dict, List, Optional

from openai import OpenAI


def sanitize_json_string(s: str) -> str:
    """
    Экранирует неэкранированные переносы строк внутри JSON-строк.

    LLM иногда возвращает JSON с реальными \\n внутри строковых значений,
    что ломает json.loads(). Эта функция проходит по тексту посимвольно,
    отслеживая состояние "внутри строки" / "вне строки", и заменяет
    сырые \\n, \\r, \\t на их escaped-версии.
    """
    result_chars = []
    in_string = False
    escape_next = False

    for char in s:
        if escape_next:
            result_chars.append(char)
            escape_next = False
            continue

        if char == '\\' and in_string:
            result_chars.append(char)
            escape_next = True
            continue

        if char == '"':
            in_string = not in_string
            result_chars.append(char)
            continue

        if in_string and char in '\n\r\t':
            if char == '\n':
                result_chars.append('\\n')
            elif char == '\r':
                result_chars.append('\\r')
            elif char == '\t':
                result_chars.append('\\t')
        else:
            result_chars.append(char)

    return ''.join(result_chars)


def parse_json_response(
    raw_response: str,
    spec_index: int,
    spec_title: str
) -> List[Dict[str, Any]]:
    """
    Парсит JSON-ответ от LLM.

    Обрабатывает:
    - Markdown код-блоки (```json ... ```)
    - Битый JSON (переносы строк внутри строк)
    - Формат {"status": "ok"} → пустой список
    - Формат {"status": "fail", "нарушения": [...]} → список нарушений

    Args:
        raw_response: сырой ответ от LLM
        spec_index: номер правила (для fallback)
        spec_title: название правила (для fallback)

    Returns:
        Список нарушений. Пустой список если status=ok.
    """
    text = raw_response.strip()

    # Убираем markdown код-блок
    if text.startswith("```"):
        lines = text.split("\n")
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        text = "\n".join(lines)

    # Санитизация JSON
    text = sanitize_json_string(text)

    try:
        result = json.loads(text)

        if isinstance(result, dict):
            # Статус OK — нарушений нет
            if result.get("status") == "ok":
                return []

            # Извлекаем нарушения и добавляем rule_index/rule_title
            violations = result.get("нарушения", [])
            for v in violations:
                v["rule_index"] = result.get("rule_index", spec_index)
                v["rule_title"] = result.get("rule_title", spec_title)
            return violations

        if isinstance(result, list):
            return result

        return []

    except json.JSONDecodeError as e:
        print(f"[WARN] Не удалось распарсить JSON: {e}", file=sys.stderr)
        print(f"[WARN] Ответ: {text[:200]}...", file=sys.stderr)
        return []


def call_llm(
    messages: List[Dict[str, str]],
    model: str,
    temperature: float = 0.0,
    base_url: Optional[str] = None,
    max_tokens: Optional[int] = None,
    reasoning_effort: Optional[str] = None,
    seed: Optional[int] = None
) -> str:
    """
    Вызывает LLM через OpenAI Chat API.

    Args:
        messages: список сообщений [{role, content}]
        model: идентификатор модели (gpt-4.1-mini, openai/gpt-oss-120b и т.д.)
        temperature: температура генерации (0.0 = детерминированный ответ)
        base_url: URL API (None = облачный OpenAI из env)
        max_tokens: лимит токенов генерации (None = по умолчанию провайдера)
        reasoning_effort: уровень reasoning для моделей gpt-oss ("low"/"medium"/"high")
        seed: фиксированный seed для воспроизводимости (особенно важен для MoE-моделей)

    Returns:
        Текст ответа LLM.
    """
    # base_url=None → стандартный клиент (OPENAI_API_KEY + OPENAI_BASE_URL из env)
    if base_url:
        client = OpenAI(base_url=base_url, api_key="none")
    else:
        client = OpenAI()

    kwargs: Dict[str, Any] = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
    }
    if max_tokens is not None:
        kwargs["max_tokens"] = max_tokens
    if seed is not None:
        kwargs["seed"] = seed
    # reasoning_effort передаётся через extra_body для vLLM / gpt-oss
    if reasoning_effort:
        kwargs["extra_body"] = {"reasoning_effort": reasoning_effort}

    response = client.chat.completions.create(**kwargs)  # type: ignore[arg-type]

    return response.choices[0].message.content or ""
