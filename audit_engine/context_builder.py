#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Построение контекста для LLM-проверок.

Формирует user_prompt для каждого правила:
- Выбирает нужные чанки из target_doc и template_doc
- Применяет context_filter (фильтрация параграфов по паттернам)
- Применяет препроцессоры (нормализация текста перед сравнением)
- Оборачивает в теги [TARGET_*] / [TEMPLATE_*]
"""

import re
from typing import Any, Callable, Dict, List, Optional

from .models import RuleSpec


def extract_matching_paragraphs(
    text: str,
    patterns: List[str],
    headers_only: bool = False
) -> str:
    """
    Извлекает из текста только строки/пункты, соответствующие паттернам.

    Используется для context_filter — сужение контекста перед отправкой в LLM.

    Args:
        text: исходный текст чанка
        patterns: список regex-паттернов для фильтрации
        headers_only: если True — только заголовки, иначе заголовок + тело до следующей секции
    """
    if not patterns:
        return text

    lines = text.split('\n')
    result_lines = []
    compiled_patterns = [re.compile(p) for p in patterns]

    if headers_only:
        # Режим "только заголовки" — берём только совпавшие строки
        for line in lines:
            stripped = line.strip()
            if any(p.match(stripped) for p in compiled_patterns):
                result_lines.append(stripped)
    else:
        # Режим "параграфы" — захватываем заголовок + всё до следующей секции
        capturing = False
        new_section_pattern = re.compile(r'^(\d+\.|\d+\.\d+\.?)\s')

        for line in lines:
            stripped = line.strip()
            matches_our_pattern = any(p.match(stripped) for p in compiled_patterns)

            if matches_our_pattern:
                capturing = True
                result_lines.append(line)
            elif capturing:
                # Если встретили новую секцию, которая НЕ в наших паттернах — стоп
                if new_section_pattern.match(stripped) and not matches_our_pattern:
                    capturing = False
                else:
                    result_lines.append(line)

    filtered_text = '\n'.join(result_lines).strip()

    if not filtered_text:
        return f"[Фильтр: не найдено пунктов по паттернам {patterns}]"

    return filtered_text


def build_context_for_rule(
    spec: RuleSpec,
    target_doc: Dict[str, Any],
    template_doc: Dict[str, Any],
    preprocessors: Optional[Dict[str, Callable[[str], str]]] = None
) -> str:
    """
    Строит контекст для правила в зависимости от типа проверки.

    Args:
        spec: спецификация правила
        target_doc: распарсенный целевой документ
        template_doc: распарсенный шаблон
        preprocessors: словарь {scope: preprocess_fn} для нормализации текста

    Логика по compare:
    - target_only: только TARGET_<scope>
    - template: TARGET_<scope> + TEMPLATE_<scope> (с нормализацией)
    - cross_check: несколько TARGET_<scope> для сравнения между собой
    """
    context_parts = []
    scopes = [spec.scope] if isinstance(spec.scope, str) else spec.scope
    preprocessors = preprocessors or {}

    def apply_filter(content: str, scope: str) -> str:
        """Применяет context_filter если задан для данного scope."""
        if spec.context_filter and scope in spec.context_filter:
            patterns = spec.context_filter[scope]
            headers_only = spec.context_filter_mode == "headers_only"
            return extract_matching_paragraphs(content, patterns, headers_only)
        return content

    def apply_preprocessor(content: str, scope: str) -> str:
        """Применяет препроцессор если зарегистрирован для scope."""
        fn = preprocessors.get(scope)
        if fn:
            return fn(content)
        return content

    if spec.compare == "target_only":
        for scope in scopes:
            content = target_doc.get(scope, "")
            content = apply_filter(str(content), scope)
            context_parts.append(f"[TARGET_{scope}]")
            context_parts.append(content)
            context_parts.append(f"[/TARGET_{scope}]")

    elif spec.compare == "template":
        for scope in scopes:
            target_content = str(target_doc.get(scope, ""))
            template_content = str(template_doc.get(scope, ""))

            # Фильтрация контекста
            target_content = apply_filter(target_content, scope)
            template_content = apply_filter(template_content, scope)

            # Препроцессинг (нормализация плейсхолдеров и т.д.)
            target_content = apply_preprocessor(target_content, scope)
            template_content = apply_preprocessor(template_content, scope)

            context_parts.append(f"[TARGET_{scope}]")
            context_parts.append(target_content)
            context_parts.append(f"[/TARGET_{scope}]")

            context_parts.append(f"[TEMPLATE_{scope}]")
            context_parts.append(template_content)
            context_parts.append(f"[/TEMPLATE_{scope}]")

    elif spec.compare == "cross_check":
        for scope in scopes:
            content = str(target_doc.get(scope, ""))
            content = apply_filter(content, scope)
            context_parts.append(f"[TARGET_{scope}]")
            context_parts.append(content)
            context_parts.append(f"[/TARGET_{scope}]")

    return "\n".join(context_parts)


def build_user_prompt(
    spec: RuleSpec,
    target_doc: Dict[str, Any],
    template_doc: Dict[str, Any],
    preprocessors: Optional[Dict[str, Callable[[str], str]]] = None
) -> str:
    """
    Формирует полный user prompt для LLM.

    Структура:
        RULE_INDEX: <номер>
        COMPARE: <тип>
        SCOPE: <чанки>
        RULE_TITLE: <заголовок>
        RULE_INSTRUCTIONS:
        - инструкция 1
        - инструкция 2

        CONTEXT:
        [TARGET_scope] ... [/TARGET_scope]
        [TEMPLATE_scope] ... [/TEMPLATE_scope]
    """
    scope_str = spec.scope if isinstance(spec.scope, str) else ", ".join(spec.scope)
    instructions_text = "\n".join(f"- {instr}" for instr in spec.instructions)
    context = build_context_for_rule(spec, target_doc, template_doc, preprocessors)

    prompt = f"""RULE_INDEX: {spec.index}
COMPARE: {spec.compare}
SCOPE: {scope_str}
RULE_TITLE: {spec.title}
RULE_INSTRUCTIONS:
{instructions_text}

CONTEXT:
{context}
"""
    return prompt
