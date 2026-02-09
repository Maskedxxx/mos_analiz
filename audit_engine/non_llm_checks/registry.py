#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Реестр non-LLM проверок.

Декоратор @register(doc_type, rule_index) регистрирует функцию-проверку.
Движок вызывает get_check(doc_type, rule_index) → callable или None.

Пример:
    @register("prikaz_ic", 1)
    def check_filename_prikaz(target_doc, config):
        ...
"""

from typing import Any, Callable, Dict, List, Optional, Tuple

# Реестр: {(doc_type, rule_index): check_fn}
_REGISTRY: Dict[Tuple[str, int], Callable] = {}


def register(doc_type: str, rule_index: int):
    """
    Декоратор для регистрации non-LLM проверки.

    Args:
        doc_type: тип документа (prikaz_ic, cheklist_eu, presentation_eu, ...)
        rule_index: номер правила

    Функция проверки должна принимать (target_doc: Dict, config: AuditConfig) -> List[Dict]
    """
    def decorator(fn: Callable):
        _REGISTRY[(doc_type, rule_index)] = fn
        return fn
    return decorator


def get_check(doc_type: str, rule_index: int) -> Optional[Callable]:
    """Возвращает зарегистрированную проверку или None."""
    return _REGISTRY.get((doc_type, rule_index))


def get_all_checks(doc_type: str) -> Dict[int, Callable]:
    """Возвращает все проверки для типа документа {rule_index: check_fn}."""
    return {
        idx: fn
        for (dt, idx), fn in _REGISTRY.items()
        if dt == doc_type
    }
