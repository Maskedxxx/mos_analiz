#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Реестр препроцессоров текста.

Декоратор @register_preprocessor(doc_type, scope) регистрирует функцию-препроцессор.
Препроцессор вызывается перед отправкой текста в LLM (для compare="template").

Пример:
    @register_preprocessor("prikaz_ic", "текст_приказа")
    def normalize_prikaz(text: str) -> str:
        ...
"""

from typing import Callable, Dict, Optional, Tuple

# Реестр: {(doc_type, scope): preprocess_fn}
_REGISTRY: Dict[Tuple[str, str], Callable[[str], str]] = {}


def register_preprocessor(doc_type: str, scope: str):
    """
    Декоратор для регистрации препроцессора.

    Args:
        doc_type: тип документа
        scope: чанк, к которому применяется препроцессор

    Функция должна принимать str → str.
    """
    def decorator(fn: Callable[[str], str]):
        _REGISTRY[(doc_type, scope)] = fn
        return fn
    return decorator


def get_preprocessors(doc_type: str) -> Dict[str, Callable[[str], str]]:
    """
    Возвращает все препроцессоры для типа документа.

    Returns:
        Словарь {scope: preprocess_fn}
    """
    return {
        scope: fn
        for (dt, scope), fn in _REGISTRY.items()
        if dt == doc_type
    }
