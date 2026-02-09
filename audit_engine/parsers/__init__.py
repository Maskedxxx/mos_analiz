#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Реестр парсеров для вторичных файлов (XLSX, CSV и др.).

Парсеры регистрируются по имени. Движок вызывает get_parser(name)
для получения функции parse(file_path) -> dict.
"""

import importlib
from typing import Callable, Dict

# Реестр: имя парсера → полный путь модуля
PARSERS: Dict[str, str] = {
    "grafik_obhod": "audit_engine.parsers.grafik_obhod",
}


def get_parser(name: str) -> Callable[[str], Dict[str, str]]:
    """
    Возвращает функцию parse() по имени парсера.

    Args:
        name: имя парсера (ключ в PARSERS)

    Returns:
        Функция parse(file_path) -> dict с чанками

    Raises:
        KeyError: если парсер не зарегистрирован
    """
    if name not in PARSERS:
        raise KeyError(
            f"Парсер '{name}' не зарегистрирован. "
            f"Доступные: {list(PARSERS.keys())}"
        )
    module = importlib.import_module(PARSERS[name])
    return module.parse
