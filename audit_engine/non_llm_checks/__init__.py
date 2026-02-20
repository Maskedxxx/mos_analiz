#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Non-LLM проверки — проверки без обращения к LLM.

Проверки регистрируются через декоратор @register(doc_type, rule_index).
При аудите движок вызывает зарегистрированную функцию для правил с llm=False.
"""

from .registry import register, get_check, get_all_checks

# Импортируем модули с проверками для автоматической регистрации
from . import filename
from . import shapka_elements
from . import prikaz_checks
from . import prikaz_tirazh_checks
