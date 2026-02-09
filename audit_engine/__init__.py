#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
audit_engine — Единый модуль Vision-аудита документов.

Один движок для всех типов документов:
новый тип = новая папка конфигов в doc_configs/, без изменений в коде.

Использование:
    from audit_engine import AuditEngine

    engine = AuditEngine("presentation_eu")
    result = engine.run("document.pptx")
"""

from .engine import AuditEngine
from .models import RuleSpec, AuditConfig, AuditResult, load_rules, load_audit_config

__all__ = [
    'AuditEngine',
    'RuleSpec',
    'AuditConfig',
    'AuditResult',
    'load_rules',
    'load_audit_config',
]
