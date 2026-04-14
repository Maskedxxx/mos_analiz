#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Smoke-тесты конфигов типов документов.

Проверяют что все config.json / rules.json валидны и имеют обязательные поля.
Быстрые, без сети, без LLM.
"""
import json
from pathlib import Path

import pytest

from conftest import DOC_CONFIGS, PROJECT_ROOT


DOC_TYPES = sorted(p.name for p in DOC_CONFIGS.iterdir() if p.is_dir())


@pytest.mark.parametrize("doc_type", DOC_TYPES)
def test_config_json_valid(doc_type):
    """config.json существует и парсится."""
    cfg_path = DOC_CONFIGS / doc_type / "config.json"
    assert cfg_path.exists(), f"Нет config.json для {doc_type}"
    with open(cfg_path, encoding="utf-8") as f:
        data = json.load(f)
    assert isinstance(data, dict)
    assert data.get("doc_type") == doc_type, f"doc_type в config != имя папки"
    assert "doc_title" in data, "Нет doc_title"


@pytest.mark.parametrize("doc_type", DOC_TYPES)
def test_rules_json_valid(doc_type):
    """rules.json или validation_rules.json существует и парсится."""
    # Пробуем оба имени — kpsc/kartochka_proekta используют validation_rules.json
    for name in ("rules.json", "validation_rules.json"):
        candidate = DOC_CONFIGS / doc_type / name
        if candidate.exists():
            rules_path = candidate
            break
    else:
        pytest.skip(f"Нет файла правил для {doc_type} (спецмодуль без rules.json)")

    with open(rules_path, encoding="utf-8") as f:
        data = json.load(f)
    # Допустимые ключи с правилами: "правила" (ру) или "rules" (en)
    rules = data.get("правила") or data.get("rules") or data.get("validation_rules", [])
    assert isinstance(rules, list), f"{rules_path.name}: правила не список"
    assert len(rules) > 0, f"{rules_path.name}: пустой список правил"
