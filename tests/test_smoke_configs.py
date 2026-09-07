#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Smoke-тесты конфигов типов документов.

Проверяют, что у каждого типа в doc_configs/ валидны config.json и файлы правил
и что у правил есть обязательные поля. Быстрые, без сети, без LLM.

Файлы правил по видам движков:
  generic-LLM  → rules_multi.json (базовый слой) + rules_methodology.json (методический слой)
                 + необязательный rules_custom.json (правила, добавленные через интерфейс);
  special      → validation_rules.json (python-валидаторы по xlsx);
  crosscheck   → rules.json (сквозная сверка комплекта).
"""
import json
from pathlib import Path

import pytest

from conftest import DOC_CONFIGS, PROJECT_ROOT


DOC_TYPES = sorted(p.name for p in DOC_CONFIGS.iterdir() if p.is_dir())

# Имена файлов правил в порядке поиска. Первый найденный — основной файл типа.
RULE_FILE_NAMES = ("rules_multi.json", "validation_rules.json", "rules.json")

# Обязательные поля одного правила generic-типа (rules_multi.json / rules_methodology.json).
GENERIC_RULE_REQUIRED = ("index", "title", "check")


def _load_json(path: Path) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _rules_list(data: dict, path: Path) -> list:
    """Достаёт список правил из файла любого вида (ключ 'rules', 'правила' или 'validation_rules')."""
    rules = data.get("rules") or data.get("правила") or data.get("validation_rules", [])
    assert isinstance(rules, list), f"{path}: правила не список"
    return rules


@pytest.mark.parametrize("doc_type", DOC_TYPES)
def test_config_json_valid(doc_type):
    """config.json существует, парсится, doc_type совпадает с именем папки, есть doc_title."""
    cfg_path = DOC_CONFIGS / doc_type / "config.json"
    assert cfg_path.exists(), f"Нет config.json для {doc_type}"
    data = _load_json(cfg_path)
    assert isinstance(data, dict)
    assert data.get("doc_type") == doc_type, "doc_type в config != имя папки"
    assert "doc_title" in data, "Нет doc_title"


@pytest.mark.parametrize("doc_type", DOC_TYPES)
def test_rules_file_valid(doc_type):
    """Основной файл правил типа существует, парсится и содержит непустой список правил."""
    for name in RULE_FILE_NAMES:
        candidate = DOC_CONFIGS / doc_type / name
        if candidate.exists():
            rules_path = candidate
            break
    else:
        pytest.skip(
            f"{doc_type}: нет файла правил ({', '.join(RULE_FILE_NAMES)}) — "
            "движок без декларативных правил (например, drivers, plan_grafik)"
        )

    rules = _rules_list(_load_json(rules_path), rules_path)
    assert len(rules) > 0, f"{rules_path.name}: пустой список правил"


@pytest.mark.parametrize("doc_type", DOC_TYPES)
def test_generic_rules_structure(doc_type):
    """У generic-типов каждое правило базового и методического слоя имеет index, title, check;
    индексы внутри файла уникальны."""
    multi = DOC_CONFIGS / doc_type / "rules_multi.json"
    if not multi.exists():
        pytest.skip(f"{doc_type}: не generic-тип (нет rules_multi.json)")

    for name in ("rules_multi.json", "rules_methodology.json", "rules_custom.json"):
        path = DOC_CONFIGS / doc_type / name
        if not path.exists():
            continue
        rules = _rules_list(_load_json(path), path)
        seen = set()
        for rule in rules:
            assert isinstance(rule, dict), f"{path}: правило не объект: {rule!r}"
            for key in GENERIC_RULE_REQUIRED:
                assert key in rule, f"{path}: у правила нет поля {key!r}: {rule.get('title')!r}"
            assert rule["index"] not in seen, f"{path}: дубликат index {rule['index']}"
            seen.add(rule["index"])
