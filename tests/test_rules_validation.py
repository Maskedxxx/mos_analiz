#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты F13 (валидация правил, 404 для неизвестного типа) и F4 (битый rules_custom.json → 409, .bak).
Файл правил подменяется на временный через monkeypatch _DOC_CONFIGS_DIR не трогаем — используем
временную копию каталога типа.
"""
import json
import os
import shutil
from pathlib import Path

import pytest

from conftest import DOC_CONFIGS


def _client():
    from fastapi.testclient import TestClient
    import main
    c = TestClient(main.app, raise_server_exceptions=False)
    assert c.post("/api/login", json={"login": os.environ["AUDIT_LOGIN"], "password": os.environ["AUDIT_PASSWORD"]}).status_code == 200
    return c


@pytest.fixture
def tmp_type(tmp_path, monkeypatch):
    """Временный каталог типов с копией akt_nachala (без правок живого doc_configs)."""
    import src.api.server as srv
    import src.audit.engine as eng
    root = tmp_path / "doc_configs"; root.mkdir()
    shutil.copytree(DOC_CONFIGS / "akt_nachala", root / "akt_nachala", ignore=shutil.ignore_patterns("template", "*.bak"))
    (root / "akt_nachala" / "rules_custom.json").write_text('{"rules": []}', encoding="utf-8")
    monkeypatch.setattr(srv, "_DOC_CONFIGS_DIR", root)
    monkeypatch.setattr(eng, "_DOC_CONFIGS_DIR", root)
    return root / "akt_nachala"


def test_unknown_type_404(tmp_type):
    c = _client()
    assert c.get("/api/types/no_such/rules").status_code == 404
    r = c.post("/api/types/no_such/rules", json={"title": "t", "check": "c", "target_sections": ["x"]})
    assert r.status_code == 404 and "не найден" in r.json()["detail"]


@pytest.mark.parametrize("title,check,frag", [("", "c", "Заголовок"), ("   ", "c", "Заголовок"), ("t", "", "Текст проверки"), ("t", "   ", "Текст проверки"), ("t", "x" * 4001, "слишком длинный")])
def test_create_validation(tmp_type, title, check, frag):
    c = _client()
    r = c.post("/api/types/akt_nachala/rules", json={"title": title, "check": check, "target_sections": ["основной_текст"]})
    assert r.status_code == 400 and frag in r.json()["detail"]
    assert json.loads((tmp_type / "rules_custom.json").read_text(encoding="utf-8"))["rules"] == []


def test_draft_validation(tmp_type):
    c = _client()
    r = c.post("/api/types/akt_nachala/rules/draft", json={"title": " ", "raw_check": "", "sections": ["основной_текст"]})
    assert r.status_code == 400


def test_broken_custom_409_and_not_overwritten(tmp_type):
    rc = tmp_type / "rules_custom.json"; rc.write_text('{"rules": [', encoding="utf-8")
    c = _client()
    assert c.get("/api/types/akt_nachala/rules").status_code == 409
    r = c.post("/api/types/akt_nachala/rules", json={"title": "t", "check": "c", "target_sections": ["основной_текст"]})
    assert r.status_code == 409 and "повреждён" in r.json()["detail"]
    assert rc.read_text(encoding="utf-8") == '{"rules": ['  # не перезаписан


def test_save_creates_bak(tmp_type):
    rc = tmp_type / "rules_custom.json"
    c = _client()
    r = c.post("/api/types/akt_nachala/rules", json={"title": "Новое", "check": "проверить", "target_sections": ["основной_текст"]})
    assert r.status_code == 200, r.text
    assert (tmp_type / "rules_custom.json.bak").read_text(encoding="utf-8") == '{"rules": []}'
    assert len(json.loads(rc.read_text(encoding="utf-8"))["rules"]) == 1


def test_broken_custom_gives_audit_warning(tmp_type):
    from src.llm.multi_rule import load_multi_rule_config
    (tmp_type / "rules_custom.json").write_text('{"rules": [', encoding="utf-8")
    cfg = load_multi_rule_config(tmp_type.parent, "akt_nachala")
    assert cfg["warnings"] and "rules_custom.json повреждён" in cfg["warnings"][0]
