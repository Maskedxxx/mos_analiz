#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты F3: повреждённые типы помечаются broken и не запускаются; необработанные ошибки — JSON 500.
"""
import json
import os
import shutil
from pathlib import Path

import pytest

from conftest import DOC_CONFIGS


def _make_type(root: Path, name: str, config, sections=None, rules=None):
    d = root / "doc_configs" / name
    d.mkdir(parents=True)
    (d / "config.json").write_text(config if isinstance(config, str) else json.dumps(config, ensure_ascii=False), encoding="utf-8")
    if sections is not None:
        (d / "sections.json").write_text(sections, encoding="utf-8")
    if rules is not None:
        (d / "rules_multi.json").write_text(rules, encoding="utf-8")
    return d


def test_list_doc_types_marks_broken(tmp_path):
    from src.audit.engine import AuditEngine
    good_cfg = {"doc_type": "good", "doc_title": "Хороший", "parser_by_ext": {".docx": "docx"}}
    _make_type(tmp_path, "good", good_cfg, '{"sections": {}}', '{"rules": []}')
    _make_type(tmp_path, "bad_json", '{"doc_type": "bad_json", "doc_title": "Битый')
    _make_type(tmp_path, "no_sections", {**good_cfg, "doc_type": "no_sections"}, None, '{"rules": []}')
    _make_type(tmp_path, "bad_rules", {**good_cfg, "doc_type": "bad_rules"}, '{"sections": {}}', '{"rules": [')
    _make_type(tmp_path, "special", {"doc_type": "special", "doc_title": "Спец", "engine": "kpsc"})
    types = {t["doc_type"]: t for t in AuditEngine.list_doc_types(base_dir=str(tmp_path))}
    assert "broken" not in types["good"] and "broken" not in types["special"]
    assert types["bad_json"]["broken"] and "config.json повреждён" in types["bad_json"]["broken_reason"]
    assert types["no_sections"]["broken"] and "нет файла sections.json" in types["no_sections"]["broken_reason"]
    assert types["bad_rules"]["broken"] and "rules_multi.json повреждён" in types["bad_rules"]["broken_reason"]


def _client():
    from fastapi.testclient import TestClient
    import main
    c = TestClient(main.app, raise_server_exceptions=False)
    r = c.post("/api/login", json={"login": os.environ["AUDIT_LOGIN"], "password": os.environ["AUDIT_PASSWORD"]})
    assert r.status_code == 200
    return c


def test_unknown_doc_type_is_400():
    c = _client()
    r = c.post("/api/audit", files={"file": ("a.docx", b"PK\x03\x04", "application/octet-stream")}, data={"doc_type": "no_such_type"})
    assert r.status_code == 400
    assert "Неизвестный тип документа" in r.json()["detail"]


def test_broken_doc_type_is_400(monkeypatch):
    import main
    from src.audit.engine import AuditEngine
    fake = [{"doc_type": "akt_nachala", "doc_title": "0.1 Акт", "broken": True, "broken_reason": "нет файла sections.json"}]
    monkeypatch.setattr(AuditEngine, "list_doc_types", staticmethod(lambda base_dir=None: fake))
    c = _client()
    r = c.post("/api/audit", files={"file": ("a.docx", b"PK\x03\x04", "application/octet-stream")}, data={"doc_type": "akt_nachala"})
    assert r.status_code == 400
    assert "настроен некорректно" in r.json()["detail"] and "sections.json" in r.json()["detail"]
    r = c.get("/api/types/akt_nachala/rules")
    assert r.status_code == 400


def test_unhandled_exception_is_json_500(monkeypatch):
    from src.audit.engine import AuditEngine

    def _boom(base_dir=None):
        raise RuntimeError("имитация сбоя")

    monkeypatch.setattr(AuditEngine, "list_doc_types", staticmethod(_boom))
    c = _client()
    r = c.get("/api/types")
    assert r.status_code == 500
    body = r.json()
    assert body["detail"].startswith("Внутренняя ошибка сервера: RuntimeError")
    assert body["technical"] == "имитация сбоя"
