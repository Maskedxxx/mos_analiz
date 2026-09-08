#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты F11: сквозная сверка /api/audit/cross принимает ровно 3 разных файла (находка 2.5).
Проверяются только отказы — они срабатывают до записи файлов и запуска аудита.
"""
import os

from conftest import DOC_CONFIGS

TEMPLATE = DOC_CONFIGS / "akt_nachala" / "template" / "template.docx"
NEED_3 = "Нужно загрузить 3 документа"


def _client():
    from fastapi.testclient import TestClient
    import main
    c = TestClient(main.app, raise_server_exceptions=False)
    assert c.post("/api/login", json={"login": os.environ["AUDIT_LOGIN"], "password": os.environ["AUDIT_PASSWORD"]}).status_code == 200
    return c


def _post(client, names):
    data = TEMPLATE.read_bytes()
    return client.post("/api/audit/cross", files=[("files", (n, data, "application/octet-stream")) for n in names])


def test_two_files_rejected():
    r = _post(_client(), ["a.docx", "b.docx"])
    assert r.status_code == 400 and NEED_3 in r.json()["detail"]


def test_four_files_rejected():
    r = _post(_client(), ["a.docx", "b.docx", "c.docx", "d.docx"])
    assert r.status_code == 400 and NEED_3 in r.json()["detail"]


def test_duplicate_names_rejected():
    r = _post(_client(), ["a.docx", "b.docx", "a.docx"])
    assert r.status_code == 400
    assert "повторяются: a.docx" in r.json()["detail"]
