#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты F12: завершённая сессия переживает рестарт — /result и /download читают session.json с диска.
"""
import json
import os
from pathlib import Path

import pytest
from openpyxl import Workbook


def _client():
    from fastapi.testclient import TestClient
    import main
    c = TestClient(main.app, raise_server_exceptions=False)
    assert c.post("/api/login", json={"login": os.environ["AUDIT_LOGIN"], "password": os.environ["AUDIT_PASSWORD"]}).status_code == 200
    return c


def _make_session_dir(root: Path, sid: str, status="done"):
    sd = root / "akt_nachala" / f"session_20260908_000000_{sid}"
    sd.mkdir(parents=True)
    Workbook().save(str(sd / "audit_result.xlsx"))
    data = {"session_id": sid, "doc_type": "akt_nachala", "filename": "a.docx", "status": status,
            "result": {"violations": [], "rules_checked": 10, "duration_sec": 1.0, "warnings": [], "unchecked": []},
            "error_message": "Документ не содержит распознаваемого текста" if status == "error" else None,
            "error_technical": "ValueError: …" if status == "error" else None, "session_dir": str(sd)}
    (sd / "session.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return sd


def test_persist_and_find_roundtrip(tmp_path, monkeypatch):
    import src.api.server as srv
    monkeypatch.setattr(srv, "_LOGS_RESULT_DIR", tmp_path)
    sd = tmp_path / "akt_nachala" / "session_x"; sd.mkdir(parents=True)
    session = {"doc_type": "akt_nachala", "filename": "a.docx", "status": "done", "result": {"violations": []},
               "session_dir": str(sd), "events": object(), "loop": object()}
    srv._persist_session("abcd1234", session)
    found = srv._find_session_on_disk("abcd1234")
    assert found and found["status"] == "done" and found["result"] == {"violations": []}
    assert "events" not in found and "loop" not in found
    assert srv._find_session_on_disk("nope") is None


def test_result_and_download_after_restart(tmp_path, monkeypatch):
    import src.api.server as srv
    monkeypatch.setattr(srv, "_LOGS_RESULT_DIR", tmp_path)
    srv.sessions.pop("done1234", None)
    _make_session_dir(tmp_path, "done1234")
    c = _client()
    r = c.get("/api/audit/done1234/result")
    assert r.status_code == 200 and r.json()["rules_checked"] == 10
    r = c.get("/api/audit/done1234/download")
    assert r.status_code == 200 and r.content[:2] == b"PK"


def test_error_session_after_restart(tmp_path, monkeypatch):
    import src.api.server as srv
    monkeypatch.setattr(srv, "_LOGS_RESULT_DIR", tmp_path)
    _make_session_dir(tmp_path, "err01234", status="error")
    c = _client()
    r = c.get("/api/audit/err01234/result")
    assert r.status_code == 500 and "не содержит распознаваемого текста" in r.json()["detail"]
    assert c.get("/api/audit/zzzz9999/result").status_code == 404
