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


# --- F19: ошибка аудита попадает в артефакты сессии ---

def test_write_error_artifacts_creates_and_appends(tmp_path, monkeypatch):
    import src.api.server as srv
    monkeypatch.setattr(srv, "_LOGS_RESULT_DIR", tmp_path)
    sd = Path(srv._write_error_artifacts(None, "akt_nachala", "ValueError: boom", "Traceback (most recent call last):\n  x"))
    assert sd.parent == tmp_path / "akt_nachala" and sd.name.startswith("session_")
    assert "ValueError: boom" in (sd / "ERROR.txt").read_text(encoding="utf-8")
    assert "Traceback" in (sd / "pipeline.log").read_text(encoding="utf-8")
    # Повторный вызов на том же каталоге: ERROR.txt не перезаписывается, pipeline.log дописывается.
    (sd / "ERROR.txt").write_text("свой текст спецдвижка", encoding="utf-8")
    assert srv._write_error_artifacts(str(sd), "akt_nachala", "ValueError: again", "tb2") == str(sd)
    assert (sd / "ERROR.txt").read_text(encoding="utf-8") == "свой текст спецдвижка"
    assert (sd / "pipeline.log").read_text(encoding="utf-8").count("Аудит завершился с ошибкой") == 2


def test_failed_audit_has_error_txt_and_session_json(tmp_path, monkeypatch):
    """Движок падает после audit_start → в каталоге сессии ERROR.txt, traceback в pipeline.log, session.json со статусом error."""
    import time
    import src.api.server as srv
    from conftest import DOC_CONFIGS
    monkeypatch.setattr(srv, "_LOGS_RESULT_DIR", tmp_path)
    sd = tmp_path / "akt_nachala" / "session_fail"

    class _FailingEngine(srv.AuditEngine):
        """Настоящий движок (list_doc_types нужен start_audit), но run падает после audit_start."""
        def run(self, target_path, progress_callback=None, **kw):
            progress_callback("audit_start", {"doc_type": "akt_nachala", "filename": "t.docx", "session_dir": str(sd)})
            (sd).mkdir(parents=True); (sd / "pipeline.log").write_text("старт\n", encoding="utf-8")
            raise ValueError("Модель не вернула результат в ожидаемом формате после 5 попыток.")
    monkeypatch.setattr(srv, "AuditEngine", _FailingEngine)
    from fastapi.testclient import TestClient
    import main
    template = DOC_CONFIGS / "akt_nachala" / "template" / "template.docx"
    # Контекстный TestClient: цикл событий живёт, пока фоновый поток аудита шлёт события в SSE-очередь.
    with TestClient(main.app, raise_server_exceptions=False) as c:
        assert c.post("/api/login", json={"login": os.environ["AUDIT_LOGIN"], "password": os.environ["AUDIT_PASSWORD"]}).status_code == 200
        r = c.post("/api/audit", files={"file": ("t.docx", template.read_bytes(), "application/octet-stream")}, data={"doc_type": "akt_nachala"})
        assert r.status_code == 200
        sid = r.json()["session_id"]
        for _ in range(100):
            if srv.sessions[sid]["status"] == "error" and (sd / "session.json").exists(): break
            time.sleep(0.1)
    assert srv.sessions[sid]["status"] == "error"
    assert "ValueError" in (sd / "ERROR.txt").read_text(encoding="utf-8")
    log = (sd / "pipeline.log").read_text(encoding="utf-8")
    assert log.startswith("старт") and "Traceback" in log
    persisted = json.loads((sd / "session.json").read_text(encoding="utf-8"))
    assert persisted["status"] == "error" and persisted["session_dir"] == str(sd)
    assert (sd / "original" / "t.docx").exists()
    assert not (srv.UPLOAD_DIR / sid).exists(), "F28: загрузка переезжает в original/, uploads/<id>/ удаляется"
