#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты F9: проверка загруженного файла до запуска аудита (веб и CLI).
"""
import os
from pathlib import Path

import pytest
from docx import Document
from openpyxl import Workbook

from conftest import DOC_CONFIGS
from src.audit.input_check import check_input_file


def test_valid_docx_passes(tmp_path):
    p = tmp_path / "ok.docx"
    Document().save(str(p))
    check_input_file(p)


def test_pdf_signature(tmp_path):
    p = tmp_path / "ok.pdf"; p.write_bytes(b"%PDF-1.4 x")
    check_input_file(p)
    bad = tmp_path / "bad.pdf"; bad.write_bytes(b"not a pdf")
    with pytest.raises(ValueError, match="не является PDF"):
        check_input_file(bad)


@pytest.mark.parametrize("name,content,fragment", [
    ("empty.docx", b"", "пустой"),
    ("zero.xlsx", b"", "пустой"),
    ("truncated.docx", b"PK\x03\x04" + b"\x00" * 100, "повреждён"),
    ("ole.docx", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 64, "защищён паролем"),
])
def test_bad_files(tmp_path, name, content, fragment):
    p = tmp_path / name; p.write_bytes(content)
    with pytest.raises(ValueError, match=fragment):
        check_input_file(p)


def test_xlsx_under_docx_mask(tmp_path):
    p = tmp_path / "renamed.docx"
    Workbook().save(str(p))
    with pytest.raises(ValueError, match="таблица Excel, а не документ Word"):
        check_input_file(p)


def test_missing_file(tmp_path):
    with pytest.raises(ValueError, match="не найден"):
        check_input_file(tmp_path / "nope.docx")


def _client():
    from fastapi.testclient import TestClient
    import main
    c = TestClient(main.app, raise_server_exceptions=False)
    assert c.post("/api/login", json={"login": os.environ["AUDIT_LOGIN"], "password": os.environ["AUDIT_PASSWORD"]}).status_code == 200
    return c


def test_api_rejects_empty_file_without_session():
    import src.api.server as srv
    c = _client()
    before = {p.name for p in srv.UPLOAD_DIR.iterdir()} if srv.UPLOAD_DIR.exists() else set()
    r = c.post("/api/audit", files={"file": ("empty.docx", b"", "application/octet-stream")}, data={"doc_type": "akt_nachala"})
    assert r.status_code == 400 and "пустой" in r.json()["detail"]
    after = {p.name for p in srv.UPLOAD_DIR.iterdir()} if srv.UPLOAD_DIR.exists() else set()
    assert after == before, "загрузка отклонённого файла должна быть удалена"
