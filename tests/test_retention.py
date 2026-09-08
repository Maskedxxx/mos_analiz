#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты F28: ротация по объёму — удаляются только целые самые старые каталоги, пока корень больше лимита;
0 — не удалять; выполняющиеся сессии (skip) не трогаются.
"""
import os
import time
from pathlib import Path

from src.audit.retention import GB, dir_size_bytes, prune_oldest


def _session(root: Path, doc_type: str, name: str, size: int, age_sec: int) -> Path:
    sd = root / doc_type / name
    (sd / "original").mkdir(parents=True)
    (sd / "original" / "doc.docx").write_bytes(b"x" * (size // 2))
    (sd / "audit_result.xlsx").write_bytes(b"y" * (size - size // 2))
    t = time.time() - age_sec
    os.utime(sd, (t, t))
    return sd


def test_prune_removes_oldest_whole_sessions(tmp_path):
    old = _session(tmp_path, "akt_nachala", "session_1", 1000, 300)
    mid = _session(tmp_path, "akt_nachala", "session_2", 1000, 200)
    new = _session(tmp_path, "forma_0_3", "session_3", 1000, 100)
    assert dir_size_bytes(tmp_path) == 3000
    removed = prune_oldest(tmp_path, 2500 / GB, "*/session_*")
    assert removed == [old]
    assert not old.exists() and mid.exists() and new.exists()
    assert (new / "original" / "doc.docx").exists(), "original/ живых сессий не трогается"


def test_zero_limit_keeps_everything(tmp_path):
    _session(tmp_path, "akt_nachala", "session_1", 1000, 300)
    assert prune_oldest(tmp_path, 0, "*/session_*") == []
    assert (tmp_path / "akt_nachala" / "session_1").exists()


def test_running_session_is_skipped(tmp_path):
    old = _session(tmp_path, "akt_nachala", "session_1", 1000, 300)
    new = _session(tmp_path, "akt_nachala", "session_2", 1000, 100)
    removed = prune_oldest(tmp_path, 500 / GB, "*/session_*", skip={"session_1"})
    assert removed == [new] and old.exists()
