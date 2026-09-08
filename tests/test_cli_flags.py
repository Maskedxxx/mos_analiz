#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тесты F18: CLI отдаёт одну строку и exit 2 вместо traceback; --rule-filter у generic-типа
предупреждает; --print-prompts печатает промпты обоих слоёв и не вызывает модель.
"""
import os
import subprocess
import sys
from pathlib import Path

import pytest

from conftest import DOC_CONFIGS

ROOT = DOC_CONFIGS.parent
TEMPLATE = DOC_CONFIGS / "akt_nachala" / "template" / "template.docx"


def _cli(*args):
    return subprocess.run([sys.executable, "main.py", *args], cwd=ROOT, capture_output=True, text=True, env=os.environ.copy())


def test_print_prompts_does_not_call_model(monkeypatch, tmp_path, capsys):
    import src.llm.multi_rule as mr
    from src.audit.engine import AuditEngine

    def _no_client(base_url=None):
        raise AssertionError("модель не должна вызываться при --print-prompts")
    monkeypatch.setattr(mr, "make_llm_client", _no_client)
    result = AuditEngine("akt_nachala").run(str(TEMPLATE), print_prompts=True, session_dir=str(tmp_path))
    out = capsys.readouterr().out
    assert out.count("USER PROMPT") == 2 and "BASE" in out and "METHODOLOGY" in out
    assert result.violations == []


def test_special_engine_with_docx_is_one_line_exit_2():
    r = _cli("--doc-type", "forma_0_3", "--target", str(TEMPLATE))
    assert r.returncode == 2
    assert "Traceback" not in r.stderr
    assert r.stderr.strip().startswith("Ошибка: файл не открывается как таблица Excel")


def test_rule_filter_on_generic_type_warns():
    r = _cli("--doc-type", "akt_nachala", "--target", str(TEMPLATE), "--rule-filter", "2", "--print-prompts")
    assert r.returncode == 0
    assert "Предупреждение: --rule-filter действует только для спецдвижков" in r.stderr
