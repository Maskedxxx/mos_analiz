#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Smoke-тесты спецдвижков (python-валидаторы по xlsx-формам).

Проверяют, что каждый спецдвижок доходит до конца на своём образце документа
и создаёт отчёт без runtime-ошибок (статусы ERROR/MISSING в отчёте — падение теста).

Образцы — реальные документы предприятий, в git не хранятся (персональные данные).
Ожидаются в tests/data/<doc_type>/sample.xlsx; без образца тест пропускается.
Откуда взять образцы — tests/data/README.md.
"""
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import pytest

from conftest import PROJECT_ROOT

DATA_DIR = PROJECT_ROOT / "tests" / "data"


@pytest.mark.parametrize(
    "doc_type,report_name,rule_filter",
    [
        ("drivers", "03_report.xlsx", None),
        ("kpsc", "validation_report.xlsx", "1.1"),
        ("kartochka_proekta_0_2", "validation_report.xlsx", None),
        ("kartochka_proekta_2_4", "validation_report.xlsx", "3"),
        ("plan_grafik", "validation_report.xlsx", None),
        ("forma_0_3", "validation_report.xlsx", None),
        ("forma_0_4", "validation_report.xlsx", None),
        ("list_prisutstviya_modul_1", "validation_report.xlsx", None),
    ],
)
def test_special_xlsx_runtime_no_crash(doc_type, report_name, rule_filter, tmp_path):
    """Спецдвижок на своём образце создаёт отчёт и не оставляет runtime-статусов."""
    sample = DATA_DIR / doc_type / "sample.xlsx"
    if not sample.exists():
        pytest.skip(f"нет фикстуры {sample.relative_to(PROJECT_ROOT)} — см. tests/data/README.md")

    from src.engines import SPECIAL_ENGINE_RUNNERS, detect_engine

    runner = SPECIAL_ENGINE_RUNNERS[detect_engine(doc_type)]
    args = SimpleNamespace(
        doc_type=doc_type,
        target=str(sample),
        parse_only=False,
        rule_filter=rule_filter,
        model=None,
        temperature=None,
        session_dir=str(tmp_path / doc_type),
    )
    result = runner(args)

    assert result is not None, f"{doc_type}: runner вернул None"
    assert result.session_dir, f"{doc_type}: нет session_dir"
    report_path = Path(result.session_dir) / report_name
    assert report_path.exists(), f"{doc_type}: не создан {report_name}"
    if report_name == "validation_report.xlsx":
        df = pd.read_excel(report_path)
        if "status" in df.columns:
            bad_statuses = set(df["status"].dropna()) & {"ERROR", "MISSING"}
            assert not bad_statuses, f"{doc_type}: в отчёте остались runtime-статусы {sorted(bad_statuses)}"
