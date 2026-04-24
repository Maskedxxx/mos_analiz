#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Smoke-тесты special XLSX engines на локальном sample-файле.

Проверяют, что специальные ветки монолита доходят до конца и создают артефакты,
даже если sample-файл не соответствует предметным правилам конкретного движка.
"""
from types import SimpleNamespace

import pandas as pd
import pytest


SAMPLE_XLSX = "docs/classification_result.xlsx"


@pytest.mark.parametrize(
    "doc_type,runner_name,report_name,rule_filter",
    [
        ("drivers", "run_drivers_special", "03_report.xlsx", None),
        ("kpsc", "run_kpsc_special", "validation_report.xlsx", "1.1"),
        ("kartochka_proekta", "run_kartochka_proekta_special", "validation_report.xlsx", "3"),
        ("plan_grafik", "run_plan_grafik_special", "validation_report.xlsx", None),
    ],
)
def test_special_xlsx_runtime_no_crash(doc_type, runner_name, report_name, rule_filter, tmp_path):
    import main

    runner = getattr(main, runner_name)
    args = SimpleNamespace(
        target=SAMPLE_XLSX,
        parse_only=False,
        rule_filter=rule_filter,
        model=None,
        temperature=None,
        session_dir=str(tmp_path / doc_type),
    )
    result = runner(args)

    assert result is not None, f"{doc_type}: runner вернул None"
    assert result.session_dir, f"{doc_type}: нет session_dir"
    report_path = result.session_dir / report_name
    assert report_path.exists(), f"{doc_type}: не создан {report_name}"
    if report_name == "validation_report.xlsx":
        df = pd.read_excel(report_path)
        if "status" in df.columns:
            bad_statuses = set(df["status"].dropna()) & {"ERROR", "MISSING"}
            assert not bad_statuses, f"{doc_type}: в отчёте остались runtime-статусы {sorted(bad_statuses)}"
