#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Спецдвижок: 2.6 План-график мероприятий.

Пайплайн:
  1. Парсинг XLSX → извлечение данных из ячеек
  2. Валидация 9 правил (все non-LLM, детерминированные)
  3. Генерация Excel-отчёта

Точка входа: run(args) → AuditResult
"""

import json
import logging
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

from ..models import AuditResult

logger = logging.getLogger(__name__)


def run(args) -> AuditResult:
    """
    Запуск аудита план-графика.

    Args:
        args: Namespace с полями target, session_dir, parse_only, rule_filter

    Returns:
        AuditResult с нарушениями
    """
    start_time = time.time()
    target_path = args.target
    session_dir = Path(args.session_dir) if args.session_dir else Path(f"logs_result/plan_grafik/session_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    session_dir.mkdir(parents=True, exist_ok=True)

    print(f"[plan_grafik] Старт аудита: {target_path}")
    print(f"[plan_grafik] Сессия: {session_dir}")

    # Шаг 1: Парсинг
    try:
        from .parser import parse_plan_grafik
        parsed = parse_plan_grafik(target_path)
    except Exception as e:
        error_msg = f"Ошибка парсинга: {e}"
        print(f"[plan_grafik] {error_msg}", file=sys.stderr)
        (session_dir / "ERROR.txt").write_text(error_msg, encoding="utf-8")
        return AuditResult(
            violations=[],
            doc_type="plan_grafik",
            session_dir=session_dir,
            duration_sec=time.time() - start_time,
            rules_checked=0,
            target_path=target_path,
        )

    # Сохраняем парсинг
    parsed_path = session_dir / "parsed.json"
    with open(parsed_path, "w", encoding="utf-8") as f:
        json.dump(parsed, f, ensure_ascii=False, indent=2, default=str)
    print(f"[plan_grafik] Парсинг сохранён: {parsed_path}")

    if getattr(args, "parse_only", False):
        return AuditResult(
            doc_type="plan_grafik", session_dir=session_dir,
            duration_sec=time.time() - start_time, target_path=target_path,
        )

    # Шаг 2: Валидация
    from .validators import run_all_validators
    violations = run_all_validators(parsed, target_path)

    # Шаг 3: Excel-отчёт
    from ..excel_reporter import save_to_excel
    xlsx_path = str(session_dir / "validation_report.xlsx")
    save_to_excel(violations, xlsx_path)
    print(f"[plan_grafik] Отчёт: {xlsx_path}")

    duration = time.time() - start_time
    print(f"[plan_grafik] Завершён за {duration:.1f} сек. Нарушений: {len(violations)}")

    return AuditResult(
        violations=violations,
        doc_type="plan_grafik",
        session_dir=session_dir,
        duration_sec=duration,
        rules_checked=9,
        target_path=target_path,
    )
