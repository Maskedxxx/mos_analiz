#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль аудита драйверов производительности.

Точка входа: run(args) — вызывается из run_audit.py для doc_type="drivers".
Пайплайн: парсинг Excel → LLM-анализ секций → сбор замечаний → Excel-отчёт.
"""

import json
import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from openai import OpenAI

from audit_engine.models import AuditResult
from .parser import parse_excel_to_json
from .analyzer import (
    analyze_sections,
    collect_remarks_and_summaries,
    export_missing_driver_report,
    load_driver_prompt,
)


def _load_config() -> Dict[str, Any]:
    """Загрузка конфига из doc_configs/drivers/config.json."""
    config_path = Path(__file__).resolve().parents[2] / "doc_configs" / "drivers" / "config.json"
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def run(args) -> AuditResult:
    """
    Запуск аудита драйверов.

    Вход: args (argparse Namespace) с полями:
        - target: путь к XLSX файлу
        - parse_only: только парсинг (опционально)
        - model: переопределить модель (опционально)
        - temperature: переопределить температуру (опционально)
        - session_dir: директория для логов (опционально)
    Выход: AuditResult с violations
    """
    start_time = time.time()
    target_path = Path(args.target)

    # Загрузка конфига
    config = _load_config()
    model = args.model or config.get("model", "gpt-4.1-mini")
    temperature = args.temperature if args.temperature is not None else config.get("temperature", 0.0)
    primary_threshold = config.get("primary_threshold", 9.0)
    fallback_threshold = config.get("fallback_threshold", 7.0)

    # Создание сессии логирования
    if args.session_dir:
        session_dir = Path(args.session_dir)
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        session_dir = Path(__file__).resolve().parents[2] / "logs_result" / "drivers" / f"session_{timestamp}"
    session_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"  Аудит драйверов производительности")
    print(f"{'='*60}")
    print(f"  Файл: {target_path.name}")
    print(f"  Модель: {model}")
    print(f"  Пороги: primary={primary_threshold}, fallback={fallback_threshold}")
    print(f"  Сессия: {session_dir}")
    print()

    # === Шаг 1: Парсинг Excel ===
    print("[1/4] Парсинг Excel...")
    parsed = parse_excel_to_json(target_path)

    # Сохраняем результат парсинга
    parsed_path = session_dir / "01_parsed.json"
    with open(parsed_path, "w", encoding="utf-8") as f:
        json.dump(parsed, f, ensure_ascii=False, indent=2)
    print(f"  Секций: {len(parsed.get('sections', []))}")
    print(f"  Сохранено: {parsed_path}")

    # === Режим --parse-only ===
    if args.parse_only:
        print(f"\n{'='*60}")
        print("TARGET (parsed):")
        print(json.dumps(parsed, ensure_ascii=False, indent=2))
        return AuditResult(
            doc_type="drivers",
            session_dir=session_dir,
            target_path=str(target_path),
            duration_sec=time.time() - start_time,
        )

    # === Шаг 2: LLM-анализ секций ===
    print("\n[2/4] LLM-анализ секций...")
    api_key = os.environ.get("OPENAI_API_KEY", "dummy")
    base_url = config.get("llm_base_url", "http://localhost:8001/v1/")

    client = OpenAI(api_key=api_key, base_url=base_url)

    # Загрузка промпта
    system_prompt = load_driver_prompt()

    def progress_cb(msg: str, idx: int, total: int):
        print(f"  [{idx}/{total}] {msg}")

    section_results = analyze_sections(
        parsed_data=parsed,
        client=client,
        primary_threshold=primary_threshold,
        fallback_threshold=fallback_threshold,
        model=model,
        temperature=temperature,
        progress_callback=progress_cb,
        system_prompt=system_prompt,
    )

    # Сохраняем результат LLM-анализа
    analysis_path = session_dir / "02_llm_analysis.json"
    with open(analysis_path, "w", encoding="utf-8") as f:
        json.dump(section_results, f, ensure_ascii=False, indent=2)
    print(f"  Сохранено: {analysis_path}")

    # === Шаг 3: Сбор замечаний ===
    print("\n[3/4] Сбор замечаний...")
    remarks, driver_summary_map = collect_remarks_and_summaries(section_results)
    print(f"  Замечаний: {len(remarks)}")

    # === Шаг 4: Excel-отчёт ===
    print("\n[4/4] Генерация Excel-отчёта...")
    report_path = session_dir / "03_report.xlsx"
    export_missing_driver_report(parsed, section_results, report_path)
    print(f"  Отчёт: {report_path}")

    # === Итог ===
    duration = time.time() - start_time
    violations = [
        {
            "rule_index": f"driver_{r.get('section_title', 'x')}_{r.get('number', '?')}",
            "rule_title": r.get("driver_name", ""),
            "section_title": r.get("section_title", ""),
            "issue": r.get("issue", ""),
        }
        for r in remarks
    ]

    print(f"\n{'='*60}")
    print(f"  Итого нарушений: {len(violations)}")
    if violations:
        for v in violations:
            print(f"  - [{v['rule_index']}] {v['rule_title']}: {v['issue']}")
    print(f"  Время: {duration:.1f} сек")
    print(f"  Сессия: {session_dir}")
    print(f"{'='*60}")

    return AuditResult(
        violations=violations,
        doc_type="drivers",
        session_dir=session_dir,
        duration_sec=duration,
        rules_checked=len(parsed.get("sections", [])),
        target_path=str(target_path),
    )
