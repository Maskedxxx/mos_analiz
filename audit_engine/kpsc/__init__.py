#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модуль валидации КПСЦ (Карта Потока Создания Ценности).

Точка входа: run(args) — вызывается из run_audit.py для doc_type="kpsc".
Пайплайн: 9 парсеров → 25 валидаторов (параллельно) → Excel-отчёт.
"""

import json
import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

from audit_engine.models import AuditResult
from .run_validations import (
    load_rules,
    run_parsers,
    run_validators_parallel,
    create_excel_report,
)


def _load_config() -> Dict[str, Any]:
    """Загрузка конфига из doc_configs/kpsc/config.json."""
    config_path = Path(__file__).resolve().parents[2] / "doc_configs" / "kpsc" / "config.json"
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def run(args) -> AuditResult:
    """
    Запуск валидации КПСЦ.

    Вход: args (argparse Namespace) с полями:
        - target: путь к XLSX файлу
        - parse_only: только парсинг (9 парсеров)
        - rule_filter: проверить только указанное правило (например '1.1')
        - model: переопределить модель
        - temperature: переопределить температуру
        - session_dir: директория для логов
    Выход: AuditResult с violations
    """
    start_time = time.time()
    target_path = Path(args.target)

    # Загрузка конфига
    config = _load_config()
    max_workers = config.get("max_workers", 5)

    # Пути к правилам
    rules_path = Path(__file__).resolve().parents[2] / "doc_configs" / "kpsc" / "validation_rules.json"

    # Создание сессии логирования
    if args.session_dir:
        session_dir = Path(args.session_dir)
    else:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        session_dir = Path(__file__).resolve().parents[2] / "logs_result" / "kpsc" / f"session_{timestamp}"
    session_dir.mkdir(parents=True, exist_ok=True)

    parser_outputs_dir = session_dir / "parser_outputs"
    validation_outputs_dir = session_dir / "validation_outputs"

    print(f"\n{'='*60}")
    print(f"  Валидация КПСЦ")
    print(f"{'='*60}")
    print(f"  Файл: {target_path.name}")
    print(f"  Сессия: {session_dir}")
    print(f"  Параллельность: {max_workers}")
    print()

    # Установка переменных среды для валидаторов (subprocess)
    os.environ["VALIDATION_RULES_PATH"] = str(rules_path)
    os.environ["LLM_BASE_URL"] = config.get("llm_base_url", "http://localhost:8001/v1/")
    os.environ["LLM_MODEL"] = config.get("model", "openai/gpt-oss-120b")
    if not os.environ.get("OPENAI_API_KEY"):
        os.environ["OPENAI_API_KEY"] = "dummy"

    # === Шаг 1: Парсинг (9 парсеров) ===
    print("[1/3] Запуск 9 парсеров...")
    parser_results = run_parsers(target_path, parser_outputs_dir)

    ok_count = sum(1 for r in parser_results.values() if r["status"] == "ok")
    err_count = sum(1 for r in parser_results.values() if r["status"] == "error")
    print(f"\n  Парсинг завершён: {ok_count} OK, {err_count} ошибок")

    # Сохраняем сводку парсинга
    with open(session_dir / "parser_summary.json", "w", encoding="utf-8") as f:
        json.dump(parser_results, f, ensure_ascii=False, indent=2)

    # === Режим --parse-only ===
    if args.parse_only:
        print(f"\n{'='*60}")
        print("Режим --parse-only: парсинг завершён")
        print(f"Результаты в: {parser_outputs_dir}")

        # Выводим имена сгенерированных файлов
        for p in sorted(parser_outputs_dir.glob("*.json")):
            print(f"  {p.name}")

        return AuditResult(
            doc_type="kpsc",
            session_dir=session_dir,
            target_path=str(target_path),
            duration_sec=time.time() - start_time,
        )

    # === Шаг 2: Валидация (25 правил) ===
    print("\n[2/3] Загрузка правил...")
    rules = load_rules(rules_path)

    # Фильтрация по правилу
    if args.rule_filter:
        rule_filter_str = str(args.rule_filter)
        rules = [r for r in rules if r["rule_index"] == rule_filter_str]
        if not rules:
            raise ValueError(f"Правило {rule_filter_str} не найдено в kpsc")
        print(f"  Фильтр: только правило {rule_filter_str}")

    print(f"  Правил к проверке: {len(rules)}")
    print()

    results = run_validators_parallel(
        rules=rules,
        parser_outputs_dir=parser_outputs_dir,
        output_dir=validation_outputs_dir,
        rules_path=rules_path,
        max_workers=max_workers,
    )

    # === Шаг 3: Отчёт ===
    print(f"\n[3/3] Генерация отчёта...")
    report_path = session_dir / "validation_report.xlsx"
    df = create_excel_report(results, report_path)
    print(f"  Отчёт: {report_path}")

    # Статистика
    total = len(df)
    passed = len(df[df["status"] == "PASS"])
    failed = len(df[df["status"] == "FAIL"])
    errors = total - passed - failed

    duration = time.time() - start_time

    print(f"\n{'='*60}")
    print(f"  Статистика:")
    print(f"    Всего правил: {total}")
    print(f"    PASS: {passed} ({passed/total*100:.1f}%)" if total > 0 else "")
    print(f"    FAIL: {failed} ({failed/total*100:.1f}%)" if total > 0 else "")
    if errors > 0:
        print(f"    ERRORS: {errors} ({errors/total*100:.1f}%)")
    print(f"  Время: {duration:.1f} сек")
    print(f"  Сессия: {session_dir}")
    print(f"{'='*60}")

    # Собираем violations (FAIL → violation)
    violations = []
    for rule, result_data, dur in results:
        if result_data.get("status") == "FAIL":
            violations.append({
                "rule_index": result_data.get("rule_index", rule.get("rule_index", "?")),
                "rule_title": result_data.get("rule_title", rule.get("rule_title", "")),
                "section": rule.get("section", ""),
                "discrepancy": result_data.get("discrepancy", ""),
            })

    if violations:
        print(f"\n  Нарушения ({len(violations)}):")
        for v in violations:
            print(f"  - [{v['rule_index']}] {v['rule_title']}: {v['discrepancy'][:80]}")

    return AuditResult(
        violations=violations,
        doc_type="kpsc",
        session_dir=session_dir,
        duration_sec=duration,
        rules_checked=total,
        target_path=str(target_path),
    )
