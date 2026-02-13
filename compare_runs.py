#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт сравнения двух тестовых прогонов.

Сравнивает rules_status.json между прогонами, определяет:
- ИСПРАВЛЕНО (FAIL → PASS) — ожидаемые фиксы
- РЕГРЕССИИ (PASS → FAIL) — новые проблемы
- БЕЗ ИЗМЕНЕНИЙ — стабильные результаты
- НОВЫЕ — правила, которых не было в старом прогоне

Использует test_analysis.json как baseline для определения ожидаемых false_fail.

Использование:
    python compare_runs.py --old run_20260210_165639 --new run_20260212_143500
    python compare_runs.py --old run_20260210_165639 --new run_20260212_143500 --doc-type kpsc
    python compare_runs.py --old run_20260210_165639 --new run_20260212_143500 --doc-type kpsc --company biznes_otel
"""

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

BASE_DIR = Path(__file__).parent
RESULTS_DIR = BASE_DIR / "test_results"
ANALYSIS_PATH = BASE_DIR / "test_configs" / "test_analysis.json"


def load_analysis_baseline() -> Dict[str, Dict]:
    """
    Загружает test_analysis.json как baseline.

    Возвращает словарь: {(doc_type, company, rule): entry}
    """
    if not ANALYSIS_PATH.exists():
        return {}

    with open(ANALYSIS_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)

    baseline = {}
    for entry in data.get("entries", []):
        key = (entry["doc_type"], entry["company"], str(entry["rule"]))
        baseline[key] = entry
    return baseline


def load_run_statuses(run_dir: Path, doc_type_filter: Optional[str] = None,
                      company_filter: Optional[str] = None) -> Dict[str, Dict[str, str]]:
    """
    Загружает rules_status.json для всех doc_type/company в прогоне.

    Возвращает: {(doc_type, company): {rule: status}}
    """
    statuses = {}

    if not run_dir.exists():
        print(f"Ошибка: директория {run_dir} не найдена", file=sys.stderr)
        sys.exit(1)

    # Перебираем doc_type/company/rules_status.json
    for doc_type_dir in sorted(run_dir.iterdir()):
        if not doc_type_dir.is_dir():
            continue
        doc_type = doc_type_dir.name

        # Пропускаем не-doc_type файлы (summary.json, summary.xlsx и т.д.)
        if doc_type.startswith("summary") or doc_type.startswith("."):
            continue

        if doc_type_filter and doc_type != doc_type_filter:
            continue

        for company_dir in sorted(doc_type_dir.iterdir()):
            if not company_dir.is_dir():
                continue
            company = company_dir.name

            if company_filter and company != company_filter:
                continue

            rules_status_path = company_dir / "rules_status.json"
            if not rules_status_path.exists():
                continue

            with open(rules_status_path, "r", encoding="utf-8") as f:
                rules = json.load(f)

            statuses[(doc_type, company)] = {str(k): v for k, v in rules.items()}

    return statuses


def compare(old_statuses: Dict, new_statuses: Dict,
            baseline: Dict) -> Dict[str, List[Dict]]:
    """
    Сравнивает два набора статусов.

    Возвращает словарь категорий:
    - fixed: FAIL → PASS
    - regression: PASS → FAIL
    - stable_fail: FAIL → FAIL
    - stable_pass: PASS → PASS
    - new_fail: правило не было в старом прогоне, теперь FAIL
    - new_pass: правило не было в старом прогоне, теперь PASS
    - disappeared: было в старом, нет в новом
    """
    results = {
        "fixed": [],
        "regression": [],
        "stable_fail": [],
        "stable_pass": [],
        "new_fail": [],
        "new_pass": [],
        "disappeared": [],
    }

    # Собираем все уникальные ключи (doc_type, company) из обоих прогонов
    all_keys = set(old_statuses.keys()) | set(new_statuses.keys())

    for key in sorted(all_keys):
        doc_type, company = key
        old_rules = old_statuses.get(key, {})
        new_rules = new_statuses.get(key, {})

        all_rules = set(old_rules.keys()) | set(new_rules.keys())

        for rule in sorted(all_rules, key=lambda r: _rule_sort_key(r)):
            old_status = old_rules.get(rule)
            new_status = new_rules.get(rule)

            # Ищем запись в baseline
            baseline_key = (doc_type, company, rule)
            baseline_entry = baseline.get(baseline_key, {})
            verdict = baseline_entry.get("verdict", "")
            root_cause = baseline_entry.get("root_cause", "")
            entry_id = baseline_entry.get("id", "")

            record = {
                "doc_type": doc_type,
                "company": company,
                "rule": rule,
                "old_status": old_status,
                "new_status": new_status,
                "baseline_verdict": verdict,
                "baseline_root_cause": root_cause,
                "baseline_id": entry_id,
            }

            if old_status is None and new_status is not None:
                # Новое правило в прогоне
                if new_status == "FAIL":
                    results["new_fail"].append(record)
                else:
                    results["new_pass"].append(record)
            elif old_status is not None and new_status is None:
                # Правило исчезло
                results["disappeared"].append(record)
            elif old_status == "FAIL" and new_status == "PASS":
                results["fixed"].append(record)
            elif old_status == "PASS" and new_status == "FAIL":
                results["regression"].append(record)
            elif old_status == "FAIL" and new_status == "FAIL":
                results["stable_fail"].append(record)
            elif old_status == "PASS" and new_status == "PASS":
                results["stable_pass"].append(record)

    return results


def _rule_sort_key(rule: str) -> Tuple:
    """Сортировка правил: '1' < '2' < '1.1' < '1.2' < '10'."""
    parts = rule.split(".")
    return tuple(int(p) if p.isdigit() else p for p in parts)


def print_report(results: Dict[str, List[Dict]], old_name: str, new_name: str,
                 doc_type_filter: Optional[str] = None):
    """Печатает отчёт сравнения в консоль."""

    title = "СРАВНЕНИЕ ПРОГОНОВ"
    if doc_type_filter:
        title += f": {doc_type_filter}"

    print(f"\n{'=' * 70}")
    print(f"  {title}")
    print(f"  OLD: {old_name}")
    print(f"  NEW: {new_name}")
    print(f"{'=' * 70}\n")

    # --- ИСПРАВЛЕНО ---
    fixed = results["fixed"]
    expected_fixes = [r for r in fixed if r["baseline_verdict"] == "false_fail"]
    unexpected_fixes = [r for r in fixed if r["baseline_verdict"] == "true_fail"]
    other_fixes = [r for r in fixed if r["baseline_verdict"] not in ("true_fail", "false_fail")]

    print(f"✅ ИСПРАВЛЕНО (FAIL → PASS): {len(fixed)}")
    if expected_fixes:
        print(f"   Ожидаемые фиксы (false_fail → PASS): {len(expected_fixes)}")
        for r in expected_fixes:
            print(f"     {r['doc_type']}/{r['company']}  rule {r['rule']}"
                  f"  (was false_fail {r['baseline_root_cause']}, entry #{r['baseline_id']})")

    if unexpected_fixes:
        print(f"\n   ⚠️  ВНИМАНИЕ: true_fail стали PASS: {len(unexpected_fixes)}")
        print(f"   Это может означать регрессию — реальные ошибки маскируются!")
        for r in unexpected_fixes:
            print(f"     {r['doc_type']}/{r['company']}  rule {r['rule']}"
                  f"  (was true_fail B, entry #{r['baseline_id']}) ← РАССЛЕДОВАТЬ!")

    if other_fixes:
        print(f"\n   Без baseline записи: {len(other_fixes)}")
        for r in other_fixes:
            print(f"     {r['doc_type']}/{r['company']}  rule {r['rule']}")

    # --- РЕГРЕССИИ ---
    regressions = results["regression"]
    print(f"\n{'❌' if regressions else '✅'} РЕГРЕССИИ (PASS → FAIL): {len(regressions)}")
    for r in regressions:
        print(f"     {r['doc_type']}/{r['company']}  rule {r['rule']}  ← НОВЫЙ FAIL!")

    # --- НОВЫЕ FAIL ---
    new_fails = results["new_fail"]
    if new_fails:
        print(f"\n⚠️  НОВЫЕ FAIL (не было в старом прогоне): {len(new_fails)}")
        for r in new_fails:
            print(f"     {r['doc_type']}/{r['company']}  rule {r['rule']}")

    # --- СТАБИЛЬНЫЕ ---
    stable_fail = results["stable_fail"]
    stable_pass = results["stable_pass"]
    print(f"\n📊 БЕЗ ИЗМЕНЕНИЙ:")
    print(f"   FAIL → FAIL: {len(stable_fail)}")

    # Показываем false_fail, которые НЕ исправились
    unfixed = [r for r in stable_fail if r["baseline_verdict"] == "false_fail"]
    if unfixed:
        print(f"   ⚠️  Не исправленные false_fail: {len(unfixed)}")
        for r in unfixed:
            print(f"     {r['doc_type']}/{r['company']}  rule {r['rule']}"
                  f"  ({r['baseline_root_cause']}, entry #{r['baseline_id']})")

    print(f"   PASS → PASS: {len(stable_pass)}")

    # --- ИСЧЕЗНУВШИЕ ---
    disappeared = results["disappeared"]
    if disappeared:
        print(f"\n⚠️  ИСЧЕЗЛИ (были в старом, нет в новом): {len(disappeared)}")
        for r in disappeared:
            print(f"     {r['doc_type']}/{r['company']}  rule {r['rule']}")

    # --- ИТОГО ---
    print(f"\n{'─' * 70}")
    total_old_fail = len(fixed) + len(stable_fail)
    total_new_fail = len(regressions) + len(stable_fail) + len(new_fails)
    print(f"  ИТОГО: {len(fixed)} fix, {len(regressions)} regression, "
          f"{len(stable_fail)} stable_fail, {len(new_fails)} new_fail")
    print(f"  FAIL в старом: {total_old_fail} → FAIL в новом: {total_new_fail}")

    if total_old_fail > 0:
        improvement = total_old_fail - total_new_fail
        print(f"  Улучшение: {'+' if improvement >= 0 else ''}{improvement} "
              f"({'меньше' if improvement > 0 else 'больше'} FAIL)")

    # Статус
    if regressions:
        print(f"\n  🔴 ЕСТЬ РЕГРЕССИИ — нужно расследование!")
    elif unexpected_fixes:
        print(f"\n  🟡 true_fail стали PASS — проверить, не маскируем ли ошибки")
    elif unfixed:
        print(f"\n  🟡 Часть false_fail не исправлена ({len(unfixed)} из {len(unfixed) + len(expected_fixes)})")
    else:
        print(f"\n  🟢 Все фиксы сработали, регрессий нет!")

    print()


def save_report_json(results: Dict[str, List[Dict]], output_path: Path):
    """Сохраняет отчёт в JSON для дальнейшей обработки."""
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"Отчёт сохранён: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Сравнение двух тестовых прогонов аудита"
    )
    parser.add_argument("--old", required=True,
                        help="Имя старого прогона (например run_20260210_165639)")
    parser.add_argument("--new", required=True,
                        help="Имя нового прогона (или 'latest')")
    parser.add_argument("--doc-type", default=None,
                        help="Фильтр по doc_type")
    parser.add_argument("--company", default=None,
                        help="Фильтр по компании")
    parser.add_argument("--json", default=None,
                        help="Сохранить отчёт в JSON")

    args = parser.parse_args()

    # Определяем директории прогонов
    old_dir = RESULTS_DIR / args.old
    if args.new == "latest":
        new_dir = RESULTS_DIR / "latest"
        # Резолвим симлинк для имени
        new_name = new_dir.resolve().name if new_dir.is_symlink() else "latest"
    else:
        new_dir = RESULTS_DIR / args.new
        new_name = args.new

    if not old_dir.exists():
        print(f"Ошибка: старый прогон не найден: {old_dir}", file=sys.stderr)
        sys.exit(1)
    if not new_dir.exists():
        print(f"Ошибка: новый прогон не найден: {new_dir}", file=sys.stderr)
        sys.exit(1)

    # Загружаем данные
    baseline = load_analysis_baseline()
    old_statuses = load_run_statuses(old_dir, args.doc_type, args.company)
    new_statuses = load_run_statuses(new_dir, args.doc_type, args.company)

    if not old_statuses:
        print(f"Ошибка: нет данных в старом прогоне для заданных фильтров", file=sys.stderr)
        sys.exit(1)
    if not new_statuses:
        print(f"Ошибка: нет данных в новом прогоне для заданных фильтров", file=sys.stderr)
        sys.exit(1)

    # Сравниваем
    results = compare(old_statuses, new_statuses, baseline)

    # Выводим отчёт
    print_report(results, args.old, new_name, args.doc_type)

    # Сохраняем JSON если запрошено
    if args.json:
        save_report_json(results, Path(args.json))


if __name__ == "__main__":
    main()
