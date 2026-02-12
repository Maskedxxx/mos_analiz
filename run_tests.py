#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Интеграционный тест-раннер для всех типов документов.

Прогоняет тестовые файлы из test_docs/ через run_audit.py,
собирает результаты валидации и генерирует сводные отчёты.

Использование:
    python run_tests.py                              # все типы
    python run_tests.py --doc-type kartochka_proekta  # один тип
    python run_tests.py --doc-type kpsc --company mapper  # один файл
    python run_tests.py --report-only                 # отчёт из последнего прогона
    python run_tests.py --timeout 600                 # таймаут 10 мин на файл
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

# Корень проекта — папка, где лежит этот скрипт
BASE_DIR = Path(__file__).resolve().parent

# Таймаут по умолчанию (секунды)
DEFAULT_TIMEOUT = 300


def _load_dotenv():
    """Загружает переменные из .env файла (если есть) в os.environ."""
    env_path = BASE_DIR / ".env"
    if not env_path.exists():
        return
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value


# ─────────────────────────────────────────────────
# Загрузка тестовой матрицы
# ─────────────────────────────────────────────────

def load_test_matrix(matrix_path: Path) -> List[Dict[str, Any]]:
    """Загружает test_matrix.json и возвращает список test_suites."""
    with open(matrix_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data["test_suites"]


def filter_suites(
    suites: List[Dict], doc_type: Optional[str], company: Optional[str]
) -> List[Dict]:
    """Фильтрация по --doc-type и --company."""
    if doc_type:
        suites = [s for s in suites if s["doc_type"] == doc_type]

    if company:
        filtered = []
        for suite in suites:
            matching_files = [
                f for f in suite["files"] if f.get("company") == company
            ]
            if matching_files:
                suite_copy = dict(suite)
                suite_copy["files"] = matching_files
                filtered.append(suite_copy)
        suites = filtered

    return suites


# ─────────────────────────────────────────────────
# Проверка зависимостей
# ─────────────────────────────────────────────────

def check_dependencies(suites: List[Dict]) -> Dict[str, bool]:
    """
    Проверяет наличие OPENAI_API_KEY и LibreOffice.
    Возвращает словарь {dependency: available}.
    """
    deps = {}

    # OPENAI_API_KEY — нужен всем движкам
    deps["OPENAI_API_KEY"] = bool(os.environ.get("OPENAI_API_KEY"))

    # LibreOffice — нужен только vision-движку
    has_vision = any(s["engine"] == "vision" for s in suites)
    if has_vision:
        libre_available = shutil.which("libreoffice") or shutil.which("soffice")
        deps["LibreOffice"] = bool(libre_available)
    else:
        deps["LibreOffice"] = True  # не требуется

    return deps


# ─────────────────────────────────────────────────
# Запуск одного тест-кейса
# ─────────────────────────────────────────────────

def run_single_test(
    doc_type: str,
    engine: str,
    file_info: Dict[str, Any],
    session_dir: Path,
    timeout: int,
    multi_file: bool = False,
) -> Dict[str, Any]:
    """
    Запускает run_audit.py для одного файла и возвращает результат.

    Формат результата:
    {
        "doc_type": str,
        "company": str,
        "engine": str,
        "status": "PASS" | "FAIL" | "ERROR",
        "exit_code": int,
        "duration_sec": float,
        "violations_count": int,
        "rules_status": {rule: status},
        "error": str | None,
        "session_dir": str,
        "target_path": str,
    }
    """
    company = file_info["company"]
    session_dir.mkdir(parents=True, exist_ok=True)

    # Определяем путь к файлу
    if multi_file:
        target_path = str(BASE_DIR / file_info["primary"])
        secondary_path = str(BASE_DIR / file_info["secondary"])
    else:
        target_path = str(BASE_DIR / file_info["path"])
        secondary_path = None

    # Собираем команду
    cmd = [
        sys.executable, str(BASE_DIR / "run_audit.py"),
        "--doc-type", doc_type,
        "--target", target_path,
        "--session-dir", str(session_dir),
    ]
    if secondary_path:
        cmd.extend(["--secondary", secondary_path])

    result = {
        "doc_type": doc_type,
        "company": company,
        "engine": engine,
        "target_path": target_path,
        "session_dir": str(session_dir),
        "status": "ERROR",
        "exit_code": -1,
        "duration_sec": 0.0,
        "violations_count": 0,
        "rules_status": {},
        "error": None,
    }

    start = time.time()

    try:
        # Запуск subprocess с таймаутом
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            cwd=str(BASE_DIR),
        )
        result["exit_code"] = proc.returncode
        result["duration_sec"] = round(time.time() - start, 1)

        # Сохраняем stdout/stderr для отладки
        if proc.stdout:
            (session_dir / "test_stdout.txt").write_text(proc.stdout, encoding="utf-8")
        if proc.stderr:
            (session_dir / "test_stderr.txt").write_text(proc.stderr, encoding="utf-8")

        # Парсим результаты в зависимости от движка
        if engine in ("kpsc", "kartochka_proekta"):
            rules_status = _parse_special_engine_results(session_dir)
        else:
            rules_status = _parse_vision_results(session_dir)

        result["rules_status"] = rules_status

        # Определяем общий статус
        if proc.returncode not in (0, 1):
            # Ненормальный код — ошибка
            result["status"] = "ERROR"
            result["error"] = f"exit code {proc.returncode}"
            if proc.stderr:
                result["error"] += f"\n{proc.stderr[-500:]}"
        elif not rules_status:
            # Нет результатов валидации — возможно ошибка парсинга
            result["status"] = "ERROR"
            result["error"] = "Нет результатов валидации в session_dir"
        else:
            # Считаем violations
            fails = [r for r, s in rules_status.items() if s == "FAIL"]
            errors = [r for r, s in rules_status.items() if s == "ERROR"]
            result["violations_count"] = len(fails)

            if errors:
                result["status"] = "ERROR"
                result["error"] = f"{len(errors)} правил с ошибками"
            elif fails:
                result["status"] = "FAIL"
            else:
                result["status"] = "PASS"

    except subprocess.TimeoutExpired:
        result["duration_sec"] = round(time.time() - start, 1)
        result["status"] = "ERROR"
        result["error"] = f"Таймаут ({timeout} сек)"
    except Exception as e:
        result["duration_sec"] = round(time.time() - start, 1)
        result["status"] = "ERROR"
        result["error"] = f"{type(e).__name__}: {e}\n{traceback.format_exc()}"

    return result


# ─────────────────────────────────────────────────
# Парсинг результатов из session_dir
# ─────────────────────────────────────────────────

def _parse_special_engine_results(session_dir: Path) -> Dict[str, str]:
    """
    Парсит результаты kpsc / kartochka_proekta.
    Читает validation_outputs/validate_*.json → {rule_index: status}.
    """
    rules_status = {}
    val_dir = session_dir / "validation_outputs"

    if not val_dir.exists():
        return rules_status

    for f in sorted(val_dir.glob("validate_*.json")):
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
            rule_idx = str(data.get("rule_index", f.stem))
            status = data.get("status", "ERROR")
            rules_status[rule_idx] = status
        except (json.JSONDecodeError, KeyError):
            rules_status[f.stem] = "ERROR"

    return rules_status


def _parse_vision_results(session_dir: Path) -> Dict[str, str]:
    """
    Парсит результаты vision-движка.
    Читает rules_summary.json (все правила) и final_results.json (violations).
    Правила, не вошедшие в violations → PASS.
    """
    rules_status = {}

    # Загружаем список всех правил
    rules_path = session_dir / "rules_summary.json"
    if rules_path.exists():
        try:
            rules = json.loads(rules_path.read_text(encoding="utf-8"))
            for r in rules:
                idx = str(r.get("index", "?"))
                rules_status[idx] = "PASS"  # по умолчанию PASS
        except (json.JSONDecodeError, KeyError):
            pass

    # Загружаем нарушения — они перекрывают PASS → FAIL
    violations_path = session_dir / "final_results.json"
    if violations_path.exists():
        try:
            violations = json.loads(violations_path.read_text(encoding="utf-8"))
            for v in violations:
                idx = str(v.get("rule_index", "?"))
                rules_status[idx] = "FAIL"
        except (json.JSONDecodeError, KeyError):
            pass

    return rules_status


# ─────────────────────────────────────────────────
# Генерация отчётов
# ─────────────────────────────────────────────────

def generate_summary(results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Генерирует сводный JSON-отчёт."""
    total = len(results)
    passed = sum(1 for r in results if r["status"] == "PASS")
    failed = sum(1 for r in results if r["status"] == "FAIL")
    errors = sum(1 for r in results if r["status"] == "ERROR")

    # Агрегация по типам
    by_type = {}
    for r in results:
        dt = r["doc_type"]
        if dt not in by_type:
            by_type[dt] = {
                "doc_type": dt,
                "engine": r["engine"],
                "files": 0,
                "pass": 0,
                "fail": 0,
                "error": 0,
                "total_rules": 0,
                "rules_pass": 0,
                "rules_fail": 0,
                "rules_error": 0,
                "duration_sec": 0.0,
            }
        entry = by_type[dt]
        entry["files"] += 1
        entry[r["status"].lower()] += 1
        entry["duration_sec"] += r["duration_sec"]

        # Считаем правила
        for _, status in r["rules_status"].items():
            entry["total_rules"] += 1
            if status == "PASS":
                entry["rules_pass"] += 1
            elif status == "FAIL":
                entry["rules_fail"] += 1
            else:
                entry["rules_error"] += 1

    return {
        "timestamp": datetime.now().isoformat(),
        "total_files": total,
        "passed": passed,
        "failed": failed,
        "errors": errors,
        "by_type": list(by_type.values()),
        "results": results,
    }


def generate_excel_report(summary: Dict[str, Any], output_path: Path):
    """
    Генерирует Excel-отчёт с 3 листами:
    1. Сводка по типам
    2. Детально по файлам
    3. Ошибки
    """
    try:
        import pandas as pd
    except ImportError:
        print("  ВНИМАНИЕ: pandas не установлен, Excel-отчёт пропущен")
        return

    # Лист 1: Сводка по типам
    type_rows = []
    for t in summary["by_type"]:
        pct = (
            round(t["rules_pass"] / t["total_rules"] * 100, 1)
            if t["total_rules"] > 0 else 0
        )
        type_rows.append({
            "doc_type": t["doc_type"],
            "engine": t["engine"],
            "файлов": t["files"],
            "правил_всего": t["total_rules"],
            "PASS": t["rules_pass"],
            "FAIL": t["rules_fail"],
            "ERROR": t["rules_error"],
            "% PASS": pct,
            "время_сек": round(t["duration_sec"], 1),
        })
    df_types = pd.DataFrame(type_rows)

    # Лист 2: Детально по файлам (каждое правило)
    detail_rows = []
    for r in summary["results"]:
        for rule, status in r["rules_status"].items():
            # Пытаемся найти discrepancy для FAIL-правил
            discrepancy = ""
            if status == "FAIL":
                discrepancy = _find_discrepancy(r, rule)

            detail_rows.append({
                "doc_type": r["doc_type"],
                "company": r["company"],
                "файл": Path(r["target_path"]).name,
                "rule": rule,
                "status": status,
                "discrepancy": discrepancy,
            })
    df_details = pd.DataFrame(detail_rows) if detail_rows else pd.DataFrame()

    # Лист 3: Ошибки
    error_rows = []
    for r in summary["results"]:
        if r["status"] == "ERROR" and r.get("error"):
            error_rows.append({
                "doc_type": r["doc_type"],
                "company": r["company"],
                "файл": Path(r["target_path"]).name,
                "error_type": r["error"].split("\n")[0][:100],
                "traceback": r["error"][:1000],
            })
    df_errors = pd.DataFrame(error_rows) if error_rows else pd.DataFrame()

    # Запись в Excel
    with pd.ExcelWriter(str(output_path), engine="openpyxl") as writer:
        df_types.to_excel(writer, sheet_name="Сводка по типам", index=False)
        if not df_details.empty:
            df_details.to_excel(writer, sheet_name="Детально по файлам", index=False)
        if not df_errors.empty:
            df_errors.to_excel(writer, sheet_name="Ошибки", index=False)

    print(f"  Excel-отчёт: {output_path}")


def _find_discrepancy(result: Dict, rule_idx: str) -> str:
    """
    Ищет текст расхождения для FAIL-правила.
    Для kpsc/kartochka — читает validate_N.json.
    Для vision — читает final_results.json.
    """
    session_dir = Path(result["session_dir"])
    engine = result["engine"]

    if engine in ("kpsc", "kartochka_proekta"):
        # Ищем validate_N.json (N может быть 1_2 для kpsc)
        val_dir = session_dir / "validation_outputs"
        prefix = rule_idx.replace(".", "_")
        candidate = val_dir / f"validate_{prefix}.json"
        if candidate.exists():
            try:
                data = json.loads(candidate.read_text(encoding="utf-8"))
                return data.get("discrepancy", "")
            except (json.JSONDecodeError, KeyError):
                pass
    else:
        # Vision — ищем в final_results.json
        fr_path = session_dir / "final_results.json"
        if fr_path.exists():
            try:
                violations = json.loads(fr_path.read_text(encoding="utf-8"))
                for v in violations:
                    if str(v.get("rule_index")) == rule_idx:
                        return v.get("discrepancy", "")
            except (json.JSONDecodeError, KeyError):
                pass

    return ""


# ─────────────────────────────────────────────────
# Консольный вывод
# ─────────────────────────────────────────────────

def print_summary_table(summary: Dict[str, Any]):
    """Выводит красивую таблицу результатов в консоль."""
    print(f"\n{'='*70}")
    print(f"  РЕЗУЛЬТАТЫ ТЕСТИРОВАНИЯ")
    print(f"  {summary['timestamp']}")
    print(f"{'='*70}")

    # Сводка по типам
    print(f"\n  {'doc_type':<25} {'engine':<10} {'файл':>4} {'PASS':>5} {'FAIL':>5} {'ERR':>4} {'%':>6} {'сек':>6}")
    print(f"  {'─'*65}")

    for t in summary["by_type"]:
        pct = (
            round(t["rules_pass"] / t["total_rules"] * 100, 1)
            if t["total_rules"] > 0 else 0
        )
        print(
            f"  {t['doc_type']:<25} {t['engine']:<10} {t['files']:>4} "
            f"{t['rules_pass']:>5} {t['rules_fail']:>5} {t['rules_error']:>4} "
            f"{pct:>5.1f}% {t['duration_sec']:>6.1f}"
        )

    # Итог
    total_rules = sum(t["total_rules"] for t in summary["by_type"])
    total_pass = sum(t["rules_pass"] for t in summary["by_type"])
    total_fail = sum(t["rules_fail"] for t in summary["by_type"])
    total_err = sum(t["rules_error"] for t in summary["by_type"])
    total_time = sum(t["duration_sec"] for t in summary["by_type"])
    total_pct = round(total_pass / total_rules * 100, 1) if total_rules > 0 else 0

    print(f"  {'─'*65}")
    print(
        f"  {'ИТОГО':<25} {'':10} {summary['total_files']:>4} "
        f"{total_pass:>5} {total_fail:>5} {total_err:>4} "
        f"{total_pct:>5.1f}% {total_time:>6.1f}"
    )

    # Файлы с ошибками
    error_files = [r for r in summary["results"] if r["status"] == "ERROR"]
    if error_files:
        print(f"\n  ОШИБКИ ({len(error_files)}):")
        for r in error_files:
            print(f"    {r['doc_type']}/{r['company']}: {r.get('error', '?')[:80]}")

    print(f"\n{'='*70}\n")


# ─────────────────────────────────────────────────
# Режим --report-only
# ─────────────────────────────────────────────────

def report_only(results_dir: Path):
    """Генерирует отчёт из последнего прогона без нового запуска."""
    latest = results_dir / "latest"
    if not latest.exists():
        print("Нет предыдущих прогонов в test_results/latest")
        sys.exit(1)

    summary_path = latest / "summary.json"
    if not summary_path.exists():
        print(f"Файл summary.json не найден в {latest}")
        sys.exit(1)

    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    print_summary_table(summary)

    # Перегенерация Excel
    generate_excel_report(summary, latest / "summary.xlsx")


# ─────────────────────────────────────────────────
# Главная логика
# ─────────────────────────────────────────────────

def main():
    """Точка входа тест-раннера."""
    # Загружаем .env (OPENAI_API_KEY и др.)
    _load_dotenv()

    parser = argparse.ArgumentParser(
        description="Интеграционный тест-раннер для аудита документов"
    )
    parser.add_argument(
        "--doc-type", default=None,
        help="Тип документа для тестирования (если не указан — все типы)"
    )
    parser.add_argument(
        "--company", default=None,
        help="Компания (фильтр по company в test_matrix)"
    )
    parser.add_argument(
        "--report-only", action="store_true",
        help="Только отчёт из последнего прогона (без нового запуска)"
    )
    parser.add_argument(
        "--timeout", type=int, default=DEFAULT_TIMEOUT,
        help=f"Таймаут на один файл в секундах (по умолчанию {DEFAULT_TIMEOUT})"
    )
    args = parser.parse_args()

    results_dir = BASE_DIR / "test_results"

    # Режим --report-only
    if args.report_only:
        report_only(results_dir)
        return

    # Загружаем тестовую матрицу
    matrix_path = BASE_DIR / "test_configs" / "test_matrix.json"
    if not matrix_path.exists():
        print(f"Файл тестовой матрицы не найден: {matrix_path}")
        sys.exit(1)

    suites = load_test_matrix(matrix_path)
    suites = filter_suites(suites, args.doc_type, args.company)

    if not suites:
        print("Нет тестов для запуска (проверьте --doc-type и --company)")
        sys.exit(1)

    # Считаем общее число файлов
    total_files = sum(len(s["files"]) for s in suites)
    total_types = len(suites)

    print(f"\n{'='*70}")
    print(f"  ТЕСТ-РАННЕР: интеграционный прогон")
    print(f"  Типов: {total_types}, файлов: {total_files}")
    print(f"  Таймаут на файл: {args.timeout} сек")
    print(f"{'='*70}\n")

    # Проверяем зависимости
    deps = check_dependencies(suites)
    if not deps["OPENAI_API_KEY"]:
        print("OPENAI_API_KEY не установлен. Аудит невозможен.")
        sys.exit(1)

    skip_vision = False
    if not deps.get("LibreOffice", True):
        print("WARNING: LibreOffice не найден — vision-типы будут пропущены")
        skip_vision = True

    # Создаём директорию прогона
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = results_dir / f"run_{timestamp}"
    run_dir.mkdir(parents=True, exist_ok=True)

    # Прогон тестов
    all_results = []
    file_idx = 0

    for suite in suites:
        doc_type = suite["doc_type"]
        engine = suite["engine"]
        multi_file = suite.get("multi_file", False)

        # Пропуск vision-типов если нет LibreOffice
        if skip_vision and engine == "vision":
            print(f"  SKIP {doc_type} (нет LibreOffice)")
            continue

        for file_info in suite["files"]:
            file_idx += 1
            company = file_info["company"]

            # Путь для результатов этого теста
            test_session_dir = run_dir / doc_type / company / "session"

            # Прогресс
            fname = Path(file_info.get("path") or file_info.get("primary", "")).name
            print(
                f"  [{file_idx}/{total_files}] {doc_type}/{company} "
                f"({fname[:40]}{'...' if len(fname) > 40 else ''})"
            )

            # Запуск
            result = run_single_test(
                doc_type=doc_type,
                engine=engine,
                file_info=file_info,
                session_dir=test_session_dir,
                timeout=args.timeout,
                multi_file=multi_file,
            )

            # Сохраняем per-test результаты
            test_out_dir = run_dir / doc_type / company
            _save_test_result(test_out_dir, result)

            # Статус в консоль
            status_icon = {"PASS": "OK", "FAIL": "FAIL", "ERROR": "ERR"}
            icon = status_icon.get(result["status"], "???")
            rules_info = ""
            if result["rules_status"]:
                p = sum(1 for s in result["rules_status"].values() if s == "PASS")
                f_cnt = sum(1 for s in result["rules_status"].values() if s == "FAIL")
                rules_info = f" ({p}P/{f_cnt}F)"
            print(f"         → {icon}{rules_info} [{result['duration_sec']}s]")

            all_results.append(result)

    # Генерация сводки
    print(f"\n  Генерация отчётов...")
    summary = generate_summary(all_results)

    # Сохраняем summary.json
    summary_path = run_dir / "summary.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    # Генерируем summary.xlsx
    generate_excel_report(summary, run_dir / "summary.xlsx")

    # Обновляем симлинк latest
    latest_link = results_dir / "latest"
    if latest_link.is_symlink() or latest_link.exists():
        latest_link.unlink()
    latest_link.symlink_to(run_dir.name)

    # Финальный вывод
    print_summary_table(summary)
    print(f"  Результаты: {run_dir}")
    print(f"  Симлинк:    {latest_link} -> {run_dir.name}")


def _save_test_result(test_dir: Path, result: Dict[str, Any]):
    """Сохраняет результаты одного теста в его директорию."""
    test_dir.mkdir(parents=True, exist_ok=True)

    # audit_result.json — основные данные
    audit_result = {
        "status": result["status"],
        "violations_count": result["violations_count"],
        "duration_sec": result["duration_sec"],
        "exit_code": result["exit_code"],
        "session_dir": result["session_dir"],
        "error": result.get("error"),
    }
    with open(test_dir / "audit_result.json", "w", encoding="utf-8") as f:
        json.dump(audit_result, f, ensure_ascii=False, indent=2)

    # rules_status.json — статус по каждому правилу
    if result["rules_status"]:
        with open(test_dir / "rules_status.json", "w", encoding="utf-8") as f:
            json.dump(result["rules_status"], f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
