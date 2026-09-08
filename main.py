# START_MODULE_CONTRACT
# PURPOSE: Тонкий entrypoint проекта. CLI (`python main.py --doc-type ... --target ...`) и re-export `app` для `uvicorn main:app`.
# INPUTS: CLI-аргументы (argparse), doc_configs/<doc_type>/config.json (для определения engine).
# OUTPUTS: AuditResult в stdout/Excel-отчёт; exit code 0/1 по `--fail-on-violations`.
# KEYWORDS: cli, entrypoint, dispatch, uvicorn.
# LINKS: src/api/server.py (FastAPI app), src/audit/engine.py (AuditEngine), src/engines.py (реестр спецдвижков и detect_engine).
# RATIONALE:
#   После пошагового рефакторинга (steps 1..K) вся бизнес-логика и инфраструктура
#   живёт в src/. main.py — это ровно две роли:
#     1. CLI: парсит argparse, диспатчит на AuditEngine (multi_rule generic путь)
#        или на спецдвижок из реестра `src/engines.py` по `engine` из config.json.
#     2. Re-export `app` для `uvicorn main:app` (этот entrypoint исторически жил
#        в main, тесты `tests/test_smoke_api.py` тоже используют `main.app`).
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import argparse
import json
import sys
from pathlib import Path

# Re-export для `uvicorn main:app` и tests/test_smoke_api.py.
from src.api.server import app  # noqa: F401

from src.audit.engine import AuditEngine
from src.audit.input_check import check_input_file
from src.doc_type_parsers.kpsc import kpsc_parse_kpsc_header_build_payload
from src.doc_type_validators.kpsc import _run_kpsc_parsers
from src.engines import SPECIAL_ENGINE_RUNNERS, detect_engine  # noqa: F401 — реестр реэкспортируется для тестов
# END_IMPORTS


# START_PATHS
# PURPOSE: Корневые пути для CLI. PROJECT_ROOT = директория main.py.
PROJECT_ROOT = Path(__file__).resolve().parent
DOC_CONFIGS_DIR = PROJECT_ROOT / "doc_configs"
# END_PATHS


# START_DISPATCH
# PURPOSE: Диспетчер CLI берёт реестр спецдвижков и detect_engine из src/engines.py —
# единственного места, где перечислены раннеры (общее с веб-бэкендом).
# END_DISPATCH


# START_CLI
def main() -> None:
    """CLI entrypoint. Диспатчит запрос на special-runner или AuditEngine."""
    parser = argparse.ArgumentParser(description="Единый аудит документов")
    parser.add_argument("--doc-type", default=None, help="Тип документа (имя папки в doc_configs/)")
    parser.add_argument("--target", default=None, help="Путь к целевому документу")
    parser.add_argument("--model", default=None, help="Модель LLM (переопределяет конфиг)")
    parser.add_argument("--temperature", type=float, default=None, help="Температура генерации")
    parser.add_argument("--rule-filter", default=None, help="Проверить только указанное правило (только для special-runner-ов: kpsc/kartochka_proekta)")
    parser.add_argument("--parse-only", action="store_true", help="Только парсинг, без проверок")
    parser.add_argument("--print-prompts", action="store_true", help="Печатать промпты без вызова LLM")
    parser.add_argument("--session-dir", default=None, help="Директория для логов сессии")
    parser.add_argument("--out-xlsx", default=None, help="Путь для сохранения Excel")
    parser.add_argument("--secondary", default=None, help="Путь к вторичному файлу")
    parser.add_argument("--list-types", action="store_true", help="Показать список доступных типов документов")
    parser.add_argument("--fail-on-violations", action="store_true", help="Вернуть exit code 1, если аудит завершился с нарушениями")
    args = parser.parse_args()

    if args.list_types:
        doc_types = AuditEngine.list_doc_types()
        if not doc_types:
            print("Нет доступных типов документов в doc_configs/")
            sys.exit(0)
        print(f"\nДоступные типы документов ({len(doc_types)}):")
        print("-" * 50)
        for dt in doc_types:
            title = f" — {dt['doc_title']}" if dt["doc_title"] else ""
            # Повреждённая конфигурация — показать причину, тип к запуску не пригоден.
            broken = f"  [НЕДОСТУПЕН: {dt['broken_reason']}]" if dt.get("broken") else ""
            print(f"  {dt['doc_type']}{title}{broken}")
        print("\nИспользование: python main.py --doc-type <doc_type> --target <file>")
        sys.exit(0)

    if not args.doc_type:
        parser.error("--doc-type обязателен (или используйте --list-types)")
    if not args.target:
        parser.error("--target обязателен")

    # Тип и файл проверяются до запуска — одна строка в stderr и exit 2 вместо traceback
    # из глубины парсера (находки 2.4b, 2.4c, 2.4d, 2.4i; CLI 3.7a).
    known = {t["doc_type"]: t for t in AuditEngine.list_doc_types()}
    if args.doc_type not in known:
        print(f"Ошибка: тип документа не найден: {args.doc_type} (см. --list-types)", file=sys.stderr)
        sys.exit(2)
    if known[args.doc_type].get("broken"):
        print(f"Ошибка: тип документа «{args.doc_type}» настроен некорректно: {known[args.doc_type]['broken_reason']}", file=sys.stderr)
        sys.exit(2)
    try:
        check_input_file(Path(args.target))
    except ValueError as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        sys.exit(2)

    engine_type = detect_engine(args.doc_type)
    if engine_type in SPECIAL_ENGINE_RUNNERS:
        result = SPECIAL_ENGINE_RUNNERS[engine_type](args)
    else:
        engine = AuditEngine(args.doc_type)
        result = engine.run(
            target_path=args.target,
            model=args.model,
            temperature=args.temperature,
            parse_only=args.parse_only,
            print_prompts=args.print_prompts,
            session_dir=args.session_dir,
            out_xlsx=args.out_xlsx,
            secondary_path=args.secondary,
        )

    if result.violations and not args.fail_on_violations:
        sys.exit(0)
    sys.exit(1 if result.violations else 0)
# END_CLI


# START_TEST_COMPAT
# PURPOSE: Алиасы используются тестами через `from main import ...`. При следующем
# рефакторинге тестов можно переключить их импорты на `src.format_parsers` /
# `src.doc_type_parsers.kpsc` / `src.doc_type_validators.kpsc` и убрать эти строки.
# Текущие потребители:
#   - tests/test_smoke_parsers.py    — parse_docx, parse_pptx
#   - tests/test_kpsc_parsers.py     — build_kpsc_header_payload, run_kpsc_parsers
#   - tests/test_smoke_pipeline.py   — AuditEngine (уже импортируется выше)
from src.format_parsers import parse_docx, parse_pptx  # noqa: F401

build_kpsc_header_payload = kpsc_parse_kpsc_header_build_payload
run_kpsc_parsers = _run_kpsc_parsers
# END_TEST_COMPAT


if __name__ == "__main__":
    main()
