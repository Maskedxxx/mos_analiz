# START_MODULE_CONTRACT
# PURPOSE: Тонкий entrypoint проекта. CLI (`python main.py --doc-type ... --target ...`) и re-export `app` для `uvicorn main:app`.
# INPUTS: CLI-аргументы (argparse), doc_configs/<doc_type>/config.json (для определения engine).
# OUTPUTS: AuditResult в stdout/Excel-отчёт; exit code 0/1 по `--fail-on-violations`.
# KEYWORDS: cli, entrypoint, dispatch, uvicorn.
# LINKS: src/api/server.py (FastAPI app), src/audit/engine.py (AuditEngine), src/doc_type_validators/{drivers,kpsc,kartochka_proekta,plan_grafik}.py (special-runner-ы).
# RATIONALE:
#   После пошагового рефакторинга (steps 1..K) вся бизнес-логика и инфраструктура
#   живёт в src/. main.py — это ровно две роли:
#     1. CLI: парсит argparse, диспатчит на AuditEngine (multi_rule generic путь)
#        или на один из 4 special-runner-ов по `engine` из config.json.
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
from src.doc_type_parsers.kpsc import kpsc_parse_kpsc_header_build_payload
from src.doc_type_validators.drivers import run_drivers_special
from src.doc_type_validators.kartochka_proekta import run_kartochka_proekta_special
from src.doc_type_validators.kpsc import _run_kpsc_parsers, run_kpsc_special
from src.doc_type_validators.plan_grafik import run_plan_grafik_special
# END_IMPORTS


# START_PATHS
# PURPOSE: Корневые пути для CLI. PROJECT_ROOT = директория main.py.
PROJECT_ROOT = Path(__file__).resolve().parent
DOC_CONFIGS_DIR = PROJECT_ROOT / "doc_configs"
# END_PATHS


# START_DISPATCH
# PURPOSE: Маппинг engine-key → special-runner callable. Используется CLI-`main()`
# для выбора пайплайна по `engine`-полю в doc_configs/<doc_type>/config.json.
SPECIAL_ENGINE_RUNNERS = {
    "drivers": run_drivers_special,
    "kpsc": run_kpsc_special,
    "kartochka_proekta": run_kartochka_proekta_special,
    "plan_grafik": run_plan_grafik_special,
}


def _detect_engine(doc_type: str) -> str:
    """
    Читает `engine` из doc_configs/<doc_type>/config.json.

    Returns:
        Имя special-движка (kpsc/kartochka_proekta/drivers/plan_grafik) или
        'vision' для generic multi_rule пути через AuditEngine.
    """
    config_path = DOC_CONFIGS_DIR / doc_type / "config.json"
    if not config_path.exists():
        return "vision"
    with open(config_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("engine", "vision")
# END_DISPATCH


# START_CLI
def main() -> None:
    """CLI entrypoint. Диспатчит запрос на special-runner или AuditEngine."""
    parser = argparse.ArgumentParser(description="Единый аудит документов")
    parser.add_argument("--doc-type", default=None, help="Тип документа (имя папки в doc_configs/)")
    parser.add_argument("--target", default=None, help="Путь к целевому документу")
    parser.add_argument("--model", default=None, help="Модель LLM (переопределяет конфиг)")
    parser.add_argument("--temperature", type=float, default=None, help="Температура генерации")
    parser.add_argument("--rule-filter", default=None, help="Проверить только указанное правило")
    parser.add_argument("--parse-only", action="store_true", help="Только парсинг, без проверок")
    parser.add_argument("--print-prompts", action="store_true", help="Печатать промпты без вызова LLM")
    parser.add_argument("--session-dir", default=None, help="Директория для логов сессии")
    parser.add_argument("--chunk-filter", default=None, help="Парсить только указанный чанк")
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
            print(f"  {dt['doc_type']}{title}")
        print("\nИспользование: python main.py --doc-type <doc_type> --target <file>")
        sys.exit(0)

    if not args.doc_type:
        parser.error("--doc-type обязателен (или используйте --list-types)")
    if not args.target:
        parser.error("--target обязателен")

    engine_type = _detect_engine(args.doc_type)
    if engine_type in SPECIAL_ENGINE_RUNNERS:
        result = SPECIAL_ENGINE_RUNNERS[engine_type](args)
    else:
        rule_filter = int(args.rule_filter) if args.rule_filter is not None else None
        engine = AuditEngine(args.doc_type)
        result = engine.run(
            target_path=args.target,
            model=args.model,
            temperature=args.temperature,
            rule_filter=rule_filter,
            parse_only=args.parse_only,
            print_prompts=args.print_prompts,
            session_dir=args.session_dir,
            chunk_filter=args.chunk_filter,
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
