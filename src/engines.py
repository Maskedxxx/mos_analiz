# START_MODULE_CONTRACT
# PURPOSE: Единый реестр спецдвижков (python-валидаторы по xlsx/комплектам, без LLM или с точечными LLM-вызовами) и определение движка типа документа. Единственное место, где перечислены раннеры: CLI (`main.py`) и веб-бэкенд (`src/api/server.py`) импортируют реестр отсюда.
# INPUTS: `doc_configs/<doc_type>/config.json` — поле `engine` (ключ реестра). Если поля нет — тип generic (LLM-путь через `AuditEngine`).
# OUTPUTS: `SPECIAL_ENGINE_RUNNERS` — {engine_key: runner(args) -> AuditResult}; `detect_engine(doc_type)` — ключ движка или `GENERIC_ENGINE`.
# KEYWORDS: registry, dispatch, special-engine, engine-key.
# LINKS: main.py (CLI-диспетчер), src/api/server.py (веб-диспетчер), src/doc_type_validators/*.py (раннеры), tests/test_registry.py (реестр ↔ конфиги).
# RATIONALE: Раньше реестр был продублирован в main.py и server.py; расхождение копий делало тип рабочим из CLI и невидимым в интерфейсе (так терялись шесть типов). Один модуль + тест согласованности с конфигами закрывают этот класс ошибок. Новый движок = строка здесь + `engine` в config.json типа.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
from pathlib import Path
from typing import Callable, Dict

from src.doc_type_validators.crosscheck import run_crosscheck_special
from src.doc_type_validators.drivers import run_drivers_special
from src.doc_type_validators.forma_0_3 import run_forma_0_3_special
from src.doc_type_validators.forma_0_4 import run_forma_0_4_special
from src.doc_type_validators.kartochka_proekta import run_kartochka_proekta_special
from src.doc_type_validators.kpsc import run_kpsc_special
from src.doc_type_validators.list_prisutstviya import run_list_prisutstviya_special
from src.doc_type_validators.plan_grafik import run_plan_grafik_special
# END_IMPORTS


# START_PATHS
# PURPOSE: Корень проекта — src/engines.py → parents[1]; конфиги типов лежат в doc_configs/.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
DOC_CONFIGS_DIR = PROJECT_ROOT / "doc_configs"
# END_PATHS


# START_REGISTRY
# PURPOSE: Маппинг engine-key → раннер спецдвижка. Раннер принимает `args` (namespace с полями
# doc_type, target, session_dir, parse_only, rule_filter, model, temperature) и возвращает AuditResult.
# Один раннер может обслуживать несколько типов (карточки 0.2/2.4, листы присутствия модулей 1/2) —
# тогда тип различается по `args.doc_type`, передавать его обязательно.

# Значение `engine` для generic-типов (поле в config.json отсутствует) — LLM-путь через AuditEngine.
GENERIC_ENGINE = "vision"

SPECIAL_ENGINE_RUNNERS: Dict[str, Callable] = {
    "drivers": run_drivers_special,
    "forma_0_3": run_forma_0_3_special,
    "forma_0_4": run_forma_0_4_special,
    "kpsc": run_kpsc_special,
    "kartochka_proekta_2_4": run_kartochka_proekta_special,
    "kartochka_proekta_0_2": run_kartochka_proekta_special,
    "list_prisutstviya": run_list_prisutstviya_special,
    "plan_grafik": run_plan_grafik_special,
    "crosscheck_2_4_0_6_0_5": run_crosscheck_special,
}


def detect_engine(doc_type: str, doc_configs_dir: Path = DOC_CONFIGS_DIR) -> str:
    """
    Назначение:
        Определяет движок типа документа по полю `engine` в его config.json.

    Вход:
        doc_type: Имя папки типа в doc_configs/.
        doc_configs_dir: Каталог конфигов (по умолчанию doc_configs/ в корне проекта).

    Выход:
        Ключ спецдвижка из SPECIAL_ENGINE_RUNNERS или GENERIC_ENGINE, если поля `engine` нет
        или config.json отсутствует (тогда тип идёт generic-путём через AuditEngine).
    """
    config_path = doc_configs_dir / doc_type / "config.json"
    if not config_path.exists():
        return GENERIC_ENGINE
    with open(config_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("engine", GENERIC_ENGINE)
# END_REGISTRY
