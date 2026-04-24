# START_MODULE_CONTRACT
# PURPOSE: Пакет doc-type-специфичных парсеров xlsx. В отличие от format_parsers/, эти парсеры знают про конкретные листы, координаты ячеек и семантику полей своего типа документа. Возвращают не универсальный ParsedDocument, а свои типизированные структуры.
# INPUTS: Импорт-время — подмодули пакета.
# OUTPUTS: Публичный API: `parse_<doc_type>` для каждого типа + соответствующие типы контрактов.
# KEYWORDS: doc-type-parsers, xlsx, structured-parsing, per-doc-type.
# LINKS: main.py (валидаторы и движки), doc_configs/<doc_type>/config.json.
# RATIONALE: Формат-парсеры универсальны (один парсер — много doc_types); doc_type-парсеры наоборот — специфичны для своего типа документа. Разные слои, разные места в репо.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_REEXPORTS
from src.doc_type_parsers.drivers import parse_excel_to_json
from src.doc_type_parsers.grafik_obhod import GrafikObhodDocument, parse_grafik_obhod
from src.doc_type_parsers.kartochka_proekta import (
    parse_dropdown,
    parse_kartochka_main,
    parse_metodika,
)
from src.doc_type_parsers.kpsc import (
    parse_kpsc_header,
    parse_kpsc_table1,
    parse_legend,
    parse_loss_digitization,
    parse_pa1_chart,
    parse_pa1_table,
    parse_pokazateli,
    parse_spaghetti_problems,
    parse_spaghetti_sheet,
)
from src.doc_type_parsers.plan_grafik import PlanGrafikDocument, parse_plan_grafik

__all__ = [
    # grafik_obhod
    "GrafikObhodDocument",
    "parse_grafik_obhod",
    # plan_grafik
    "PlanGrafikDocument",
    "parse_plan_grafik",
    # kartochka_proekta
    "parse_dropdown",
    "parse_kartochka_main",
    "parse_metodika",
    # kpsc
    "parse_kpsc_header",
    "parse_kpsc_table1",
    "parse_legend",
    "parse_loss_digitization",
    "parse_pa1_chart",
    "parse_pa1_table",
    "parse_pokazateli",
    "parse_spaghetti_problems",
    "parse_spaghetti_sheet",
    # drivers
    "parse_excel_to_json",
]
# END_REEXPORTS
