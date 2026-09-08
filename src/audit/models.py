# START_MODULE_CONTRACT
# PURPOSE: Базовые контракты данных аудита — dataclass-модели и loader JSON-конфигов doc_type. Не зависят от движка/runner-ов; импортируются всеми слоями (main.py CLI, AuditEngine, doc_type-runner-ы).
# INPUTS: doc_configs/<doc_type>/config.json (читается `load_audit_config`).
# OUTPUTS: `SecondaryFileConfig`, `AuditConfig`, `AuditResult`, `load_audit_config`.
# KEYWORDS: dataclass, models, contracts, audit-config, audit-result.
# LINKS: config/llm.py (LLM_CONFIG — defaults для AuditConfig), main.py (CLI/AuditEngine), src/doc_type_validators/*.py (runner-ы возвращают AuditResult).
# RATIONALE:
#   Раньше эти модели лежали в main.py. Все 4 doc_type-runner-а (kpsc/kartochka/
#   drivers/plan_grafik) делали `from main import AuditResult` лениво внутри
#   функций — потому что main.py импортирует runner-ы сверху, а runner-ы
#   нуждаются в `AuditResult` для возврата результата (циркулярка). Вынос
#   моделей в отдельный модуль решает циркулярку.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

from config.llm import LLM_CONFIG
# END_IMPORTS


# START_SECONDARY_FILE_CONFIG
@dataclass
class SecondaryFileConfig:
    """
    Конфигурация вторичного файла для multi-file аудитов.

    Используется, когда аудит включает связку документов (например DOCX + XLSX).

    Поля:
        type — тип файла ("xlsx", "csv" и т.д.)
        parser — имя парсера (например "grafik_obhod")
        chunk_prefix — префикс для ключей вторичного файла (по умолчанию "xlsx_")
    """
    type: str
    parser: str
    chunk_prefix: str = "xlsx_"
# END_SECONDARY_FILE_CONFIG


# START_AUDIT_CONFIG
@dataclass
class AuditConfig:
    """
    Конфигурация аудита для конкретного типа документа.

    Загружается из doc_configs/<doc_type>/config.json через `load_audit_config`.

    Поля:
        doc_type — идентификатор типа документа (prikaz_ic, cheklist_eu, ...)
        doc_title — человекочитаемое название
        model — модель LLM для проверок (default из LLM_CONFIG.default_model)
        filename_pattern — ожидаемое имя файла для non-LLM проверки (подстрока)
        filename_keywords — список ключевых слов для проверки имени файла
        max_workers — количество параллельных LLM-запросов
        temperature — температура генерации
        secondary_file — конфиг вторичного файла (для multi-file аудитов)
        parser_by_ext — карта «расширение файла → имя парсера» для generic-пути
                        AuditEngine. Пример: {".pdf": "paddle", ".docx": "docx"}.
                        Пусто, если doc_type обслуживается special-движком.
        engine — имя кастомного движка (kpsc, kartochka_proekta, ...). Заполнено
                 только для special-движков; взаимоисключающе с parser_by_ext.
        llm_base_url, llm_max_tokens, reasoning_effort, llm_seed — параметры
            LLM-клиента (defaults из LLM_CONFIG).
        config_dir — путь к папке с конфигами (автоматически).
    """
    doc_type: str
    doc_title: str = ""
    model: str = field(default_factory=lambda: LLM_CONFIG.default_model)
    filename_pattern: str = ""
    filename_keywords: Optional[List[str]] = None
    max_workers: int = 1
    temperature: float = 0.0
    secondary_file: Optional[SecondaryFileConfig] = None
    # Выбор парсера по расширению файла для generic-пути. Для special-движков — пусто.
    parser_by_ext: Dict[str, str] = field(default_factory=dict)
    # Имя special-движка (kpsc, kartochka_proekta, ...). Взаимоисключающе с parser_by_ext.
    engine: Optional[str] = None
    llm_base_url: Optional[str] = field(default_factory=lambda: LLM_CONFIG.base_url)
    llm_max_tokens: int = field(default_factory=lambda: LLM_CONFIG.default_max_tokens)
    reasoning_effort: Optional[str] = field(default_factory=lambda: LLM_CONFIG.default_reasoning_effort)
    llm_seed: Optional[int] = field(default_factory=lambda: LLM_CONFIG.default_seed)
    config_dir: Path = field(default_factory=Path)
# END_AUDIT_CONFIG


# START_AUDIT_RESULT
@dataclass
class AuditResult:
    """
    Результат аудита.

    Поля:
        violations — список нарушений [{rule_index, rule_title, ...}]
        doc_type — тип документа
        session_dir — путь к директории сессии с логами
        duration_sec — продолжительность аудита в секундах
        rules_checked — количество проверенных правил
        target_path — путь к проверенному документу
        warnings — предупреждения парсера (например, не распознанные страницы PDF);
                   показываются пользователю и попадают в Excel листом «Предупреждения»
        unchecked_rules — правила, по которым модель не вернула вердикт (F15):
                   [{index, title, layer}]. Помечаются «НЕ ПРОВЕРЕНО», не выдаются за пройденные
    """
    violations: List[Dict[str, Any]] = field(default_factory=list)
    doc_type: str = ""
    session_dir: Path = field(default_factory=Path)
    duration_sec: float = 0.0
    rules_checked: int = 0
    target_path: str = ""
    warnings: List[str] = field(default_factory=list)
    unchecked_rules: List[Dict[str, Any]] = field(default_factory=list)
# END_AUDIT_RESULT


# START_LOAD_AUDIT_CONFIG
def load_audit_config(config_dir: Path) -> AuditConfig:
    """
    Назначение:
        Загружает AuditConfig из папки конфигов doc_configs/<doc_type>/.

    Вход:
        config_dir: путь к папке (содержит config.json).

    Выход:
        Готовый `AuditConfig` с заполненным `secondary_file` если он есть.

    Логика:
        XOR-валидация источника правды по рантайму: ровно одно из
        `parser_by_ext` (generic-путь через AuditEngine) или `engine` (special-
        движок) должно быть задано в config.json. Оба пустых/оба заполненных —
        ошибка конфигурации.
    """
    config_path = config_dir / "config.json"
    with open(config_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    # XOR-валидация источника правды по рантайму: либо parser_by_ext, либо engine.
    parser_by_ext = data.get("parser_by_ext", {})
    engine = data.get("engine", None)
    has_parser_map = isinstance(parser_by_ext, dict) and bool(parser_by_ext)
    has_engine = isinstance(engine, str) and bool(engine)
    if has_parser_map and has_engine:
        raise ValueError(f"Конфиг {config_path}: одновременно заданы parser_by_ext и engine — должен быть ровно один источник правды.")
    if not has_parser_map and not has_engine:
        raise ValueError(f"Конфиг {config_path}: не задан ни parser_by_ext (generic), ни engine (special). Укажите ровно одно.")
    config = AuditConfig(
        doc_type=data["doc_type"],
        doc_title=data.get("doc_title", data["doc_type"]),
        model=data.get("model", LLM_CONFIG.default_model),
        filename_pattern=data.get("filename_pattern", ""),
        filename_keywords=data.get("filename_keywords", None),
        max_workers=data.get("max_workers", 1),
        temperature=data.get("temperature", 0.0),
        parser_by_ext=parser_by_ext if has_parser_map else {},
        engine=engine if has_engine else None,
        llm_base_url=data.get("llm_base_url", LLM_CONFIG.base_url),
        llm_max_tokens=data.get("llm_max_tokens", LLM_CONFIG.default_max_tokens),
        reasoning_effort=data.get("reasoning_effort", LLM_CONFIG.default_reasoning_effort),
        llm_seed=data.get("llm_seed", LLM_CONFIG.default_seed),
        config_dir=config_dir,
    )
    if "secondary_file" in data:
        sf = data["secondary_file"]
        config.secondary_file = SecondaryFileConfig(
            type=sf["type"],
            parser=sf["parser"],
            chunk_prefix=sf.get("chunk_prefix", "xlsx_"),
        )
    return config
# END_LOAD_AUDIT_CONFIG
