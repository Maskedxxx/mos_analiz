# START_MODULE_CONTRACT
# PURPOSE: Collapse all backend Python runtime and business logic into one readable source monolith.
# INPUTS: CLI args, FastAPI requests, document files, doc_configs assets, model service endpoints.
# OUTPUTS: Audit results, Excel reports, session logs, FastAPI responses, CLI exit codes.
# KEYWORDS: monolith, audit-engine, fastapi, cli, parsers, validators, reporting.
# LINKS: doc_configs/, docs/monolith_refactor_journal.md, grace_methodology.md.
# RATIONALE: Keep runtime truth in one file and remove legacy backend file indirection.
# END_MODULE_CONTRACT

# START_IMPORTS
from __future__ import annotations

from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pathlib import Path
from pydantic import BaseModel
from sse_starlette.sse import EventSourceResponse
from types import SimpleNamespace
from typing import Any, Callable, Dict, Iterable, List, Optional, Set, Tuple, Union
from xml.etree import ElementTree as ET

import argparse
import asyncio
import base64
import faulthandler
import hashlib
import importlib
import io
import json
import logging
import os
import openpyxl
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import unicodedata
import uuid
import zipfile
import xml.etree.ElementTree as ET

try:
    import fitz
except ImportError:
    fitz = None  # type: ignore[assignment]

try:
    import numpy as np
except ImportError:
    np = None  # type: ignore[assignment]

try:
    import pandas as pd
except ImportError:
    pd = None  # type: ignore[assignment]

try:
    import requests
except ImportError:
    requests = None  # type: ignore[assignment]

try:
    from PIL import Image
except ImportError:
    Image = None  # type: ignore[assignment]

try:
    from docx import Document
except ImportError:
    Document = None  # type: ignore[assignment]

try:
    from openai import AsyncOpenAI, OpenAI
except ImportError:
    AsyncOpenAI = None  # type: ignore[assignment]
    OpenAI = None  # type: ignore[assignment]

def _should_disable_local_reasoning(base_url: Optional[str]) -> bool:
    if not base_url:
        return False
    return ':11437' in str(base_url)


def _merge_disable_thinking_extra_body(extra_body: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    merged = dict(extra_body or {})
    chat_template_kwargs = dict(merged.get('chat_template_kwargs') or {})
    chat_template_kwargs.setdefault('enable_thinking', False)
    merged['chat_template_kwargs'] = chat_template_kwargs
    return merged


def _patch_openai_chat_client(client: Any, base_url: Optional[str]) -> Any:
    if not _should_disable_local_reasoning(base_url):
        return client
    completions = getattr(getattr(client, 'chat', None), 'completions', None)
    create_fn = getattr(completions, 'create', None)
    if completions is None or create_fn is None:
        return client
    if getattr(create_fn, '_codex_disable_thinking_patch', False):
        return client
    if asyncio.iscoroutinefunction(create_fn):

        async def patched_create(*args, **kwargs):
            kwargs['extra_body'] = _merge_disable_thinking_extra_body(kwargs.get('extra_body'))
            return await create_fn(*args, **kwargs)
    else:

        def patched_create(*args, **kwargs):
            kwargs['extra_body'] = _merge_disable_thinking_extra_body(kwargs.get('extra_body'))
            return create_fn(*args, **kwargs)
    patched_create._codex_disable_thinking_patch = True  # type: ignore[attr-defined]
    completions.create = patched_create
    return client


if OpenAI is not None:
    _OPENAI_SYNC_CLIENT_CTOR = OpenAI

    def OpenAI(*args, **kwargs):
        # Prevent silent infinite hangs on local provider completions.
        kwargs.setdefault("timeout", float(os.environ.get("OPENAI_TIMEOUT_SEC", "45")))
        client = _OPENAI_SYNC_CLIENT_CTOR(*args, **kwargs)
        return _patch_openai_chat_client(client, kwargs.get('base_url'))


if AsyncOpenAI is not None:
    _OPENAI_ASYNC_CLIENT_CTOR = AsyncOpenAI

    def AsyncOpenAI(*args, **kwargs):
        kwargs.setdefault("timeout", float(os.environ.get("OPENAI_TIMEOUT_SEC", "45")))
        client = _OPENAI_ASYNC_CLIENT_CTOR(*args, **kwargs)
        return _patch_openai_chat_client(client, kwargs.get('base_url'))

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.chart._chart import ChartBase
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.utils import range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

try:
    from pptx import Presentation
    from pptx.util import Emu
except ImportError:
    Presentation = None  # type: ignore[assignment]
    Emu = None  # type: ignore[assignment]

try:
    from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential
except ImportError:
    def retry(*args, **kwargs):
        def decorator(fn):
            return fn
        return decorator

    def stop_after_attempt(*args, **kwargs):
        return None

    def wait_exponential(*args, **kwargs):
        return None

    def retry_if_exception_type(*args, **kwargs):
        return None

# END_IMPORTS

# START_PATHS_AND_SETTINGS
# PURPOSE: Centralize repository-relative runtime paths and auth/session defaults.
# INPUTS: Environment variables and the monolith location on disk.
# OUTPUTS: Shared path constants and default auth values used by CLI and API.
# KEYWORDS: paths, settings, uploads, doc-configs, auth.
# LINKS: doc_configs/, uploads/, logs_result/.
# RATIONALE: One monolith still needs one explicit filesystem contract.
PROJECT_ROOT = Path(__file__).resolve().parent
DOC_CONFIGS_DIR = PROJECT_ROOT / 'doc_configs'
LOGS_RESULT_DIR = PROJECT_ROOT / 'logs_result'

# END_PATHS_AND_SETTINGS

# START_DOC_CONFIG_LOADING
# PURPOSE: Keep small repository-level helpers close to the doc-config contract.
# INPUTS: doc_type identifiers and doc_configs paths.
# OUTPUTS: Resolved config directories and engine strings.
# KEYWORDS: doc-configs, engine, parser, discovery.
# LINKS: doc_configs/<doc_type>/config.json.
# RATIONALE: CLI, API, and tests all depend on the same config discovery behavior.
def resolve_doc_config_dir(doc_type: str) -> Path:
    return DOC_CONFIGS_DIR / doc_type

def read_doc_type_engine(doc_type: str) -> str:
    config_path = resolve_doc_config_dir(doc_type) / 'config.json'
    if not config_path.exists():
        return 'vision'
    with open(config_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data.get('engine', 'vision')

# END_DOC_CONFIG_LOADING

# START_MODELS
# PURPOSE: Hold core dataclasses, config structures, and parser/model data contracts.
# INPUTS: JSON config payloads, parsed documents, validation results.
# OUTPUTS: Stable in-memory contracts shared across the monolith.
# KEYWORDS: dataclass, config, result, chunk, vision, paddle.
# LINKS: audit_engine/models.py, audit_engine/vision_parser/config.py, audit_engine/paddle_parser/config.py.
# RATIONALE: Shared contracts must be defined before any parser or service code.

# START_SOURCE_MODELS
# PURPOSE: Inlined source from audit_engine/models.py.
@dataclass
class RuleSpec:
    """
    Спецификация одного правила проверки.

    Поля:
        index — номер правила (1, 2, 3...)
        title — название правила
        scope — чанк(и) документа для проверки
        compare — тип сравнения: template / target_only / cross_check
        llm — использует ли LLM (False = non-LLM проверка)
        instructions — текстовые инструкции для LLM (из поля "content" в JSON)
        context_filter — фильтры контекста по чанкам {scope: [patterns]}
        context_filter_mode — режим фильтрации: "paragraphs" или "headers_only"
    """
    index: int
    title: str
    scope: Union[str, List[str]]
    compare: str
    llm: bool
    instructions: List[str] = field(default_factory=list)
    context_filter: Dict[str, List[str]] = field(default_factory=dict)
    context_filter_mode: str = 'paragraphs'
    max_chars: Optional[int] = None

@dataclass
class SecondaryFileConfig:
    """
    Конфигурация вторичного файла для multi-file аудитов.

    Используется, когда аудит включает связку документов (например DOCX + XLSX).

    Поля:
        type — тип файла ("xlsx", "csv" и т.д.)
        parser — имя парсера из audit_engine/parsers/ (например "grafik_obhod")
        chunk_prefix — префикс для чанков вторичного файла (по умолчанию "xlsx_")
    """
    type: str
    parser: str
    chunk_prefix: str = 'xlsx_'

@dataclass
class AuditConfig:
    """
    Конфигурация аудита для конкретного типа документа.

    Загружается из doc_configs/<doc_type>/config.json.

    Поля:
        doc_type — идентификатор типа документа (prikaz_ic, cheklist_eu, ...)
        doc_title — человекочитаемое название
        model — модель LLM для проверок
        filename_pattern — ожидаемое имя файла для non-LLM проверки (подстрока)
        filename_keywords — список ключевых слов для проверки имени файла (все должны быть в имени)
        max_workers — количество параллельных LLM-запросов
        temperature — температура генерации
        secondary_file — конфиг вторичного файла (для multi-file аудитов)
        parser_by_ext — карта «расширение файла → имя парсера» для generic-пути AuditEngine.
                        Пример: {".pdf": "paddle", ".docx": "docx"}. Пусто, если doc_type
                        обслуживается special-движком.
        engine — имя кастомного движка для doc_type (kpsc, kartochka_proekta, ...).
                 Заполнено только для special-движков; взаимоисключающе с parser_by_ext.
        llm_base_url, llm_max_tokens, reasoning_effort, llm_seed — параметры LLM-клиента.
        config_dir — путь к папке с конфигами (автоматически)
        rules_path — путь к rules.json (автоматически)
    """
    doc_type: str
    doc_title: str = ''
    model: str = field(default_factory=lambda: LLM_CONFIG.default_model)
    filename_pattern: str = ''
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
    rules_path: Path = field(default_factory=Path)

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
    """
    violations: List[Dict[str, Any]] = field(default_factory=list)
    doc_type: str = ''
    session_dir: Path = field(default_factory=Path)
    duration_sec: float = 0.0
    rules_checked: int = 0
    target_path: str = ''

def load_rules(rules_path: str) -> List[RuleSpec]:
    """
    Загружает правила из JSON-файла.

    Формат JSON:
    {
        "правила": [
            {"index": 1, "title": "...", "scope": "...", "compare": "...", "llm": true/false, "content": [...]}
        ]
    }
    """
    with open(rules_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    rules = []
    for rule in data.get('правила', []):
        spec = RuleSpec(index=rule['index'], title=rule['title'], scope=rule['scope'], compare=rule['compare'], llm=rule['llm'], instructions=rule.get('content', []), context_filter=rule.get('context_filter', {}), context_filter_mode=rule.get('context_filter_mode', 'paragraphs'), max_chars=rule.get('max_chars', None))
        rules.append(spec)
    return rules

def load_audit_config(config_dir: Path) -> AuditConfig:
    """
    Загружает AuditConfig из папки конфигов.

    Ищет config.json, rules.json, template/*.

    Правило выбора рантайма: ровно одно из двух полей должно быть задано —
    либо `parser_by_ext` (generic-путь через AuditEngine), либо `engine`
    (special-движок). Оба пустых/оба заполненных — ошибка конфигурации.
    """
    config_path = config_dir / 'config.json'
    with open(config_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    # XOR-валидация источника правды по рантайму: либо parser_by_ext, либо engine.
    parser_by_ext = data.get('parser_by_ext', {})
    engine = data.get('engine', None)
    has_parser_map = isinstance(parser_by_ext, dict) and bool(parser_by_ext)
    has_engine = isinstance(engine, str) and bool(engine)
    if has_parser_map and has_engine:
        raise ValueError(f"Конфиг {config_path}: одновременно заданы parser_by_ext и engine — должен быть ровно один источник правды.")
    if not has_parser_map and not has_engine:
        raise ValueError(f"Конфиг {config_path}: не задан ни parser_by_ext (generic), ни engine (special). Укажите ровно одно.")
    config = AuditConfig(
        doc_type=data['doc_type'],
        doc_title=data.get('doc_title', data['doc_type']),
        model=data.get('model', LLM_CONFIG.default_model),
        filename_pattern=data.get('filename_pattern', ''),
        filename_keywords=data.get('filename_keywords', None),
        max_workers=data.get('max_workers', 1),
        temperature=data.get('temperature', 0.0),
        parser_by_ext=parser_by_ext if has_parser_map else {},
        engine=engine if has_engine else None,
        llm_base_url=data.get('llm_base_url', LLM_CONFIG.base_url),
        llm_max_tokens=data.get('llm_max_tokens', LLM_CONFIG.default_max_tokens),
        reasoning_effort=data.get('reasoning_effort', LLM_CONFIG.default_reasoning_effort),
        llm_seed=data.get('llm_seed', LLM_CONFIG.default_seed),
        config_dir=config_dir,
        rules_path=config_dir / 'rules.json',
    )
    if 'secondary_file' in data:
        sf = data['secondary_file']
        config.secondary_file = SecondaryFileConfig(type=sf['type'], parser=sf['parser'], chunk_prefix=sf.get('chunk_prefix', 'xlsx_'))
    return config

models_module = SimpleNamespace(RuleSpec=RuleSpec, SecondaryFileConfig=SecondaryFileConfig, AuditConfig=AuditConfig, AuditResult=AuditResult, load_rules=load_rules, load_audit_config=load_audit_config)

# END_SOURCE_MODELS

# END_MODELS

# START_PREPROCESSORS
# PURPOSE: Normalize document chunks before LLM comparison or semantic filtering.
# INPUTS: doc_type, scope name, raw chunk text.
# OUTPUTS: Preprocessed text registered by doc_type and scope.
# KEYWORDS: preprocessors, registry, normalization.
# LINKS: audit_engine/preprocessors/*.py.
# RATIONALE: Preprocessors remain data-shaping business logic even inside one file.

# END_PREPROCESSORS

# START_PARSERS
# PURPOSE: Extract structured text from docx, pptx, pdf, xlsx, OCR, and special-engine sources.
# INPUTS: document file paths, chunk configs, model endpoints.
# OUTPUTS: Parsed document dictionaries and parser artifacts.
# KEYWORDS: parsers, docx, pptx, paddle, ocr, vision, xlsx.
# LINKS: audit_engine/*parser*.py, audit_engine/*/parser*.py.
# RATIONALE: Parsing is the widest dependency fan-out, so it sits before validation and runtime orchestration.

# START_SOURCE_FORMAT_PARSERS
# PURPOSE: Bridge low-level DOCX/PPTX format parsing from src/format_parsers.py into the monolith runtime.
# INPUTS: Import-time module loading and parser function calls from _parse_document.
# OUTPUTS: Stable parser callables and compatibility namespaces used by existing runtime code and tests.
# KEYWORDS: bridge, format-parser, docx, pptx, compatibility, weak-coupling.
# LINKS: src/format_parsers.py, tests/test_smoke_parsers.py.
# RATIONALE: The monolith should orchestrate parser usage, not hold low-level format-reader implementations inline.
from src.format_parsers import _extract_elements_in_order
from src.format_parsers import _extract_slide_text
from src.format_parsers import _ocr_header_images
from src.format_parsers import _textframe_to_text
from src.format_parsers import docx_parser__table_to_html
from src.format_parsers import docx_parser_logger
from src.format_parsers import parse_docx
from src.format_parsers import parse_pptx
from src.format_parsers import pptx_parser__table_to_html
from src.format_parsers import pptx_parser_logger
from config.llm import LLM_CONFIG
from config.parsers import PARSERS_CONFIG
from src.format_parsers.pdf import parse_pdf
from src.format_parsers.pdf._clients import VLMClient
from src.doc_type_parsers.grafik_obhod import parse_grafik_obhod
from src.doc_type_parsers.plan_grafik import parse_plan_grafik
from src.doc_type_parsers.kartochka_proekta import parse_dropdown, parse_kartochka_main, parse_metodika
from src.doc_type_parsers.kpsc import (
    kpsc_parse_kpsc_header_build_payload,
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
from src.doc_type_parsers.drivers import parse_excel_to_json
from src.doc_type_validators.kartochka_proekta import (
    KARTOCHKA_VALIDATOR_MODULE_DISPATCH,
    run_kartochka_validator_module,
)
from src.doc_type_validators.kpsc import (
    KPSC_VALIDATOR_MODULE_DISPATCH,
    run_kpsc_validator_module,
)
from src.doc_type_validators.drivers import run_drivers_special
from src.doc_type_validators.plan_grafik import run_all_validators
from src.llm import (
    call_llm,
    load_methodology_config,
    load_multi_rule_config,
    parse_json_response,
    resolve_runtime_llm_model,
    run_multi_rule_audit,
)

docx_parser_module = SimpleNamespace(
    logger=docx_parser_logger,
    parse_docx=parse_docx,
    _ocr_header_images=_ocr_header_images,
    _extract_elements_in_order=_extract_elements_in_order,
    _table_to_html=docx_parser__table_to_html,
)

pptx_parser_module = SimpleNamespace(
    logger=pptx_parser_logger,
    parse_pptx=parse_pptx,
    _extract_slide_text=_extract_slide_text,
    _textframe_to_text=_textframe_to_text,
    _table_to_html=pptx_parser__table_to_html,
)
# END_SOURCE_FORMAT_PARSERS

# END_PARSERS

# START_VALIDATORS
# PURPOSE: Evaluate deterministic and LLM-backed validation logic for special engines.
# INPUTS: parser outputs, validation rules, workbook data, model endpoints.
# OUTPUTS: Rule results, discrepancies, and validation JSON payloads.
# KEYWORDS: validators, kpsc, kartochka, plan-grafik.
# LINKS: audit_engine/kpsc/validation_scripts/, audit_engine/kartochka_proekta/validation_scripts/, audit_engine/plan_grafik/validators.py.
# RATIONALE: Validation logic is the business core that the monolith must expose directly.



# END_VALIDATORS


# START_REPORTING
# PURPOSE: Persist session logs, prompts, responses, and Excel reports.
# INPUTS: runtime events, prompts, parsed docs, violations.
# OUTPUTS: log folders, JSON artifacts, Excel workbooks.
# KEYWORDS: logging, reports, excel, session-artifacts.
# LINKS: audit_engine/logger.py, audit_engine/excel_reporter.py.
# RATIONALE: Reporting must remain close to runtime truth for post-run inspection.

# START_SOURCE_LOGGER
# PURPOSE: Inlined source from audit_engine/logger.py.
class PipelineLogger:
    """
    Логгер для всего пайплайна аудита.
    Сохраняет все этапы в отдельные файлы для отладки.
    """

    def __init__(self, log_dir: Path, doc_type: str=''):
        """
        Инициализация логгера.

        Args:
            log_dir: директория для логов сессии
            doc_type: тип документа (для заголовка лога)
        """
        self.log_dir = log_dir
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.parsed_dir = log_dir / 'parsed_docs'
        self.parsed_dir.mkdir(exist_ok=True)
        self.main_log = log_dir / 'pipeline.log'
        self._init_main_log(doc_type)

    def _init_main_log(self, doc_type: str):
        """Инициализирует главный лог-файл."""
        with open(self.main_log, 'w', encoding='utf-8') as f:
            f.write(f'{'=' * 80}\n')
            f.write(f'AUDIT PIPELINE LOG — {doc_type or 'Universal Audit Engine'}\n')
            f.write(f'Started: {datetime.now().isoformat()}\n')
            f.write(f'{'=' * 80}\n\n')

    def log(self, message: str):
        """Записывает сообщение в главный лог и stderr."""
        timestamp = datetime.now().strftime('%H:%M:%S')
        with open(self.main_log, 'a', encoding='utf-8') as f:
            f.write(f'[{timestamp}] {message}\n')
        print(f'[{timestamp}] {message}', file=sys.stderr)

    def log_parsed_doc(self, doc: Dict[str, Any], name: str):
        """Сохраняет распарсенный документ в JSON."""
        filepath = self.parsed_dir / f'{name}.json'
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(doc, f, ensure_ascii=False, indent=2)
        self.log(f'📄 Распарсенный документ сохранён: {filepath}')

    def log_final_results(self, violations: List[Dict]):
        """Сохраняет финальные результаты."""
        filepath = self.log_dir / 'final_results.json'
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(violations, f, ensure_ascii=False, indent=2)
        self.log(f'📊 Финальные результаты сохранены: {filepath}')

logger_module = SimpleNamespace(PipelineLogger=PipelineLogger)

# END_SOURCE_LOGGER

# START_SOURCE_EXCEL_REPORTER
# PURPOSE: Inlined source from audit_engine/excel_reporter.py.
_DISPLAY_HEADERS = ['№', 'Проверка', 'Тип проверки', 'Статус', 'Целевой документ', 'Различие']

_FONT_SIZE = 14

_HEADER_FONT = Font(name='Calibri', size=_FONT_SIZE, bold=True, color='FFFFFF')

_HEADER_FILL = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')

_CELL_FONT = Font(name='Calibri', size=_FONT_SIZE)

_WRAP_ALIGNMENT = Alignment(wrap_text=True, vertical='top')

_CENTER_ALIGNMENT = Alignment(horizontal='center', vertical='top')

_THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

_MIN_WIDTHS = [6, 35, 18, 10, 40, 40]

_MAX_WIDTHS = [6, 50, 18, 10, 60, 60]

_OK_FILL = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')

_FAIL_FILL = PatternFill(start_color='FCE4EC', end_color='FCE4EC', fill_type='solid')

_OK_FONT = Font(name='Calibri', size=_FONT_SIZE, bold=True, color='1F7A1F')

_FAIL_FONT = Font(name='Calibri', size=_FONT_SIZE, bold=True, color='CC0000')

_METH_OK_FILL = PatternFill(start_color='D6EAF8', end_color='D6EAF8', fill_type='solid')

_METH_FAIL_FILL = PatternFill(start_color='FADBD8', end_color='FADBD8', fill_type='solid')

def _layer_label(layer: str) -> str:
    if layer == 'methodology':
        return 'Методическая'
    return 'Базовая'

def save_to_excel(violations: List[Dict[str, Any]], output_path: str, all_rules: Optional[List[Any]]=None, multi_rules: Optional[List[Dict[str, Any]]]=None) -> None:
    """
    Сохраняет результаты аудита в форматированный Excel.

    Показывает ВСЕ правила: ОК если нарушений нет, FAIL с деталями если есть.
    Колонка «Тип проверки» разделяет базовые и методические правила.

    Args:
        violations: список нарушений (list of dicts, каждый с полем "layer")
        output_path: путь для сохранения .xlsx
        all_rules: список всех RuleSpec (legacy формат)
        multi_rules: список dict-правил из rules_multi.json + rules_methodology.json
    """
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = 'Результаты проверки'
    for col_idx, header in enumerate(_DISPLAY_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _CENTER_ALIGNMENT
        cell.border = _THIN_BORDER
    ws.row_dimensions[1].height = 30
    violations_by_rule = {}
    for v in violations:
        key = (v.get('rule_index', 0), v.get('layer', 'base'))
        violations_by_rule.setdefault(key, []).append(v)
    rows_data = []
    if multi_rules:
        for rule in sorted(multi_rules, key=lambda r: (0 if r.get('layer', 'base') == 'base' else 1, r.get('index', 0))):
            idx = rule.get('index', 0)
            layer = rule.get('layer', 'base')
            key = (idx, layer)
            rule_violations = violations_by_rule.get(key, [])
            if rule_violations:
                for v in rule_violations:
                    rows_data.append({'index': idx, 'title': rule.get('title', ''), 'layer': layer, 'status': 'FAIL', 'target': v.get('Целевой документ', ''), 'diff': v.get('Различие', '')})
            else:
                rows_data.append({'index': idx, 'title': rule.get('title', ''), 'layer': layer, 'status': 'ОК', 'target': '', 'diff': ''})
    elif all_rules:
        for rule in sorted(all_rules, key=lambda r: r.index):
            key = (rule.index, 'base')
            rule_violations = violations_by_rule.get(key, [])
            if rule_violations:
                for v in rule_violations:
                    rows_data.append({'index': rule.index, 'title': rule.title, 'layer': 'base', 'status': 'FAIL', 'target': v.get('Целевой документ', ''), 'diff': v.get('Различие', '')})
            else:
                rows_data.append({'index': rule.index, 'title': rule.title, 'layer': 'base', 'status': 'ОК', 'target': '', 'diff': ''})
    else:
        for v in sorted(violations, key=lambda v: v.get('rule_index', 0)):
            rows_data.append({'index': v.get('rule_index', ''), 'title': v.get('правило', ''), 'layer': v.get('layer', 'base'), 'status': 'FAIL', 'target': v.get('Целевой документ', ''), 'diff': v.get('Различие', '')})
    for row_idx, row in enumerate(rows_data, start=2):
        is_ok = row['status'] == 'ОК'
        is_meth = row.get('layer') == 'methodology'
        if is_meth:
            fill = _METH_OK_FILL if is_ok else _METH_FAIL_FILL
        else:
            fill = _OK_FILL if is_ok else _FAIL_FILL
        values = [row['index'], row['title'], _layer_label(row.get('layer', 'base')), row['status'], row['target'], row['diff']]
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = _CELL_FONT
            cell.border = _THIN_BORDER
            cell.fill = fill
            if col_idx == 1:
                cell.alignment = _CENTER_ALIGNMENT
            elif col_idx in (3, 4):
                cell.alignment = _CENTER_ALIGNMENT
                if col_idx == 4:
                    cell.font = _OK_FONT if is_ok else _FAIL_FONT
            else:
                cell.alignment = _WRAP_ALIGNMENT
    for col_idx in range(1, len(_DISPLAY_HEADERS) + 1):
        max_len = len(str(ws.cell(row=1, column=col_idx).value))
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value:
                    lines = str(cell.value).split('\n')
                    longest = max((len(line) for line in lines))
                    max_len = max(max_len, longest)
        min_w = _MIN_WIDTHS[col_idx - 1]
        max_w = _MAX_WIDTHS[col_idx - 1]
        width = min(max(max_len + 2, min_w), max_w)
        ws.column_dimensions[get_column_letter(col_idx)].width = width
    ws.freeze_panes = 'A2'
    wb.save(output_path)

excel_reporter_module = SimpleNamespace(_DISPLAY_HEADERS=_DISPLAY_HEADERS, _FONT_SIZE=_FONT_SIZE, _HEADER_FONT=_HEADER_FONT, _HEADER_FILL=_HEADER_FILL, _CELL_FONT=_CELL_FONT, _WRAP_ALIGNMENT=_WRAP_ALIGNMENT, _CENTER_ALIGNMENT=_CENTER_ALIGNMENT, _THIN_BORDER=_THIN_BORDER, _MIN_WIDTHS=_MIN_WIDTHS, _MAX_WIDTHS=_MAX_WIDTHS, _OK_FILL=_OK_FILL, _FAIL_FILL=_FAIL_FILL, _OK_FONT=_OK_FONT, _FAIL_FONT=_FAIL_FONT, _METH_OK_FILL=_METH_OK_FILL, _METH_FAIL_FILL=_METH_FAIL_FILL, _layer_label=_layer_label, save_to_excel=save_to_excel)

# END_SOURCE_EXCEL_REPORTER
# END_REPORTING

# START_RUNTIME_SERVICES
# PURPOSE: Orchestrate the generic audit engine, special engines, prompt building, and LLM calls.
# INPUTS: doc_type, parsed docs, rule specs, templates, API requests, CLI args.
# OUTPUTS: AuditResult objects, session directories, and engine dispatch decisions.
# KEYWORDS: runtime, engine, multi-rule, drivers, kpsc, kartochka, plan-grafik.
# LINKS: audit_engine/engine.py, audit_engine/*/__init__.py, audit_engine/multi_rule.py.
# RATIONALE: The monolith should expose one runtime graph instead of scattered module entrypoints.



# START_SOURCE_KPSC_RUN_VALIDATIONS
# PURPOSE: Inlined source from audit_engine/kpsc/run_validations.py.

def kpsc_run_validations_load_rules(rules_path: Path) -> List[Dict]:
    """Загрузка правил из validation_rules.json."""
    with open(rules_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data['rules']

def kpsc_run_validations_get_validator_module_name(rule_index: str) -> str:
    """Конвертирует rule_index в имя модуля валидатора: '1.1' → 'validate_1_1_*'."""
    prefix = rule_index.replace('.', '_')
    scripts_dir = Path(__file__).parent / 'validation_scripts'
    pattern = f'validate_{prefix}_*.py'
    matches = list(scripts_dir.glob(pattern))
    if not matches:
        raise FileNotFoundError(f'Валидатор для правила {rule_index} не найден (паттерн: {pattern})')
    if len(matches) > 1:
        raise ValueError(f'Найдено несколько валидаторов для правила {rule_index}: {matches}')
    module_stem = matches[0].stem
    return f'audit_engine.kpsc.validation_scripts.{module_stem}'

def kpsc_run_validations_create_excel_report(results: List[Tuple[Dict, Dict, float]], report_path: Path):
    """Создание Excel-отчёта с результатами валидации."""
    rows = []
    for rule, result_data, duration in results:
        rows.append({'rule_index': rule['rule_index'], 'section': rule.get('section', ''), 'rule_title': rule.get('rule_title', ''), 'status': result_data.get('status', 'UNKNOWN'), 'discrepancy': result_data.get('discrepancy', ''), 'duration_sec': round(duration, 2)})
    df = pd.DataFrame(rows)
    df = df.sort_values('rule_index')
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Validation Results', index=False)
        worksheet = writer.sheets['Validation Results']
        for idx, col in enumerate(df.columns):
            max_length = max(df[col].astype(str).apply(len).max(), len(col))
            from openpyxl.utils import get_column_letter
            worksheet.column_dimensions[get_column_letter(idx + 1)].width = min(max_length + 2, 50)
    return df


# END_SOURCE_KPSC_RUN_VALIDATIONS

# START_SOURCE_KPSC
# PURPOSE: Inlined source from audit_engine/kpsc/__init__.py.
def kpsc__load_config() -> Dict[str, Any]:
    """Загрузка конфига из doc_configs/kpsc/config.json."""
    config_path = Path(__file__).resolve().parents[2] / 'doc_configs' / 'kpsc' / 'config.json'
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def kpsc_run(args) -> AuditResult:
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
    config = kpsc__load_config()
    max_workers = config.get('max_workers', 5)
    rules_path = Path(__file__).resolve().parents[2] / 'doc_configs' / 'kpsc' / 'validation_rules.json'
    if args.session_dir:
        session_dir = Path(args.session_dir)
    else:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        session_dir = Path(__file__).resolve().parents[2] / 'logs_result' / 'kpsc' / f'session_{timestamp}'
    session_dir.mkdir(parents=True, exist_ok=True)
    parser_outputs_dir = session_dir / 'parser_outputs'
    validation_outputs_dir = session_dir / 'validation_outputs'
    print(f'\n{'=' * 60}')
    print(f'  Валидация КПСЦ')
    print(f'{'=' * 60}')
    print(f'  Файл: {target_path.name}')
    print(f'  Сессия: {session_dir}')
    print(f'  Параллельность: {max_workers}')
    print()
    os.environ['VALIDATION_RULES_PATH'] = str(rules_path)
    os.environ['LLM_BASE_URL'] = config.get('llm_base_url', LLM_CONFIG.base_url)
    os.environ['LLM_MODEL'] = config.get('model', LLM_CONFIG.default_model)
    if not os.environ.get('OPENAI_API_KEY'):
        os.environ['OPENAI_API_KEY'] = 'dummy'
    print('[1/3] Запуск 9 парсеров...')
    parser_results = kpsc_run_validations_run_parsers(target_path, parser_outputs_dir)
    ok_count = sum((1 for r in parser_results.values() if r['status'] == 'ok'))
    err_count = sum((1 for r in parser_results.values() if r['status'] == 'error'))
    print(f'\n  Парсинг завершён: {ok_count} OK, {err_count} ошибок')
    with open(session_dir / 'parser_summary.json', 'w', encoding='utf-8') as f:
        json.dump(parser_results, f, ensure_ascii=False, indent=2)
    if args.parse_only:
        print(f'\n{'=' * 60}')
        print('Режим --parse-only: парсинг завершён')
        print(f'Результаты в: {parser_outputs_dir}')
        for p in sorted(parser_outputs_dir.glob('*.json')):
            print(f'  {p.name}')
        return AuditResult(doc_type='kpsc', session_dir=session_dir, target_path=str(target_path), duration_sec=time.time() - start_time)
    print('\n[2/3] Загрузка правил...')
    rules = kpsc_run_validations_load_rules(rules_path)
    if args.rule_filter:
        rule_filter_str = str(args.rule_filter)
        rules = [r for r in rules if r['rule_index'] == rule_filter_str]
        if not rules:
            raise ValueError(f'Правило {rule_filter_str} не найдено в kpsc')
        print(f'  Фильтр: только правило {rule_filter_str}')
    print(f'  Правил к проверке: {len(rules)}')
    print()
    results = kpsc_run_validations_run_validators_parallel(rules=rules, parser_outputs_dir=parser_outputs_dir, output_dir=validation_outputs_dir, rules_path=rules_path, max_workers=max_workers)
    print(f'\n[3/3] Генерация отчёта...')
    report_path = session_dir / 'validation_report.xlsx'
    df = kpsc_run_validations_create_excel_report(results, report_path)
    print(f'  Отчёт: {report_path}')
    total = len(df)
    passed = len(df[df['status'] == 'PASS'])
    failed = len(df[df['status'] == 'FAIL'])
    errors = total - passed - failed
    duration = time.time() - start_time
    print(f'\n{'=' * 60}')
    print(f'  Статистика:')
    print(f'    Всего правил: {total}')
    print(f'    PASS: {passed} ({passed / total * 100:.1f}%)' if total > 0 else '')
    print(f'    FAIL: {failed} ({failed / total * 100:.1f}%)' if total > 0 else '')
    if errors > 0:
        print(f'    ERRORS: {errors} ({errors / total * 100:.1f}%)')
    print(f'  Время: {duration:.1f} сек')
    print(f'  Сессия: {session_dir}')
    print(f'{'=' * 60}')
    violations = []
    for rule, result_data, dur in results:
        if result_data.get('status') == 'FAIL':
            violations.append({'rule_index': result_data.get('rule_index', rule.get('rule_index', '?')), 'rule_title': result_data.get('rule_title', rule.get('rule_title', '')), 'section': rule.get('section', ''), 'discrepancy': result_data.get('discrepancy', '')})
    if violations:
        print(f'\n  Нарушения ({len(violations)}):')
        for v in violations:
            print(f'  - [{v['rule_index']}] {v['rule_title']}: {v['discrepancy'][:80]}')
    return AuditResult(violations=violations, doc_type='kpsc', session_dir=session_dir, duration_sec=duration, rules_checked=total, target_path=str(target_path))

kpsc_module = SimpleNamespace(_load_config=kpsc__load_config, run=kpsc_run)

# END_SOURCE_KPSC

# START_SOURCE_KARTOCHKA_PROEKTA_RUN_VALIDATIONS
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/run_validations.py.

def kartochka_proekta_run_validations_load_rules(rules_path: Path) -> List[Dict]:
    """Загрузка правил из validation_rules.json."""
    with open(rules_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data['rules']

def kartochka_proekta_run_validations_create_excel_report(results: List[Tuple[Dict, Dict, float]], report_path: Path):
    """Создание Excel-отчёта с результатами валидации."""
    rows = []
    for rule, result_data, duration in results:
        rows.append({'rule_index': rule['rule_index'], 'section': rule.get('section', ''), 'rule_title': rule.get('rule_title', ''), 'status': result_data.get('status', 'UNKNOWN'), 'discrepancy': result_data.get('discrepancy', ''), 'duration_sec': round(duration, 2)})
    df = pd.DataFrame(rows)
    df = df.sort_values('rule_index')
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Validation Results', index=False)
        worksheet = writer.sheets['Validation Results']
        for idx, col in enumerate(df.columns):
            max_length = max(df[col].astype(str).apply(len).max(), len(col))
            from openpyxl.utils import get_column_letter
            worksheet.column_dimensions[get_column_letter(idx + 1)].width = min(max_length + 2, 50)
    return df


# END_SOURCE_KARTOCHKA_PROEKTA_RUN_VALIDATIONS

# START_SOURCE_KARTOCHKA_PROEKTA
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/__init__.py.
def kartochka_proekta__load_config() -> Dict[str, Any]:
    """Загрузка конфига из doc_configs/kartochka_proekta/config.json."""
    config_path = Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'config.json'
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_run(args) -> AuditResult:
    """
    Запуск валидации «Карточка проекта».

    Вход: args (argparse Namespace) с полями:
        - target: путь к XLSX файлу
        - parse_only: только парсинг (3 парсера)
        - rule_filter: проверить только указанное правило (например '9')
        - model: переопределить модель
        - temperature: переопределить температуру
        - session_dir: директория для логов
    Выход: AuditResult с violations
    """
    start_time = time.time()
    target_path = Path(args.target)
    config = kartochka_proekta__load_config()
    max_workers = config.get('max_workers', 5)
    rules_path = Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json'
    if args.session_dir:
        session_dir = Path(args.session_dir)
    else:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        session_dir = Path(__file__).resolve().parents[2] / 'logs_result' / 'kartochka_proekta' / f'session_{timestamp}'
    session_dir.mkdir(parents=True, exist_ok=True)
    parser_outputs_dir = session_dir / 'parser_outputs'
    validation_outputs_dir = session_dir / 'validation_outputs'
    print(f'\n{'=' * 60}')
    print(f'  Валидация «Карточка проекта»')
    print(f'{'=' * 60}')
    print(f'  Файл: {target_path.name}')
    print(f'  Сессия: {session_dir}')
    print(f'  Параллельность: {max_workers}')
    print()
    os.environ['VALIDATION_RULES_PATH'] = str(rules_path)
    print('[1/3] Запуск 3 парсеров...')
    parser_results = kartochka_proekta_run_validations_run_parsers(target_path, parser_outputs_dir)
    ok_count = sum((1 for r in parser_results.values() if r['status'] == 'ok'))
    err_count = sum((1 for r in parser_results.values() if r['status'] == 'error'))
    print(f'\n  Парсинг завершён: {ok_count} OK, {err_count} ошибок')
    with open(session_dir / 'parser_summary.json', 'w', encoding='utf-8') as f:
        json.dump(parser_results, f, ensure_ascii=False, indent=2)
    if args.parse_only:
        print(f'\n{'=' * 60}')
        print('Режим --parse-only: парсинг завершён')
        print(f'Результаты в: {parser_outputs_dir}')
        for p in sorted(parser_outputs_dir.glob('*.json')):
            print(f'  {p.name}')
        return AuditResult(doc_type='kartochka_proekta', session_dir=session_dir, target_path=str(target_path), duration_sec=time.time() - start_time)
    print('\n[2/3] Загрузка правил...')
    rules = kartochka_proekta_run_validations_load_rules(rules_path)
    if args.rule_filter:
        rule_filter_str = str(args.rule_filter)
        rules = [r for r in rules if r['rule_index'] == rule_filter_str]
        if not rules:
            raise ValueError(f'Правило {rule_filter_str} не найдено в kartochka_proekta')
        print(f'  Фильтр: только правило {rule_filter_str}')
    print(f'  Правил к проверке: {len(rules)}')
    print()
    results = kartochka_proekta_run_validations_run_validators_parallel(rules=rules, parser_outputs_dir=parser_outputs_dir, output_dir=validation_outputs_dir, rules_path=rules_path, max_workers=max_workers)
    print(f'\n[3/3] Генерация отчёта...')
    report_path = session_dir / 'validation_report.xlsx'
    df = kartochka_proekta_run_validations_create_excel_report(results, report_path)
    print(f'  Отчёт: {report_path}')
    total = len(df)
    passed = len(df[df['status'] == 'PASS'])
    failed = len(df[df['status'] == 'FAIL'])
    errors = total - passed - failed
    duration = time.time() - start_time
    print(f'\n{'=' * 60}')
    print(f'  Статистика:')
    print(f'    Всего правил: {total}')
    if total > 0:
        print(f'    PASS: {passed} ({passed / total * 100:.1f}%)')
        print(f'    FAIL: {failed} ({failed / total * 100:.1f}%)')
    if errors > 0:
        print(f'    ERRORS: {errors} ({errors / total * 100:.1f}%)')
    print(f'  Время: {duration:.1f} сек')
    print(f'  Сессия: {session_dir}')
    print(f'{'=' * 60}')
    violations = []
    for rule, result_data, dur in results:
        if result_data.get('status') == 'FAIL':
            violations.append({'rule_index': result_data.get('rule_index', rule.get('rule_index', '?')), 'rule_title': result_data.get('rule_title', rule.get('rule_title', '')), 'section': rule.get('section', ''), 'discrepancy': result_data.get('discrepancy', '')})
    if violations:
        print(f'\n  Нарушения ({len(violations)}):')
        for v in violations:
            print(f'  - [{v['rule_index']}] {v['rule_title']}: {v['discrepancy'][:80]}')
    return AuditResult(violations=violations, doc_type='kartochka_proekta', session_dir=session_dir, duration_sec=duration, rules_checked=total, target_path=str(target_path))

kartochka_proekta_module = SimpleNamespace(_load_config=kartochka_proekta__load_config, run=kartochka_proekta_run)

# END_SOURCE_KARTOCHKA_PROEKTA


# START_SOURCE_ENGINE
# PURPOSE: Inlined source from audit_engine/engine.py.
class AuditEngine:
    """
    Единый движок аудита документов.

    Использование:
        engine = AuditEngine("presentation_eu")
        result = engine.run("document.pptx")

    Или с кастомным base_dir:
        engine = AuditEngine("prikaz_ic", base_dir="/path/to/project")
    """

    def __init__(self, doc_type: str, base_dir: Optional[str]=None, config_dir: Optional[str]=None):
        """
        Инициализация движка.

        Args:
            doc_type: тип документа (имя папки в doc_configs/)
            base_dir: корневая директория проекта (по умолчанию — родитель audit_engine/)
            config_dir: путь к папке конфигов (переопределяет base_dir + doc_configs)
        """
        self.doc_type = doc_type
        if base_dir:
            self.base_dir = Path(base_dir)
        else:
            self.base_dir = Path(__file__).parent.parent
        if config_dir:
            self.config_path = Path(config_dir)
        else:
            self.config_path = self.base_dir / 'doc_configs' / doc_type
        self.config: AuditConfig = load_audit_config(self.config_path)
        self.logger: Optional[PipelineLogger] = None

    def run(self, target_path: str, model: Optional[str]=None, temperature: Optional[float]=None, rule_filter: Optional[int]=None, parse_only: bool=False, print_prompts: bool=False, session_dir: Optional[str]=None, chunk_filter: Optional[str]=None, out_xlsx: Optional[str]=None, secondary_path: Optional[str]=None, progress_callback: Optional[callable]=None) -> AuditResult:
        """
        Запуск полного цикла аудита.

        Args:
            target_path: путь к целевому документу
            model: модель LLM (переопределяет конфиг)
            temperature: температура (переопределяет конфиг)
            rule_filter: проверить только одно правило
            parse_only: режим только парсинга
            print_prompts: режим отладки — выводить промпты без LLM
            session_dir: директория для логов
            chunk_filter: парсить только указанный чанк
            out_xlsx: путь для сохранения Excel
            secondary_path: путь к вторичному файлу (XLSX для multi-file аудитов)
            progress_callback: колбэк прогресса для веб-UI (type: str, data: dict)

        Returns:
            AuditResult с нарушениями и статистикой
        """
        start_time = time.time()

        def _emit(event_type: str, data: dict=None):
            if progress_callback:
                progress_callback(event_type, data or {})
        model = model or self.config.model
        temperature = temperature if temperature is not None else self.config.temperature
        if session_dir:
            session_path = Path(session_dir)
        else:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            session_path = self.base_dir / 'logs_result' / self.doc_type / f'session_{timestamp}'
        self.logger = PipelineLogger(session_path, self.config.doc_title)
        self.logger.log(f'🚀 Запуск аудита: {self.config.doc_title} ({self.doc_type})')
        self.logger.log(f'   Сессия: {session_path}')
        self.logger.log(f'   Целевой документ: {target_path}')
        if secondary_path:
            self.logger.log(f'   Вторичный файл: {secondary_path}')
        self.logger.log(f'   Модель: {model}')
        # Логируем карту парсеров по расширениям; блоки дополнительного контекста
        # выводим, только если соответствующий парсер реально задействован хотя бы для
        # одного расширения.
        self.logger.log(f'   Парсеры по расширениям: {self.config.parser_by_ext}')
        parser_names = set(self.config.parser_by_ext.values())
        if 'paddle' in parser_names:
            # Значения берём из PDF infra-конфига (config/parsers.json), не из AuditConfig:
            # AuditConfig описывает doc_type, а сервисы PDF общие для всего рантайма.
            pdf_cfg = self.pdf_parser_config
            layout_src = pdf_cfg.layout.base_url or f'local: {pdf_cfg.layout.model}'
            self.logger.log(f'   Layout: {layout_src}')
            self.logger.log(f'   VLM: {pdf_cfg.vlm.model or "авто (через /v1/models)"}')
            self.logger.log(f'   VLM URL: {pdf_cfg.vlm.base_url}')
        if self.config.llm_base_url:
            self.logger.log(f'   LLM URL: {self.config.llm_base_url}')
        if not Path(target_path).exists():
            self.logger.log(f'❌ Целевой документ не найден: {target_path}')
            raise FileNotFoundError(f'Целевой документ не найден: {target_path}')
        if self.config.secondary_file and (not secondary_path):
            self.logger.log(f'⚠️ Конфиг требует вторичный файл ({self.config.secondary_file.type}), но --secondary не указан')
        if secondary_path and (not Path(secondary_path).exists()):
            self.logger.log(f'❌ Вторичный файл не найден: {secondary_path}')
            raise FileNotFoundError(f'Вторичный файл не найден: {secondary_path}')
        self.logger.log(f'📋 Загрузка правил...')
        all_rules = load_rules(str(self.config.rules_path))
        self.logger.log(f'   Загружено {len(all_rules)} правил')
        _emit('audit_start', {'doc_type': self.doc_type, 'filename': Path(target_path).name, 'total_rules': len(all_rules)})
        rules = all_rules
        chunks_to_parse = None
        if rule_filter is not None:
            rules = [r for r in all_rules if r.index == rule_filter]
            if not rules:
                self.logger.log(f'❌ Правило #{rule_filter} не найдено')
                raise ValueError(f'Правило #{rule_filter} не найдено в {self.doc_type}')
            rule = rules[0]
            chunks_to_parse = [rule.scope] if isinstance(rule.scope, str) else list(rule.scope)
            self.logger.log(f'   ⚠️ Фильтр: только правило #{rule_filter}')
        if chunk_filter:
            chunks_to_parse = [chunk_filter]
        target_ext = Path(target_path).suffix.lower()
        _is_secondary_upload = False
        if self.config.secondary_file and target_ext == f'.{self.config.secondary_file.type}' and (not secondary_path):
            _is_secondary_upload = True
            self.logger.log(f'📊 Загружен {target_ext} — парсим как вторичный файл')
            _emit('parsing_target', {})
            target_doc = self._parse_secondary(target_path)
            _emit('parsing_target_done', {})
        else:
            _emit('parsing_target', {})
            target_doc = self._parse_document(target_path, chunks_to_parse, session_path / 'vision_target')
            _emit('parsing_target_done', {})
            if secondary_path and self.config.secondary_file:
                secondary_chunks = self._parse_secondary(secondary_path)
                target_doc.update(secondary_chunks)
                self.logger.log(f'   📊 Merged {len(secondary_chunks)} чанков из вторичного файла')
        if parse_only:
            self.logger.log(f'✅ Режим --parse-only: парсинг завершён')
            print(f'\n{'=' * 60}')
            print('TARGET:')
            print(json.dumps(target_doc, ensure_ascii=False, indent=2))
            return AuditResult(doc_type=self.doc_type, session_dir=session_path, target_path=target_path, duration_sec=time.time() - start_time)
        # Активный LLM-путь — multi_rule: парсер отдаёт raw_text, промпт собирается
        # по sections.json + rules_multi.json / rules_methodology.json.
        doc_configs_dir = self.config.config_dir.parent
        mr_config = load_multi_rule_config(doc_configs_dir, self.doc_type)
        if mr_config is None:
            raise ValueError(
                f"Для doc_type={self.doc_type} не найдены sections.json + rules_multi.json. "
                f"Без них multi-rule аудит запустить нельзя; legacy single-rule путь удалён."
            )
        self.logger.log(f'🔍 Запуск multi-rule аудита (базовый слой, {len(mr_config['rules'])} правил)...')
        _emit('checking_rules', {'total': len(mr_config['rules'])})
        mr_result = run_multi_rule_audit(
            parsed=target_doc,
            sections=mr_config['sections'],
            rules=mr_config['rules'],
            include_scopes=mr_config.get('include_scopes'),
            filename=target_doc.get('filename', Path(target_path).name),
            llm_base_url=self.config.llm_base_url,
            llm_model=self.config.model,
            session_dir=session_path,
            layer='base',
        )
        violations = mr_result['violations']
        self.logger.log(f'   Base usage: prompt={mr_result['usage']['prompt_tokens']} completion={mr_result['usage']['completion_tokens']} total={mr_result['usage']['total_tokens']}')
        meth_config = load_methodology_config(doc_configs_dir, self.doc_type)
        if meth_config is not None:
            self.logger.log(f'🔍 Запуск методического слоя ({len(meth_config['rules'])} правил, источник: {meth_config.get('source', 'МР/МУ')[:80]})...')
            meth_result = run_multi_rule_audit(
                parsed=target_doc,
                sections=mr_config['sections'],
                rules=meth_config['rules'],
                include_scopes=meth_config.get('include_scopes') or mr_config.get('include_scopes'),
                filename=target_doc.get('filename', Path(target_path).name),
                llm_base_url=self.config.llm_base_url,
                llm_model=self.config.model,
                session_dir=session_path,
                layer='methodology',
            )
            violations.extend(meth_result['violations'])
            self.logger.log(f'   Methodology usage: prompt={meth_result['usage']['prompt_tokens']} completion={meth_result['usage']['completion_tokens']} total={meth_result['usage']['total_tokens']}')
            self.logger.log(f'   Methodology violations: {len(meth_result['violations'])}')
        _all_multi_rules = [{**r, 'layer': 'base'} for r in mr_config['rules']]
        if meth_config is not None:
            _all_multi_rules.extend(({**r, 'layer': 'methodology'} for r in meth_config['rules']))
        _emit('checking_rules_done', {'violations': len(violations)})
        self.logger.log_final_results(violations)
        print(json.dumps(violations, ensure_ascii=False, indent=2))
        if out_xlsx:
            xlsx_path = out_xlsx
        else:
            xlsx_path = str(session_path / 'audit_result.xlsx')
        save_to_excel(violations, xlsx_path, all_rules=None, multi_rules=_all_multi_rules)
        self.logger.log(f'📊 Excel сохранён: {xlsx_path}')
        duration = time.time() - start_time
        self.logger.log(f'{'=' * 60}')
        self.logger.log(f'📊 Итого нарушений: {len(violations)}')
        if violations:
            by_rule: Dict[Any, int] = {}
            for v in violations:
                idx = v.get('rule_index', '?')
                by_rule.setdefault(idx, 0)
                by_rule[idx] += 1
            self.logger.log('   По правилам:')
            for idx in sorted(by_rule.keys()):
                self.logger.log(f'   - Правило #{idx}: {by_rule[idx]} нарушений')
        self.logger.log(f'✅ Аудит завершён за {duration:.1f} сек. Сессия: {session_path}')
        return AuditResult(violations=violations, doc_type=self.doc_type, session_dir=session_path, duration_sec=duration, rules_checked=len(_all_multi_rules), target_path=target_path)

    def _parse_document(self, file_path: str, chunks_to_parse: Optional[List[str]], vision_log_dir: Path) -> Dict[str, Any]:
        """
        Парсит документ выбранным парсером.

        Выбор парсера: `config.parser_by_ext[<расширение файла>]`. Никаких
        автоматических переопределений — один конфиг описывает всю цепочку.
        Если расширение не описано в карте, бросается ValueError.

        Параметр `chunks_to_parse` оставлен в сигнатуре для обратной совместимости
        с вызывающим кодом, но не используется: форматные парсеры теперь всегда
        отдают весь текст документа целиком (`ParsedDocument` с `raw_text`).
        """
        del chunks_to_parse  # параметр больше не нужен формат-парсерам
        file_ext = Path(file_path).suffix.lower()
        # Единый источник правды — parser_by_ext. Никаких скрытых override по формату.
        if file_ext not in self.config.parser_by_ext:
            raise ValueError(f"Для doc_type={self.doc_type} не сконфигурирован парсер для расширения {file_ext!r}. parser_by_ext={self.config.parser_by_ext}")
        parser_name = self.config.parser_by_ext[file_ext]
        if parser_name == 'docx':
            self.logger.log(f'📄 DOCX-парсинг: {file_path}...')
            # Один и тот же VLMClient используется и для PDF-страниц, и для OCR шапок DOCX.
            vlm_cfg = self.pdf_parser_config.vlm
            vlm_client = VLMClient(
                base_url=vlm_cfg.base_url,
                model=vlm_cfg.model,
                api_key=vlm_cfg.api_key,
                max_tokens=vlm_cfg.max_tokens,
                temperature=vlm_cfg.temperature,
            )
            doc = parse_docx(file_path, vlm=vlm_client, header_ocr_cfg=self.docx_header_ocr_config)
        elif parser_name == 'pptx':
            self.logger.log(f'📄 PPTX-парсинг: {file_path}...')
            doc = parse_pptx(file_path)
        elif parser_name == 'paddle':
            self.logger.log(f'📄 PDF-парсинг (Paddle): {file_path}...')
            doc = parse_pdf(file_path, self.pdf_parser_config, log_dir=vision_log_dir)
        else:
            # Единственные поддерживаемые парсеры сейчас — docx/pptx/paddle. Любое другое
            # значение — ошибка конфигурации и причина явно падать, а не молча продолжать.
            raise ValueError(f"Неподдерживаемый парсер {parser_name!r} для {file_ext!r}. Допустимы: docx, pptx, paddle.")
        self.logger.log_parsed_doc(doc, Path(file_path).stem)
        return doc

    def _parse_secondary(self, file_path: str) -> Dict[str, Any]:
        """
        Парсит вторичный файл через зарегистрированный парсер.

        Добавляет префикс из config.secondary_file.chunk_prefix к ключам,
        а также ключ <prefix>filename с именем файла.
        """
        from .parsers import get_parser
        spec = self.config.secondary_file
        parser_fn = get_parser(spec.parser)
        raw_chunks = parser_fn(file_path)
        prefixed = {f'{spec.chunk_prefix}{k}': v for k, v in raw_chunks.items()}
        prefixed[f'{spec.chunk_prefix}filename'] = Path(file_path).name
        self.logger.log_parsed_doc(prefixed, f'secondary_{spec.type}')
        return prefixed

    @staticmethod
    def list_doc_types(base_dir: Optional[str]=None) -> List[Dict[str, str]]:
        """
        Список доступных типов документов.

        Сканирует doc_configs/ и возвращает список {doc_type, doc_title}.
        """
        if base_dir:
            configs_dir = Path(base_dir) / 'doc_configs'
        else:
            configs_dir = Path(__file__).parent.parent / 'doc_configs'
        result = []
        if not configs_dir.exists():
            return result
        for config_dir in sorted(configs_dir.iterdir()):
            config_file = config_dir / 'config.json'
            if config_file.exists():
                try:
                    with open(config_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    result.append({'doc_type': data.get('doc_type', config_dir.name), 'doc_title': data.get('doc_title', '')})
                except (json.JSONDecodeError, KeyError):
                    result.append({'doc_type': config_dir.name, 'doc_title': '(ошибка чтения config.json)'})
        return result

engine_module = SimpleNamespace(AuditEngine=AuditEngine)

# END_SOURCE_ENGINE
# END_RUNTIME_SERVICES

# START_API
# PURPOSE: Expose the backend over FastAPI with auth, health, audit, SSE, and download routes.
# INPUTS: HTTP requests, uploaded files, session ids.
# OUTPUTS: JSON responses, SSE events, file downloads.
# KEYWORDS: fastapi, health, login, audit, sse.
# LINKS: api_server.py.
# RATIONALE: `uvicorn main:app` must be the active API entrypoint after the collapse.

# START_SOURCE_API_SERVER
# PURPOSE: Inlined source from api_server.py.
faulthandler.enable(file=sys.stderr)

try:
    import torch
    if torch.cuda.is_available():
        torch.cuda.init()
        _dummy = torch.zeros(1, device='cuda:0')
        del _dummy
        print(f'[CUDA] Контекст зарезервирован на {torch.cuda.get_device_name(0)}')
    else:
        print('[CUDA] GPU не обнаружен — пропуск CUDA warmup (модели на удалённом сервере)')
except Exception as e:
    print(f'[CUDA] Не удалось зарезервировать контекст: {e}')

app = FastAPI(title='AuditAPI')

_audit_lock = threading.Lock()

_queue_counter = 0

_queue_counter_lock = threading.Lock()

app.add_middleware(CORSMiddleware, allow_origins=['http://localhost:5173', 'http://192.168.20.118:5173'], allow_credentials=True, allow_methods=['*'], allow_headers=['*'])

AUTH_LOGIN = os.getenv('AUDIT_LOGIN', 'admin')

AUTH_PASSWORD = os.getenv('AUDIT_PASSWORD', 'mos186124kva')

AUTH_TOKEN = os.getenv('AUDIT_TOKEN', 'audit-session-token-2026')

UPLOAD_DIR = Path('uploads')

UPLOAD_DIR.mkdir(exist_ok=True)

sessions: Dict[str, Dict[str, Any]] = {}

class LoginRequest(BaseModel):
    """Тело запроса авторизации."""
    login: str
    password: str

def _check_auth(request: Request):
    """Проверка авторизации через cookie или header."""
    token = request.cookies.get('auth_token')
    if not token:
        auth_header = request.headers.get('Authorization', '')
        if auth_header.startswith('Bearer '):
            token = auth_header[7:]
    if token != AUTH_TOKEN:
        raise HTTPException(status_code=401, detail='Не авторизован')

@app.get('/api/health')
async def health_check():
    """Проверка здоровья всех сервисов. Без авторизации."""
    from datetime import datetime
    from urllib.request import urlopen
    from urllib.error import URLError
    services = {}
    try:
        r = urlopen('http://172.16.10.35:11438/v1/models', timeout=5)
        services['paddleocr'] = {'status': 'ok', 'code': r.status}
    except Exception as e:
        services['paddleocr'] = {'status': 'error', 'detail': str(e)}
    try:
        r = urlopen('http://172.16.10.35:11437/v1/models', timeout=5)
        services['llm'] = {'status': 'ok', 'code': r.status}
    except Exception as e:
        services['llm'] = {'status': 'error', 'detail': str(e)}
    try:
        import torch as _torch
        cuda_ok = _torch.cuda.is_available()
        services['cuda'] = {'status': 'ok' if cuda_ok else 'skip', 'device': _torch.cuda.get_device_name(0) if cuda_ok else 'нет GPU (модели на удалённом сервере)'}
    except ImportError:
        services['cuda'] = {'status': 'skip', 'device': 'torch не установлен (модели на удалённом сервере)'}
    except Exception as e:
        services['cuda'] = {'status': 'error', 'detail': str(e)}
    overall = all((s['status'] in ('ok', 'skip') for s in services.values()))
    return {'status': 'ok' if overall else 'degraded', 'services': services, 'timestamp': datetime.now().isoformat()}

@app.post('/api/login')
async def login(body: LoginRequest):
    """Авторизация. Возвращает token и ставит cookie."""
    valid_logins = {AUTH_LOGIN, 'guest'}
    if body.login not in valid_logins or body.password != AUTH_PASSWORD:
        raise HTTPException(status_code=401, detail='Неверный логин или пароль')
    from fastapi.responses import JSONResponse
    response = JSONResponse({'ok': True, 'token': AUTH_TOKEN})
    response.set_cookie('auth_token', AUTH_TOKEN, httponly=True, max_age=86400, samesite='lax')
    return response

@app.get('/api/types')
async def get_types(request: Request):
    """Список доступных типов документов."""
    _check_auth(request)
    return AuditEngine.list_doc_types()

@app.post('/api/audit')
async def start_audit(request: Request, file: UploadFile=File(...), doc_type: str=Form(...)):
    """Запуск аудита: принимает файл + тип, возвращает session_id."""
    _check_auth(request)
    file_ext = Path(file.filename).suffix.lower() if file.filename else ''
    allowed = _get_allowed_extensions(doc_type)
    if file_ext not in allowed:
        ext_list = ', '.join(sorted(allowed))
        raise HTTPException(status_code=400, detail=f'Неподдерживаемый формат файла «{file_ext}». Для типа «{doc_type}» допустимы: {ext_list}')
    session_id = uuid.uuid4().hex[:8]
    upload_path = UPLOAD_DIR / session_id / file.filename
    upload_path.parent.mkdir(parents=True, exist_ok=True)
    with open(upload_path, 'wb') as f:
        shutil.copyfileobj(file.file, f)
    event_queue = asyncio.Queue()
    loop = asyncio.get_event_loop()
    sessions[session_id] = {'status': 'running', 'events': event_queue, 'loop': loop, 'result': None, 'doc_type': doc_type, 'filename': file.filename, 'upload_path': str(upload_path), 'session_dir': None}
    thread = threading.Thread(target=_run_audit_thread, args=(session_id, doc_type, str(upload_path)), daemon=True)
    thread.start()
    return {'session_id': session_id}

@app.get('/api/audit/{session_id}/events')
async def audit_events(session_id: str):
    """SSE-стрим прогресса аудита."""
    session = sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail='Сессия не найдена')

    async def event_generator():
        queue = session['events']
        while True:
            try:
                event = await asyncio.wait_for(queue.get(), timeout=300)
            except asyncio.TimeoutError:
                yield {'event': 'ping', 'data': '{}'}
                continue
            yield {'event': event['type'], 'data': json.dumps(event['data'], ensure_ascii=False)}
            if event['type'] in ('complete', 'error'):
                break
    return EventSourceResponse(event_generator())

@app.get('/api/audit/{session_id}/result')
async def get_result(request: Request, session_id: str):
    """Получить результат аудита (violations + статистика)."""
    _check_auth(request)
    session = sessions.get(session_id)
    if not session:
        raise HTTPException(status_code=404, detail='Сессия не найдена')
    if session['status'] == 'running':
        raise HTTPException(status_code=202, detail='Аудит ещё выполняется')
    if session['status'] == 'error':
        raise HTTPException(status_code=500, detail='Аудит завершился с ошибкой')
    return session['result']

@app.get('/api/audit/{session_id}/download')
async def download_report(request: Request, session_id: str):
    """Скачать Excel-отчёт."""
    _check_auth(request)
    session = sessions.get(session_id)
    if not session or not session.get('session_dir'):
        raise HTTPException(status_code=404, detail='Отчёт не найден')
    session_dir = Path(session['session_dir'])
    xlsx_path = session_dir / 'audit_result.xlsx'
    if not xlsx_path.exists():
        xlsx_files = list(session_dir.glob('*.xlsx'))
        if not xlsx_files:
            raise HTTPException(status_code=404, detail='Excel не сформирован')
        xlsx_path = xlsx_files[0]
    return FileResponse(str(xlsx_path), filename=f'audit_{session_id}.xlsx', media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def api_server__detect_engine(doc_type: str) -> str:
    """
    Определяет тип движка по полю "engine" в config.json.

    Returns:
        str: значение поля "engine" или "standard" если поля нет.
    """
    config_path = Path(__file__).parent / 'doc_configs' / doc_type / 'config.json'
    if config_path.exists():
        with open(config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data.get('engine', 'standard')
    return 'standard'

_STANDARD_EXTENSIONS = {'.docx', '.doc', '.pdf', '.odt', '.rtf'}

_XLSX_EXTENSIONS = {'.xlsx'}

_PPTX_EXTENSIONS = {'.pptx'}

def _get_allowed_extensions(doc_type: str) -> set:
    """
    Возвращает допустимые расширения файлов для данного doc_type.

    Логика:
      - Спецдвижки (drivers, kpsc, kartochka_proekta) → .xlsx
      - parser: "pptx" → .pptx
      - secondary_file → .docx/.doc/.pdf + .xlsx
      - Остальные (paddle) → .docx/.doc/.pdf/.odt/.rtf

    Args:
        doc_type: тип документа
    Returns:
        set расширений (с точкой, нижний регистр)
    """
    config_path = Path(__file__).parent / 'doc_configs' / doc_type / 'config.json'
    if not config_path.exists():
        return _STANDARD_EXTENSIONS
    with open(config_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    engine = data.get('engine', 'standard')
    if engine in api_server_SPECIAL_ENGINES:
        return _XLSX_EXTENSIONS
    parser = data.get('parser', 'paddle')
    if parser == 'pptx':
        return _PPTX_EXTENSIONS
    allowed = set(_STANDARD_EXTENSIONS)
    if data.get('secondary_file'):
        sec_type = data['secondary_file'].get('type', '')
        if sec_type:
            allowed.add(f'.{sec_type}')
    return allowed

api_server_SPECIAL_ENGINES = {'drivers': 'audit_engine.drivers', 'kpsc': 'audit_engine.kpsc', 'kartochka_proekta': 'audit_engine.kartochka_proekta', 'plan_grafik': 'audit_engine.plan_grafik'}

def _run_special_engine(engine_type: str, doc_type: str, target_path: str, session_dir: str) -> 'AuditResult':
    """
    Запуск спецдвижка (kpsc, drivers, kartochka_proekta).

    Спецдвижки ожидают argparse Namespace — эмулируем его.

    Args:
        engine_type: ключ из SPECIAL_ENGINES
        doc_type: тип документа
        target_path: путь к файлу
        session_dir: директория для логов
    Returns:
        AuditResult
    """
    import importlib
    from types import SimpleNamespace
    module = importlib.import_module(api_server_SPECIAL_ENGINES[engine_type])
    args = SimpleNamespace(target=target_path, parse_only=False, rule_filter=None, model=None, temperature=None, session_dir=session_dir)
    return module.run(args)

def _run_audit_thread(session_id: str, doc_type: str, target_path: str):
    """Запуск аудита в фоновом потоке с очередью и progress_callback."""
    global _queue_counter
    session = sessions[session_id]
    loop = session['loop']
    queue = session['events']

    def progress_callback(event_type: str, data: dict):
        """Колбэк из engine — пушит события в SSE-очередь."""
        asyncio.run_coroutine_threadsafe(queue.put({'type': event_type, 'data': data}), loop)
    with _queue_counter_lock:
        _queue_counter += 1
        position = _queue_counter
    if position > 1:
        progress_callback('queue', {'position': position - 1})
        print(f'[QUEUE] Сессия {session_id} встала в очередь, позиция {position - 1}')
    try:
        with _audit_lock:
            progress_callback('queue', {'position': 0})
            try:
                engine_type = api_server__detect_engine(doc_type)
                if engine_type in api_server_SPECIAL_ENGINES:
                    progress_callback('audit_start', {'doc_type': doc_type, 'filename': Path(target_path).name, 'engine': engine_type})
                    from datetime import datetime as dt
                    timestamp = dt.now().strftime('%Y%m%d_%H%M%S')
                    session_dir = str(Path(__file__).parent / 'logs_result' / doc_type / f'session_{timestamp}')
                    result = _run_special_engine(engine_type, doc_type, target_path, session_dir)
                else:
                    engine = AuditEngine(doc_type)
                    result = engine.run(target_path, progress_callback=progress_callback)
                session['result'] = {'violations': result.violations, 'rules_checked': result.rules_checked, 'duration_sec': result.duration_sec}
                session['session_dir'] = str(result.session_dir)
                session['status'] = 'done'
                progress_callback('complete', session['result'])
            except Exception as e:
                err_msg = f'{type(e).__name__}: {e}'
                traceback.print_exc(file=sys.stderr)
                sys.stderr.flush()
                print(f'[AUDIT ERROR] {err_msg}', file=sys.stderr, flush=True)
                session['status'] = 'error'
                progress_callback('error', {'message': err_msg})
    finally:
        try:
            sd = session.get('session_dir')
            if sd:
                orig_dir = Path(sd) / 'original'
                orig_dir.mkdir(exist_ok=True)
                shutil.copy2(target_path, orig_dir / Path(target_path).name)
        except Exception:
            pass
        with _queue_counter_lock:
            _queue_counter -= 1


# END_SOURCE_API_SERVER
# END_API


# START_RUNTIME_INTEGRATION
# PURPOSE: Replace legacy module-string and file-script dispatch with direct monolith callables.
# INPUTS: parser names, validator rule ids, doc_type engine ids, CLI args, API requests.
# OUTPUTS: direct parser callables, direct validator execution, repo-root-safe paths, active CLI/API entrypoints.
# KEYWORDS: integration, dispatch, direct-calls, path-fix, monolith-runtime.
# LINKS: docs/monolith_refactor_journal.md, doc_configs/.
# RATIONALE: The generated monolith must own runtime behavior instead of importing deleted backend files.

SECONDARY_PARSER_DISPATCH = {
    "grafik_obhod": parse_grafik_obhod,
}

KPSC_PARSER_MODULE_DISPATCH = {
    "parse_kpsc_header": parse_kpsc_header,
    "parse_kpsc_table1": parse_kpsc_table1,
    "parse_legend": parse_legend,
    "parse_loss_digitization": parse_loss_digitization,
    "parse_pa1_chart": parse_pa1_chart,
    "parse_pa1_table": parse_pa1_table,
    "parse_pokazateli": parse_pokazateli,
    "parse_spaghetti_sheet": parse_spaghetti_sheet,
    "parse_spaghetti_problems": parse_spaghetti_problems,
}

KARTOCHKA_PARSER_MODULE_DISPATCH = {
    "parse_kartochka_main": parse_kartochka_main,
    "parse_metodika": parse_metodika,
    "parse_dropdown": parse_dropdown,
}



kpsc_run_validations_PARSER_MODULES = list(KPSC_PARSER_MODULE_DISPATCH)
kartochka_proekta_run_validations_PARSER_MODULES = list(KARTOCHKA_PARSER_MODULE_DISPATCH)


def get_parser(name: str) -> Callable[[str], Dict[str, str]]:
    if name not in SECONDARY_PARSER_DISPATCH:
        raise KeyError(f"Парсер '{name}' не зарегистрирован. Доступные: {sorted(SECONDARY_PARSER_DISPATCH)}")
    return SECONDARY_PARSER_DISPATCH[name]






def kpsc_run_validations_run_parsers(xlsx_path: Path, output_dir: Path) -> Dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    results: Dict[str, Any] = {}
    for short_name, parser_fn in KPSC_PARSER_MODULE_DISPATCH.items():
        try:
            parser_fn(xlsx_path, output_dir)
            results[short_name] = {"status": "ok"}
            print(f"  [OK] {short_name}")
        except Exception as e:
            results[short_name] = {"status": "error", "error": str(e)}
            print(f"  [ERR] {short_name}: {e}")
    return results


def kpsc_run_validations_run_single_validator(
    rule: Dict[str, Any],
    parser_outputs_dir: Path,
    output_dir: Path,
    rules_path: Path,
) -> Tuple[Dict[str, Any], Dict[str, Any], float]:
    rule_index = rule["rule_index"]
    start = time.time()
    module_ns = KPSC_VALIDATOR_MODULE_DISPATCH.get(rule_index)
    if module_ns is None:
        return rule, {
            "rule_index": rule_index,
            "status": "MISSING",
            "discrepancy": f"Валидатор для правила {rule_index} не зарегистрирован в main.py",
        }, 0.0

    os.environ["VALIDATION_RULES_PATH"] = str(rules_path)
    output_file = output_dir / f"validate_{rule_index.replace('.', '_')}.json"
    try:
        result_data = run_kpsc_validator_module(module_ns, parser_outputs_dir, output_file)
        return rule, result_data, time.time() - start
    except FileNotFoundError as e:
        result_data = {
            "rule_index": rule_index,
            "status": "FAIL",
            "discrepancy": f"Недостаточно данных парсинга для правила {rule_index}: {e}",
        }
        output_file.parent.mkdir(parents=True, exist_ok=True)
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(result_data, f, ensure_ascii=False, indent=2)
        return rule, result_data, time.time() - start
    except Exception as e:
        return rule, {
            "rule_index": rule_index,
            "status": "ERROR",
            "discrepancy": f"Исключение: {e}",
        }, time.time() - start


def kpsc_run_validations_run_validators_parallel(
    rules: List[Dict[str, Any]],
    parser_outputs_dir: Path,
    output_dir: Path,
    rules_path: Path,
    max_workers: int = 5,
) -> List[Tuple[Dict[str, Any], Dict[str, Any], float]]:
    output_dir.mkdir(parents=True, exist_ok=True)
    results: List[Tuple[Dict[str, Any], Dict[str, Any], float]] = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(
                kpsc_run_validations_run_single_validator,
                rule,
                parser_outputs_dir,
                output_dir,
                rules_path,
            ): rule
            for rule in rules
        }
        for future in as_completed(futures):
            rule = futures[future]
            try:
                rule_data, result_data, duration = future.result()
                results.append((rule_data, result_data, duration))
                status = result_data.get("status", "?")
                emoji = "+" if status == "PASS" else "-" if status == "FAIL" else "!"
                print(f"  [{emoji}] [{len(results)}/{len(rules)}] {result_data.get('rule_index', '?')}: {status} ({duration:.1f}s)")
            except Exception as e:
                results.append((rule, {
                    "rule_index": rule["rule_index"],
                    "status": "ERROR",
                    "discrepancy": f"Future exception: {e}",
                }, 0.0))
    return results


def kartochka_proekta_run_validations_run_parsers(xlsx_path: Path, output_dir: Path) -> Dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    results: Dict[str, Any] = {}
    for short_name, parser_fn in KARTOCHKA_PARSER_MODULE_DISPATCH.items():
        try:
            parser_fn(xlsx_path, output_dir)
            results[short_name] = {"status": "ok"}
            print(f"  [OK] {short_name}")
        except Exception as e:
            results[short_name] = {"status": "error", "error": str(e)}
            print(f"  [ERR] {short_name}: {e}")
    return results


def kartochka_proekta_run_validations_run_single_validator(
    rule: Dict[str, Any],
    parser_outputs_dir: Path,
    output_dir: Path,
    rules_path: Path,
) -> Tuple[Dict[str, Any], Dict[str, Any], float]:
    rule_index = rule["rule_index"]
    start = time.time()
    module_ns = KARTOCHKA_VALIDATOR_MODULE_DISPATCH.get(rule_index)
    if module_ns is None:
        return rule, {
            "rule_index": rule_index,
            "status": "MISSING",
            "discrepancy": f"Валидатор для правила {rule_index} не зарегистрирован в main.py",
        }, 0.0

    os.environ["VALIDATION_RULES_PATH"] = str(rules_path)
    output_file = output_dir / f"validate_{rule_index.replace('.', '_')}.json"
    try:
        result_data = run_kartochka_validator_module(module_ns, parser_outputs_dir, output_file)
        return rule, result_data, time.time() - start
    except Exception as e:
        return rule, {
            "rule_index": rule_index,
            "status": "ERROR",
            "discrepancy": f"Исключение: {e}",
        }, time.time() - start


def kartochka_proekta_run_validations_run_validators_parallel(
    rules: List[Dict[str, Any]],
    parser_outputs_dir: Path,
    output_dir: Path,
    rules_path: Path,
    max_workers: int = 5,
) -> List[Tuple[Dict[str, Any], Dict[str, Any], float]]:
    output_dir.mkdir(parents=True, exist_ok=True)
    results: List[Tuple[Dict[str, Any], Dict[str, Any], float]] = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {
            executor.submit(
                kartochka_proekta_run_validations_run_single_validator,
                rule,
                parser_outputs_dir,
                output_dir,
                rules_path,
            ): rule
            for rule in rules
        }
        for future in as_completed(futures):
            rule = futures[future]
            try:
                rule_data, result_data, duration = future.result()
                results.append((rule_data, result_data, duration))
                status = result_data.get("status", "?")
                emoji = "+" if status == "PASS" else "-" if status == "FAIL" else "!"
                print(f"  [{emoji}] [{len(results)}/{len(rules)}] {result_data.get('rule_index', '?')}: {status} ({duration:.1f}s)")
            except Exception as e:
                results.append((rule, {
                    "rule_index": rule["rule_index"],
                    "status": "ERROR",
                    "discrepancy": f"Future exception: {e}",
                }, 0.0))
    return results





def kpsc__load_config() -> Dict[str, Any]:
    with open(DOC_CONFIGS_DIR / "kpsc" / "config.json", "r", encoding="utf-8") as f:
        return json.load(f)


def _kpsc_core_sheet_missing(parser_results: Dict[str, Dict[str, Any]]) -> bool:
    # A workbook without the "КПСЦ" sheet is not a valid KPSC input.
    required = ("parse_kpsc_header", "parse_kpsc_table1")
    hits = 0
    for parser_name in required:
        parser_state = parser_results.get(parser_name, {})
        error_text = str(parser_state.get("error", ""))
        if parser_state.get("status") == "error" and "Лист КПСЦ не найден" in error_text:
            hits += 1
    return hits == len(required)


def _kpsc_fast_fail_results(
    rules: List[Dict[str, Any]],
    output_dir: Path,
    discrepancy: str,
) -> List[Tuple[Dict[str, Any], Dict[str, Any], float]]:
    output_dir.mkdir(parents=True, exist_ok=True)
    results: List[Tuple[Dict[str, Any], Dict[str, Any], float]] = []
    for rule in rules:
        result_data = {
            "rule_index": rule["rule_index"],
            "rule_title": rule.get("rule_title", ""),
            "status": "FAIL",
            "discrepancy": discrepancy,
        }
        output_file = output_dir / f"validate_{rule['rule_index'].replace('.', '_')}.json"
        with open(output_file, "w", encoding="utf-8") as f:
            json.dump(result_data, f, ensure_ascii=False, indent=2)
        results.append((rule, result_data, 0.0))
    return results


def kartochka_proekta__load_config() -> Dict[str, Any]:
    with open(DOC_CONFIGS_DIR / "kartochka_proekta" / "config.json", "r", encoding="utf-8") as f:
        return json.load(f)



def run_kpsc_special(args) -> AuditResult:
    start_time = time.time()
    target_path = Path(args.target)
    config = kpsc__load_config()
    max_workers = config.get("max_workers", 5)
    rules_path = DOC_CONFIGS_DIR / "kpsc" / "validation_rules.json"
    session_dir = Path(args.session_dir) if args.session_dir else LOGS_RESULT_DIR / "kpsc" / f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    session_dir.mkdir(parents=True, exist_ok=True)
    parser_outputs_dir = session_dir / "parser_outputs"
    validation_outputs_dir = session_dir / "validation_outputs"

    os.environ["VALIDATION_RULES_PATH"] = str(rules_path)
    os.environ["LLM_BASE_URL"] = config.get("llm_base_url", "http://localhost:8001/v1/")
    os.environ["LLM_MODEL"] = resolve_runtime_llm_model(config.get("model", "openai/gpt-oss-120b"))
    os.environ.setdefault("OPENAI_API_KEY", "dummy")

    parser_results = kpsc_run_validations_run_parsers(target_path, parser_outputs_dir)
    with open(session_dir / "parser_summary.json", "w", encoding="utf-8") as f:
        json.dump(parser_results, f, ensure_ascii=False, indent=2)

    if getattr(args, "parse_only", False):
        return AuditResult(doc_type="kpsc", session_dir=session_dir, target_path=str(target_path), duration_sec=time.time() - start_time)

    rules = kpsc_run_validations_load_rules(rules_path)
    if args.rule_filter:
        rules = [r for r in rules if r["rule_index"] == str(args.rule_filter)]
        if not rules:
            raise ValueError(f"Правило {args.rule_filter} не найдено в kpsc")

    if _kpsc_core_sheet_missing(parser_results):
        print("  [WARN] Лист 'КПСЦ' отсутствует. Формируем FAIL-результаты без вызова LLM.")
        results = _kpsc_fast_fail_results(
            rules=rules,
            output_dir=validation_outputs_dir,
            discrepancy="Файл не похож на КПСЦ: отсутствует основной лист 'КПСЦ', поэтому правила КПСЦ завершены как FAIL без LLM-вызова.",
        )
    else:
        results = kpsc_run_validations_run_validators_parallel(
            rules=rules,
            parser_outputs_dir=parser_outputs_dir,
            output_dir=validation_outputs_dir,
            rules_path=rules_path,
            max_workers=max_workers,
        )
    df = kpsc_run_validations_create_excel_report(results, session_dir / "validation_report.xlsx")
    violations = []
    for rule, result_data, _ in results:
        if result_data.get("status") == "FAIL":
            violations.append({
                "rule_index": result_data.get("rule_index", rule.get("rule_index", "?")),
                "rule_title": result_data.get("rule_title", rule.get("rule_title", "")),
                "section": rule.get("section", ""),
                "discrepancy": result_data.get("discrepancy", ""),
            })
    return AuditResult(
        violations=violations,
        doc_type="kpsc",
        session_dir=session_dir,
        duration_sec=time.time() - start_time,
        rules_checked=len(df),
        target_path=str(target_path),
    )


def run_kartochka_proekta_special(args) -> AuditResult:
    start_time = time.time()
    target_path = Path(args.target)
    config = kartochka_proekta__load_config()
    max_workers = config.get("max_workers", 5)
    rules_path = DOC_CONFIGS_DIR / "kartochka_proekta" / "validation_rules.json"
    session_dir = Path(args.session_dir) if args.session_dir else LOGS_RESULT_DIR / "kartochka_proekta" / f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    session_dir.mkdir(parents=True, exist_ok=True)
    parser_outputs_dir = session_dir / "parser_outputs"
    validation_outputs_dir = session_dir / "validation_outputs"

    os.environ["VALIDATION_RULES_PATH"] = str(rules_path)
    os.environ["LLM_BASE_URL"] = config.get("llm_base_url", LLM_CONFIG.base_url)
    os.environ["LLM_MODEL"] = resolve_runtime_llm_model(config.get("model", LLM_CONFIG.default_model))
    os.environ.setdefault("OPENAI_API_KEY", "dummy")
    parser_results = kartochka_proekta_run_validations_run_parsers(target_path, parser_outputs_dir)
    with open(session_dir / "parser_summary.json", "w", encoding="utf-8") as f:
        json.dump(parser_results, f, ensure_ascii=False, indent=2)

    if getattr(args, "parse_only", False):
        return AuditResult(doc_type="kartochka_proekta", session_dir=session_dir, target_path=str(target_path), duration_sec=time.time() - start_time)

    rules = kartochka_proekta_run_validations_load_rules(rules_path)
    if args.rule_filter:
        rules = [r for r in rules if r["rule_index"] == str(args.rule_filter)]
        if not rules:
            raise ValueError(f"Правило {args.rule_filter} не найдено в kartochka_proekta")

    results = kartochka_proekta_run_validations_run_validators_parallel(
        rules=rules,
        parser_outputs_dir=parser_outputs_dir,
        output_dir=validation_outputs_dir,
        rules_path=rules_path,
        max_workers=max_workers,
    )
    df = kartochka_proekta_run_validations_create_excel_report(results, session_dir / "validation_report.xlsx")
    violations = []
    for rule, result_data, _ in results:
        if result_data.get("status") == "FAIL":
            violations.append({
                "rule_index": result_data.get("rule_index", rule.get("rule_index", "?")),
                "rule_title": result_data.get("rule_title", rule.get("rule_title", "")),
                "section": rule.get("section", ""),
                "discrepancy": result_data.get("discrepancy", ""),
            })
    return AuditResult(
        violations=violations,
        doc_type="kartochka_proekta",
        session_dir=session_dir,
        duration_sec=time.time() - start_time,
        rules_checked=len(df),
        target_path=str(target_path),
    )


def run_plan_grafik_special(args) -> AuditResult:
    start_time = time.time()
    target_path = str(args.target)
    session_dir = Path(args.session_dir) if args.session_dir else LOGS_RESULT_DIR / "plan_grafik" / f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    session_dir.mkdir(parents=True, exist_ok=True)
    try:
        parsed = parse_plan_grafik(target_path)
    except Exception as e:
        error_msg = f"Ошибка парсинга: {e}"
        (session_dir / "ERROR.txt").write_text(error_msg, encoding="utf-8")
        return AuditResult(
            violations=[],
            doc_type="plan_grafik",
            session_dir=session_dir,
            duration_sec=time.time() - start_time,
            rules_checked=0,
            target_path=target_path,
        )
    with open(session_dir / "parsed.json", "w", encoding="utf-8") as f:
        json.dump(parsed, f, ensure_ascii=False, indent=2, default=str)
    if getattr(args, "parse_only", False):
        return AuditResult(doc_type="plan_grafik", session_dir=session_dir, duration_sec=time.time() - start_time, target_path=target_path)
    violations = run_all_validators(parsed, target_path)
    save_to_excel(violations, str(session_dir / "validation_report.xlsx"))
    return AuditResult(
        violations=violations,
        doc_type="plan_grafik",
        session_dir=session_dir,
        duration_sec=time.time() - start_time,
        rules_checked=9,
        target_path=target_path,
    )


SPECIAL_ENGINE_RUNNERS = {
    "drivers": run_drivers_special,
    "kpsc": run_kpsc_special,
    "kartochka_proekta": run_kartochka_proekta_special,
    "plan_grafik": run_plan_grafik_special,
}


def _audit_engine_init(self, doc_type: str, base_dir: Optional[str] = None, config_dir: Optional[str] = None) -> None:
    self.doc_type = doc_type
    self.base_dir = Path(base_dir) if base_dir else PROJECT_ROOT
    self.config_path = Path(config_dir) if config_dir else self.base_dir / "doc_configs" / doc_type
    self.config = load_audit_config(self.config_path)
    # Парсер-конфиг — это singleton PARSERS_CONFIG из config/parsers.py.
    # Берём ссылки на нужные суб-конфиги, чтобы не тянуть в runtime сам объект.
    self.pdf_parser_config = PARSERS_CONFIG.pdf
    self.docx_header_ocr_config = PARSERS_CONFIG.docx_header_ocr
    self.logger = None


def _audit_engine_parse_secondary(self, file_path: str) -> Dict[str, Any]:
    spec = self.config.secondary_file
    parser_fn = get_parser(spec.parser)
    raw_chunks = parser_fn(file_path)
    prefixed = {f"{spec.chunk_prefix}{key}": value for key, value in raw_chunks.items()}
    prefixed[f"{spec.chunk_prefix}filename"] = Path(file_path).name
    self.logger.log_parsed_doc(prefixed, f"secondary_{spec.type}")
    return prefixed


def _audit_engine_list_doc_types(base_dir: Optional[str] = None) -> List[Dict[str, str]]:
    configs_dir = Path(base_dir) / "doc_configs" if base_dir else DOC_CONFIGS_DIR
    result = []
    if not configs_dir.exists():
        return result
    for config_dir in sorted(configs_dir.iterdir()):
        config_file = config_dir / "config.json"
        if not config_file.exists():
            continue
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            result.append({"doc_type": data.get("doc_type", config_dir.name), "doc_title": data.get("doc_title", "")})
        except (json.JSONDecodeError, KeyError):
            result.append({"doc_type": config_dir.name, "doc_title": "(ошибка чтения config.json)"})
    return result


AuditEngine.__init__ = _audit_engine_init
AuditEngine._parse_secondary = _audit_engine_parse_secondary
AuditEngine.list_doc_types = staticmethod(_audit_engine_list_doc_types)

UPLOAD_DIR = PROJECT_ROOT / "uploads"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


def api_server__detect_engine(doc_type: str) -> str:
    return read_doc_type_engine(doc_type)


api_server_SPECIAL_ENGINES = dict(SPECIAL_ENGINE_RUNNERS)


def _run_special_engine(engine_type: str, doc_type: str, target_path: str, session_dir: str) -> AuditResult:
    runner = api_server_SPECIAL_ENGINES[engine_type]
    args = SimpleNamespace(
        target=target_path,
        parse_only=False,
        rule_filter=None,
        model=None,
        temperature=None,
        session_dir=session_dir,
    )
    return runner(args)




def run_audit__detect_engine(doc_type: str) -> str:
    return read_doc_type_engine(doc_type)


run_audit_SPECIAL_ENGINES = dict(SPECIAL_ENGINE_RUNNERS)


def main() -> None:
    parser = argparse.ArgumentParser(description="Единый аудит документов — monolith backend")
    parser.add_argument("--doc-type", default=None, help="Тип документа (имя папки в doc_configs/)")
    parser.add_argument("--target", default=None, help="Путь к целевому документу")
    parser.add_argument("--model", default=None, help="Модель OpenAI (переопределяет конфиг)")
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

    engine_type = run_audit__detect_engine(args.doc_type)
    if engine_type in run_audit_SPECIAL_ENGINES:
        result = run_audit_SPECIAL_ENGINES[engine_type](args)
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




build_kpsc_header_payload = kpsc_parse_kpsc_header_build_payload
run_kpsc_parsers = kpsc_run_validations_run_parsers

# END_RUNTIME_INTEGRATION

if __name__ == "__main__":
    main()
