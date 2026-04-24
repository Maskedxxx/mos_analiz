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

# START_SOURCE_PLAN_GRAFIK_VALIDATORS
# PURPOSE: Inlined source from audit_engine/plan_grafik/validators.py.
def run_all_validators(parsed: Dict[str, Any], target_path: str) -> List[Dict[str, Any]]:
    """
    Запускает все 9 валидаторов последовательно.

    Args:
        parsed: результат парсинга из parser.py
        target_path: путь к файлу (для проверки имени)

    Returns:
        Список нарушений
    """
    violations = []
    validators = [validate_1_filename, validate_2_approval, validate_3_simple_headers, validate_4_complex_headers, validate_5_dates, validate_6_responsible, validate_7_calc_columns, validate_8_signature, validate_9_formulas]
    for validator in validators:
        try:
            result = validator(parsed, target_path)
            violations.extend(result)
        except Exception as e:
            violations.append({'rule_index': 0, 'rule_title': f'Ошибка валидатора {validator.__name__}', 'Целевой документ': str(e), 'Различие': 'Внутренняя ошибка валидатора'})
    return violations

def validate_1_filename(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 1: Проверка имени файла."""
    from pathlib import Path
    filename = Path(target_path).stem.lower()
    keywords = ['2.6', 'план', 'график']
    missing = []
    for kw in keywords:
        if kw == '2.6':
            if '2.6' not in filename and '2_6' not in filename:
                missing.append(kw)
        elif kw.lower() not in filename:
            missing.append(kw)
    if missing:
        return [{'rule_index': 1, 'rule_title': 'Проверка названия файла', 'Целевой документ': Path(target_path).name, 'Различие': f'Отсутствуют ключевые слова: {', '.join(missing)}'}]
    return []

def validate_2_approval(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 2: Проверка блока УТВЕРЖДАЮ."""
    approval = parsed.get('approval', {})
    violations = []
    marker = approval.get('marker', '')
    if 'утверждаю' not in marker.lower():
        violations.append({'rule_index': 2, 'rule_title': 'Проверка блока УТВЕРЖДАЮ', 'Целевой документ': f'DM8: {marker or 'пусто'}', 'Различие': 'Отсутствует слово «УТВЕРЖДАЮ»'})
    position = approval.get('position', '')
    if not position or (position == 'Генеральный директор' and len(position) < 5):
        pass
    if not position:
        violations.append({'rule_index': 2, 'rule_title': 'Проверка блока УТВЕРЖДАЮ', 'Целевой документ': f'DM9: пусто', 'Различие': 'Не указана должность подписанта'})
    company = approval.get('company', '')
    is_placeholder = not company or ('___' in company and (not re.search('[А-Яа-яA-Za-z]{3,}', company.replace('ООО', '').replace('АО', '').replace('ЗАО', ''))))
    if is_placeholder:
        violations.append({'rule_index': 2, 'rule_title': 'Проверка блока УТВЕРЖДАЮ', 'Целевой документ': f'DM10: {company or 'пусто'}', 'Различие': 'Наименование организации не заполнено (плейсхолдер)'})
    fio = approval.get('fio', '')
    is_fio_placeholder = not fio or 'фио' in fio.lower() or (fio.count('_') > 3 and (not re.search('[А-Яа-я]{2,}', fio.replace('ФИО', ''))))
    if is_fio_placeholder:
        violations.append({'rule_index': 2, 'rule_title': 'Проверка блока УТВЕРЖДАЮ', 'Целевой документ': f'DM11: {fio or 'пусто'}', 'Различие': 'ФИО подписанта не заполнено'})
    return violations

def validate_3_simple_headers(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 3: Проверка простых заголовков таблицы."""
    headers = parsed.get('headers_row16', {})
    if not headers:
        return [{'rule_index': 3, 'rule_title': 'Проверка заголовков таблицы', 'Целевой документ': 'Строка 16 пустая', 'Различие': 'Заголовки таблицы отсутствуют'}]
    all_text = ' '.join(headers.values()).lower()
    required = {'мероприятие': 'Мероприятие', 'ответственн': 'Ответственный', 'начало': 'Начало мероприятия', 'окончани': 'Окончание мероприятия', 'статус': 'Статус'}
    missing = []
    for keyword, name in required.items():
        if keyword not in all_text:
            missing.append(name)
    if missing:
        return [{'rule_index': 3, 'rule_title': 'Проверка заголовков таблицы', 'Целевой документ': ', '.join(headers.values())[:200], 'Различие': f'Отсутствуют столбцы: {', '.join(missing)}'}]
    return []

def validate_4_complex_headers(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 4: Проверка сложных двухуровневых заголовков."""
    h16 = parsed.get('headers_row16', {})
    h17 = parsed.get('headers_row17', {})
    all_text = ' '.join(list(h16.values()) + list(h17.values())).lower()
    required_sub = {'выработк': 'Влияние на показатель выработка', 'запас': 'Влияние на показатель запасы', 'впп': 'Влияние на показатель ВПП', 'проблем': '№ проблемы из КПСЦ', 'комментари': 'Комментарии'}
    missing = []
    for keyword, name in required_sub.items():
        if keyword not in all_text:
            missing.append(name)
    if missing:
        return [{'rule_index': 4, 'rule_title': 'Проверка структуры заголовков', 'Целевой документ': f'Строки 16-17: {len(h16)} + {len(h17)} столбцов', 'Различие': f'Отсутствуют подзаголовки: {', '.join(missing)}'}]
    return []

def validate_5_dates(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 5: Проверка дат (формат + логика)."""
    violations = []
    dates = parsed.get('dates', {})
    start_raw = dates.get('start_raw')
    end_raw = dates.get('end_raw')
    if not start_raw:
        violations.append({'rule_index': 5, 'rule_title': 'Проверка дат мероприятий', 'Целевой документ': 'I12: пусто', 'Различие': 'Дата начала мероприятий не заполнена'})
    if not end_raw:
        violations.append({'rule_index': 5, 'rule_title': 'Проверка дат мероприятий', 'Целевой документ': 'I13: пусто', 'Различие': 'Дата окончания мероприятий не заполнена'})
    if start_raw and end_raw:
        try:
            start_dt = start_raw if isinstance(start_raw, datetime) else datetime.strptime(str(start_raw)[:10], '%Y-%m-%d')
            end_dt = end_raw if isinstance(end_raw, datetime) else datetime.strptime(str(end_raw)[:10], '%Y-%m-%d')
            if end_dt <= start_dt:
                violations.append({'rule_index': 5, 'rule_title': 'Проверка дат мероприятий', 'Целевой документ': f'Начало: {start_dt.strftime('%d.%m.%Y')}, Окончание: {end_dt.strftime('%d.%m.%Y')}', 'Различие': 'Дата окончания не позже даты начала'})
        except (ValueError, TypeError):
            pass
    data_rows = parsed.get('data_rows', [])
    plan_rows_no_dates = []
    plan_rows_bad_logic = []
    for row_data in data_rows:
        if row_data.get('plan_fact', '').lower() != 'план':
            continue
        row_num = row_data.get('row', '?')
        start = row_data.get('start_date')
        end = row_data.get('end_date')
        if not start and (not end):
            plan_rows_no_dates.append(str(row_num))
        elif start and end:
            try:
                s = start if isinstance(start, datetime) else datetime.strptime(str(start)[:10], '%Y-%m-%d')
                e = end if isinstance(end, datetime) else datetime.strptime(str(end)[:10], '%Y-%m-%d')
                if e < s:
                    plan_rows_bad_logic.append(str(row_num))
            except (ValueError, TypeError):
                pass
    if plan_rows_no_dates:
        violations.append({'rule_index': 5, 'rule_title': 'Проверка дат мероприятий', 'Целевой документ': f'Строки без дат: {', '.join(plan_rows_no_dates[:10])}', 'Различие': 'Даты начала/окончания не заполнены для плановых мероприятий'})
    if plan_rows_bad_logic:
        violations.append({'rule_index': 5, 'rule_title': 'Проверка дат мероприятий', 'Целевой документ': f'Строки с нарушением логики: {', '.join(plan_rows_bad_logic[:10])}', 'Различие': 'Дата окончания раньше даты начала'})
    return violations

def validate_6_responsible(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 6: Проверка заполненности ответственных (строки «План»)."""
    data_rows = parsed.get('data_rows', [])
    empty_rows = []
    for row_data in data_rows:
        if row_data.get('plan_fact', '').lower() != 'план':
            continue
        responsible = row_data.get('responsible', '').strip()
        if not responsible:
            row_num = row_data.get('row', '?')
            problem = row_data.get('problem_num', '?')
            empty_rows.append(f'строка {row_num} (проблема №{problem})')
    if empty_rows:
        return [{'rule_index': 6, 'rule_title': 'Проверка заполненности ответственных', 'Целевой документ': f'Пустые: {', '.join(empty_rows[:10])}', 'Различие': 'Не указан ответственный за мероприятие (строки «План»)'}]
    return []

def validate_7_calc_columns(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 7: Проверка наличия расчётных столбцов (Статус, Отклонения, Комментарии)."""
    sc = parsed.get('status_columns', {})
    violations = []
    if not sc.get('status_present'):
        violations.append({'rule_index': 7, 'rule_title': 'Проверка расчётных столбцов', 'Целевой документ': 'Столбец DJ (Статус): отсутствует', 'Различие': 'Столбец «Статус» не найден в заголовках таблицы'})
    if not sc.get('comments_present'):
        violations.append({'rule_index': 7, 'rule_title': 'Проверка расчётных столбцов', 'Целевой документ': 'Столбец DM (Комментарии): отсутствует', 'Различие': 'Столбец «Комментарии» не найден в заголовках таблицы'})
    formulas = parsed.get('formulas', {})
    dk18 = formulas.get('DK18', {})
    dl18 = formulas.get('DL18', {})
    if not dk18.get('is_formula'):
        violations.append({'rule_index': 7, 'rule_title': 'Проверка расчётных столбцов', 'Целевой документ': f'DK18: {dk18.get('value', 'пусто')}', 'Различие': 'Столбец «Отклонение по началу» не содержит формулу'})
    if not dl18.get('is_formula'):
        violations.append({'rule_index': 7, 'rule_title': 'Проверка расчётных столбцов', 'Целевой документ': f'DL18: {dl18.get('value', 'пусто')}', 'Различие': 'Столбец «Отклонение по окончанию» не содержит формулу'})
    return violations

def validate_8_signature(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 8: Проверка блока подписи внизу документа."""
    approval = parsed.get('approval', {})
    date_line = approval.get('date_line', '')
    signature = parsed.get('signature', '').strip()
    if not date_line and (not signature):
        return [{'rule_index': 8, 'rule_title': 'Проверка блока подписи', 'Целевой документ': 'отсутствует', 'Различие': 'Строка подписи с датой не найдена в документе'}]
    return []

def validate_9_formulas(parsed: Dict, target_path: str) -> List[Dict]:
    """Rule 9: Проверка целостности формул (не заменены на значения, нет #REF)."""
    formulas = parsed.get('formulas', {})
    violations = []
    dk17 = formulas.get('DK17', {})
    if not dk17.get('is_formula'):
        violations.append({'rule_index': 9, 'rule_title': 'Проверка целостности формул', 'Целевой документ': f'DK17: {dk17.get('value', 'пусто')}', 'Различие': 'Формула длительности заменена на значение или отсутствует'})
    n12 = formulas.get('N12', {})
    if not n12.get('is_formula'):
        violations.append({'rule_index': 9, 'rule_title': 'Проверка целостности формул', 'Целевой документ': f'N12: {n12.get('value', 'пусто')}', 'Различие': 'Формулы Ганта заменены на значения или отсутствуют'})
    for cell_name, cell_data in formulas.items():
        if cell_data.get('has_error'):
            violations.append({'rule_index': 9, 'rule_title': 'Проверка целостности формул', 'Целевой документ': f'{cell_name}: {cell_data.get('value', '')}', 'Различие': 'Формула содержит ошибку #REF! (сломанная ссылка)'})
    return violations

plan_grafik_validators_module = SimpleNamespace(run_all_validators=run_all_validators, validate_1_filename=validate_1_filename, validate_2_approval=validate_2_approval, validate_3_simple_headers=validate_3_simple_headers, validate_4_complex_headers=validate_4_complex_headers, validate_5_dates=validate_5_dates, validate_6_responsible=validate_6_responsible, validate_7_calc_columns=validate_7_calc_columns, validate_8_signature=validate_8_signature, validate_9_formulas=validate_9_formulas)

# END_SOURCE_PLAN_GRAFIK_VALIDATORS

# START_SOURCE_KPSC_VALIDATE_1_1_KPSC_TEXT
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_1_1_kpsc_text.py.
kpsc_validate_1_1_kpsc_text_RULE_INDEX = '1.1'

kpsc_validate_1_1_kpsc_text_RULE_TITLE = "Наличие текста 'КПСЦ' в заголовке"

kpsc_validate_1_1_kpsc_text_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_1_1_kpsc_text_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_1_1_kpsc_text_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def kpsc_validate_1_1_kpsc_text_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'title': data['fields'].get('title', '')}

def kpsc_validate_1_1_kpsc_text_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nЗаголовок документа: "{extracted_data['title']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_1_1_kpsc_text_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_1_1_kpsc_text_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_1_1_kpsc_text_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_1_1_kpsc_text_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_1_1_kpsc_text_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_1_1_kpsc_text_RULE_INDEX}: {kpsc_validate_1_1_kpsc_text_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_1_1_kpsc_text_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_1_1_kpsc_text_RULE_INDEX}: {kpsc_validate_1_1_kpsc_text_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_1_1_kpsc_text_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_1_1_kpsc_text_RULE_INDEX}')
    rule = kpsc_validate_1_1_kpsc_text_load_rule(kpsc_validate_1_1_kpsc_text_RULE_INDEX)
    if args.verbose:
        kpsc_validate_1_1_kpsc_text_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_1_1_kpsc_text_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_1_1_kpsc_text_log_step(log_file, 2, 'Загрузка данных', f'Source file: {args.parser_outputs / 'kpsc_header_v2.json'}\nData loaded:\n{json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_1_1_kpsc_text_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_1_1_kpsc_text_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_1_1_kpsc_text_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_1_1_kpsc_text_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_1_1_kpsc_text_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_1_1_kpsc_text_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_1_1_kpsc_text_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_1_1_kpsc_text_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_1_1_kpsc_text_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_1_1_kpsc_text_module = SimpleNamespace(RULE_INDEX=kpsc_validate_1_1_kpsc_text_RULE_INDEX, RULE_TITLE=kpsc_validate_1_1_kpsc_text_RULE_TITLE, TARGET_DOC=kpsc_validate_1_1_kpsc_text_TARGET_DOC, load_rule=kpsc_validate_1_1_kpsc_text_load_rule, load_data=kpsc_validate_1_1_kpsc_text_load_data, extract_relevant_data=kpsc_validate_1_1_kpsc_text_extract_relevant_data, build_prompt=kpsc_validate_1_1_kpsc_text_build_prompt, call_llm=kpsc_validate_1_1_kpsc_text_call_llm, save_result=kpsc_validate_1_1_kpsc_text_save_result, log_step=kpsc_validate_1_1_kpsc_text_log_step, main=kpsc_validate_1_1_kpsc_text_main)

# END_SOURCE_KPSC_VALIDATE_1_1_KPSC_TEXT

# START_SOURCE_KPSC_VALIDATE_1_2_COMPANY_NAME
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_1_2_company_name.py.
kpsc_validate_1_2_company_name_RULE_INDEX = '1.2'

kpsc_validate_1_2_company_name_RULE_TITLE = 'Наличие названия предприятия в формате ООО "наименование"'

kpsc_validate_1_2_company_name_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_1_2_company_name_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_1_2_company_name_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def kpsc_validate_1_2_company_name_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'title': data['fields'].get('title', ''), 'organization': data['fields'].get('organization', ''), 'workbook': data.get('meta', {}).get('workbook', '')}

def kpsc_validate_1_2_company_name_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nЗаголовок документа: "{extracted_data['title']}"\nНазвание организации: "{extracted_data.get('organization', '')}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_1_2_company_name_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_1_2_company_name_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_1_2_company_name_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_1_2_company_name_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_1_2_company_name_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_1_2_company_name_RULE_INDEX}: {kpsc_validate_1_2_company_name_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_1_2_company_name_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_1_2_company_name_RULE_INDEX}: {kpsc_validate_1_2_company_name_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_1_2_company_name_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_1_2_company_name_RULE_INDEX}')
    rule = kpsc_validate_1_2_company_name_load_rule(kpsc_validate_1_2_company_name_RULE_INDEX)
    if args.verbose:
        kpsc_validate_1_2_company_name_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_1_2_company_name_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_1_2_company_name_log_step(log_file, 2, 'Загрузка данных', f'Source file: {args.parser_outputs / 'kpsc_header_v2.json'}\nData loaded:\n{json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_1_2_company_name_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_1_2_company_name_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    import re
    ORG_PATTERN = '(ООО|АО|ПАО|ОАО|ЗАО)\\s*[«"\\\']?.+?[»"\\\']?'
    org = extracted.get('organization', '') or ''
    title = extracted.get('title', '') or ''
    workbook = extracted.get('workbook', '') or ''
    workbook_name = Path(workbook).stem if workbook else ''
    org_found = bool(re.search(ORG_PATTERN, org))
    title_has_org = bool(re.search(ORG_PATTERN, title))
    filename_has_org = bool(re.search(ORG_PATTERN, workbook_name))
    if org_found or title_has_org or filename_has_org:
        print(f"[4/6] Детерминированная проверка: найдено '{org or extracted['title']}' → PASS")
        result = {'rule_index': rule['rule_index'], 'rule_title': rule['rule_title'], 'target_document': kpsc_validate_1_2_company_name_TARGET_DOC, 'status': 'PASS', 'discrepancy': ''}
        if args.verbose:
            kpsc_validate_1_2_company_name_log_step(log_file, 4, 'Детерминированная проверка', f"organization='{org}', title_has_org={title_has_org} → PASS без LLM")
    else:
        print('[4/6] Формирование промпта с правилом из ТЗ')
        prompt = kpsc_validate_1_2_company_name_build_prompt(extracted, rule)
        if args.verbose:
            kpsc_validate_1_2_company_name_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
        print('[5/6] Вызов LLM')
        result = kpsc_validate_1_2_company_name_call_llm(prompt, api_key)
        if args.verbose:
            kpsc_validate_1_2_company_name_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_1_2_company_name_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_1_2_company_name_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_1_2_company_name_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_1_2_company_name_module = SimpleNamespace(RULE_INDEX=kpsc_validate_1_2_company_name_RULE_INDEX, RULE_TITLE=kpsc_validate_1_2_company_name_RULE_TITLE, TARGET_DOC=kpsc_validate_1_2_company_name_TARGET_DOC, load_rule=kpsc_validate_1_2_company_name_load_rule, load_data=kpsc_validate_1_2_company_name_load_data, extract_relevant_data=kpsc_validate_1_2_company_name_extract_relevant_data, build_prompt=kpsc_validate_1_2_company_name_build_prompt, call_llm=kpsc_validate_1_2_company_name_call_llm, save_result=kpsc_validate_1_2_company_name_save_result, log_step=kpsc_validate_1_2_company_name_log_step, main=kpsc_validate_1_2_company_name_main)

# END_SOURCE_KPSC_VALIDATE_1_2_COMPANY_NAME

# START_SOURCE_KPSC_VALIDATE_1_3_FLOW_NAME
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_1_3_flow_name.py.
kpsc_validate_1_3_flow_name_RULE_INDEX = '1.3'

kpsc_validate_1_3_flow_name_RULE_TITLE = 'Наличие названия потока в формате "имя потока"'

kpsc_validate_1_3_flow_name_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_1_3_flow_name_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_1_3_flow_name_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def kpsc_validate_1_3_flow_name_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'flow_name': data['fields'].get('flow_name', '')}

def kpsc_validate_1_3_flow_name_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nНазвание потока (flow_name): "{extracted_data['flow_name']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_1_3_flow_name_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_1_3_flow_name_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_1_3_flow_name_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_1_3_flow_name_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_1_3_flow_name_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_1_3_flow_name_RULE_INDEX}: {kpsc_validate_1_3_flow_name_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_1_3_flow_name_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_1_3_flow_name_RULE_INDEX}: {kpsc_validate_1_3_flow_name_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_1_3_flow_name_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_1_3_flow_name_RULE_INDEX}')
    rule = kpsc_validate_1_3_flow_name_load_rule(kpsc_validate_1_3_flow_name_RULE_INDEX)
    if args.verbose:
        kpsc_validate_1_3_flow_name_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_1_3_flow_name_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_1_3_flow_name_log_step(log_file, 2, 'Загрузка данных', f'Source file: {args.parser_outputs / 'kpsc_header_v2.json'}\nData loaded:\n{json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_1_3_flow_name_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_1_3_flow_name_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_1_3_flow_name_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_1_3_flow_name_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_1_3_flow_name_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_1_3_flow_name_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_1_3_flow_name_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_1_3_flow_name_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_1_3_flow_name_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_1_3_flow_name_module = SimpleNamespace(RULE_INDEX=kpsc_validate_1_3_flow_name_RULE_INDEX, RULE_TITLE=kpsc_validate_1_3_flow_name_RULE_TITLE, TARGET_DOC=kpsc_validate_1_3_flow_name_TARGET_DOC, load_rule=kpsc_validate_1_3_flow_name_load_rule, load_data=kpsc_validate_1_3_flow_name_load_data, extract_relevant_data=kpsc_validate_1_3_flow_name_extract_relevant_data, build_prompt=kpsc_validate_1_3_flow_name_build_prompt, call_llm=kpsc_validate_1_3_flow_name_call_llm, save_result=kpsc_validate_1_3_flow_name_save_result, log_step=kpsc_validate_1_3_flow_name_log_step, main=kpsc_validate_1_3_flow_name_main)

# END_SOURCE_KPSC_VALIDATE_1_3_FLOW_NAME

# START_SOURCE_KPSC_VALIDATE_1_4_RESPONSIBLE
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_1_4_responsible.py.
kpsc_validate_1_4_responsible_RULE_INDEX = '1.4'

kpsc_validate_1_4_responsible_RULE_TITLE = 'Заполнение поля "Ответственный за поток" (ФИО)'

kpsc_validate_1_4_responsible_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_1_4_responsible_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_1_4_responsible_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def kpsc_validate_1_4_responsible_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    responsible = data['fields'].get('responsible', '') or ''
    words = responsible.strip().split()
    return {'responsible': responsible, 'word_count': len(words), 'char_length': len(responsible.strip())}

def kpsc_validate_1_4_responsible_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nОтветственный за поток (responsible): "{extracted_data['responsible']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_1_4_responsible_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_1_4_responsible_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_1_4_responsible_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_1_4_responsible_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_1_4_responsible_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_1_4_responsible_RULE_INDEX}: {kpsc_validate_1_4_responsible_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_1_4_responsible_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_1_4_responsible_RULE_INDEX}: {kpsc_validate_1_4_responsible_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_1_4_responsible_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_1_4_responsible_RULE_INDEX}')
    rule = kpsc_validate_1_4_responsible_load_rule(kpsc_validate_1_4_responsible_RULE_INDEX)
    if args.verbose:
        kpsc_validate_1_4_responsible_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_1_4_responsible_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_1_4_responsible_log_step(log_file, 2, 'Загрузка данных', f'Source file: {args.parser_outputs / 'kpsc_header_v2.json'}\nData loaded:\n{json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_1_4_responsible_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_1_4_responsible_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    if extracted['word_count'] >= 2 and extracted['char_length'] >= 5:
        print('[4/6] Детерминированная проверка: ФИО содержит ≥2 слов, ≥5 символов → PASS')
        result = {'rule_index': rule['rule_index'], 'rule_title': rule['rule_title'], 'target_document': kpsc_validate_1_4_responsible_TARGET_DOC, 'status': 'PASS', 'discrepancy': ''}
        if args.verbose:
            kpsc_validate_1_4_responsible_log_step(log_file, 4, 'Детерминированная проверка', f'word_count={extracted['word_count']}, char_length={extracted['char_length']} → PASS без LLM')
    else:
        print('[4/6] Формирование промпта с правилом из ТЗ')
        prompt = kpsc_validate_1_4_responsible_build_prompt(extracted, rule)
        if args.verbose:
            kpsc_validate_1_4_responsible_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
        print('[5/6] Вызов LLM')
        result = kpsc_validate_1_4_responsible_call_llm(prompt, api_key)
        if args.verbose:
            kpsc_validate_1_4_responsible_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_1_4_responsible_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_1_4_responsible_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_1_4_responsible_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_1_4_responsible_module = SimpleNamespace(RULE_INDEX=kpsc_validate_1_4_responsible_RULE_INDEX, RULE_TITLE=kpsc_validate_1_4_responsible_RULE_TITLE, TARGET_DOC=kpsc_validate_1_4_responsible_TARGET_DOC, load_rule=kpsc_validate_1_4_responsible_load_rule, load_data=kpsc_validate_1_4_responsible_load_data, extract_relevant_data=kpsc_validate_1_4_responsible_extract_relevant_data, build_prompt=kpsc_validate_1_4_responsible_build_prompt, call_llm=kpsc_validate_1_4_responsible_call_llm, save_result=kpsc_validate_1_4_responsible_save_result, log_step=kpsc_validate_1_4_responsible_log_step, main=kpsc_validate_1_4_responsible_main)

# END_SOURCE_KPSC_VALIDATE_1_4_RESPONSIBLE

# START_SOURCE_KPSC_VALIDATE_1_5_DATE_DEVELOPED
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_1_5_date_developed.py.
kpsc_validate_1_5_date_developed_RULE_INDEX = '1.5'

kpsc_validate_1_5_date_developed_RULE_TITLE = 'Заполнение даты разработки'

kpsc_validate_1_5_date_developed_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_1_5_date_developed_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_1_5_date_developed_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def kpsc_validate_1_5_date_developed_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'date_developed': data['fields'].get('date_developed', '')}

def kpsc_validate_1_5_date_developed_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nДата разработки (date_developed): "{extracted_data['date_developed']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_1_5_date_developed_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_1_5_date_developed_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_1_5_date_developed_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_1_5_date_developed_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_1_5_date_developed_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_1_5_date_developed_RULE_INDEX}: {kpsc_validate_1_5_date_developed_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_1_5_date_developed_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_1_5_date_developed_RULE_INDEX}: {kpsc_validate_1_5_date_developed_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_1_5_date_developed_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_1_5_date_developed_RULE_INDEX}')
    rule = kpsc_validate_1_5_date_developed_load_rule(kpsc_validate_1_5_date_developed_RULE_INDEX)
    if args.verbose:
        kpsc_validate_1_5_date_developed_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_1_5_date_developed_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_1_5_date_developed_log_step(log_file, 2, 'Загрузка данных', f'Source file: {args.parser_outputs / 'kpsc_header_v2.json'}\nData loaded:\n{json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_1_5_date_developed_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_1_5_date_developed_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_1_5_date_developed_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_1_5_date_developed_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_1_5_date_developed_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_1_5_date_developed_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_1_5_date_developed_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_1_5_date_developed_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_1_5_date_developed_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_1_5_date_developed_module = SimpleNamespace(RULE_INDEX=kpsc_validate_1_5_date_developed_RULE_INDEX, RULE_TITLE=kpsc_validate_1_5_date_developed_RULE_TITLE, TARGET_DOC=kpsc_validate_1_5_date_developed_TARGET_DOC, load_rule=kpsc_validate_1_5_date_developed_load_rule, load_data=kpsc_validate_1_5_date_developed_load_data, extract_relevant_data=kpsc_validate_1_5_date_developed_extract_relevant_data, build_prompt=kpsc_validate_1_5_date_developed_build_prompt, call_llm=kpsc_validate_1_5_date_developed_call_llm, save_result=kpsc_validate_1_5_date_developed_save_result, log_step=kpsc_validate_1_5_date_developed_log_step, main=kpsc_validate_1_5_date_developed_main)

# END_SOURCE_KPSC_VALIDATE_1_5_DATE_DEVELOPED

# START_SOURCE_KPSC_VALIDATE_1_6_DATE_IMPLEMENTATION
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_1_6_date_implementation.py.
kpsc_validate_1_6_date_implementation_RULE_INDEX = '1.6'

kpsc_validate_1_6_date_implementation_RULE_TITLE = 'Заполнение даты реализации'

kpsc_validate_1_6_date_implementation_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_1_6_date_implementation_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_1_6_date_implementation_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def kpsc_validate_1_6_date_implementation_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'date_implementation': data['fields'].get('date_implementation', '')}

def kpsc_validate_1_6_date_implementation_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nДата реализации (date_implementation): "{extracted_data['date_implementation']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_1_6_date_implementation_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_1_6_date_implementation_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_1_6_date_implementation_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_1_6_date_implementation_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_1_6_date_implementation_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_1_6_date_implementation_RULE_INDEX}: {kpsc_validate_1_6_date_implementation_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_1_6_date_implementation_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_1_6_date_implementation_RULE_INDEX}: {kpsc_validate_1_6_date_implementation_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_1_6_date_implementation_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_1_6_date_implementation_RULE_INDEX}')
    rule = kpsc_validate_1_6_date_implementation_load_rule(kpsc_validate_1_6_date_implementation_RULE_INDEX)
    if args.verbose:
        kpsc_validate_1_6_date_implementation_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_1_6_date_implementation_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_1_6_date_implementation_log_step(log_file, 2, 'Загрузка данных', f'Source file: {args.parser_outputs / 'kpsc_header_v2.json'}\nData loaded:\n{json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_1_6_date_implementation_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_1_6_date_implementation_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_1_6_date_implementation_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_1_6_date_implementation_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_1_6_date_implementation_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_1_6_date_implementation_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_1_6_date_implementation_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_1_6_date_implementation_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_1_6_date_implementation_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_1_6_date_implementation_module = SimpleNamespace(RULE_INDEX=kpsc_validate_1_6_date_implementation_RULE_INDEX, RULE_TITLE=kpsc_validate_1_6_date_implementation_RULE_TITLE, TARGET_DOC=kpsc_validate_1_6_date_implementation_TARGET_DOC, load_rule=kpsc_validate_1_6_date_implementation_load_rule, load_data=kpsc_validate_1_6_date_implementation_load_data, extract_relevant_data=kpsc_validate_1_6_date_implementation_extract_relevant_data, build_prompt=kpsc_validate_1_6_date_implementation_build_prompt, call_llm=kpsc_validate_1_6_date_implementation_call_llm, save_result=kpsc_validate_1_6_date_implementation_save_result, log_step=kpsc_validate_1_6_date_implementation_log_step, main=kpsc_validate_1_6_date_implementation_main)

# END_SOURCE_KPSC_VALIDATE_1_6_DATE_IMPLEMENTATION

# START_SOURCE_KPSC_VALIDATE_1_7_COMPILED_BY
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_1_7_compiled_by.py.
kpsc_validate_1_7_compiled_by_RULE_INDEX = '1.7'

kpsc_validate_1_7_compiled_by_RULE_TITLE = 'Заполнение поля "Кто составил" (ФИО)'

kpsc_validate_1_7_compiled_by_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_1_7_compiled_by_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_1_7_compiled_by_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    with open(header_file, 'r', encoding='utf-8') as f:
        return json.load(f)

def kpsc_validate_1_7_compiled_by_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    return {'compiled_by': data['fields'].get('compiled_by', '')}

def kpsc_validate_1_7_compiled_by_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nКто составил (compiled_by): "{extracted_data['compiled_by']}"\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_1_7_compiled_by_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_1_7_compiled_by_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_1_7_compiled_by_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_1_7_compiled_by_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_1_7_compiled_by_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_1_7_compiled_by_RULE_INDEX}: {kpsc_validate_1_7_compiled_by_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_1_7_compiled_by_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_1_7_compiled_by_RULE_INDEX}: {kpsc_validate_1_7_compiled_by_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_1_7_compiled_by_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_1_7_compiled_by_RULE_INDEX}')
    rule = kpsc_validate_1_7_compiled_by_load_rule(kpsc_validate_1_7_compiled_by_RULE_INDEX)
    if args.verbose:
        kpsc_validate_1_7_compiled_by_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_1_7_compiled_by_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_1_7_compiled_by_log_step(log_file, 2, 'Загрузка данных', f'Source file: {args.parser_outputs / 'kpsc_header_v2.json'}\nData loaded:\n{json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_1_7_compiled_by_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_1_7_compiled_by_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_1_7_compiled_by_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_1_7_compiled_by_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_1_7_compiled_by_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_1_7_compiled_by_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_1_7_compiled_by_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_1_7_compiled_by_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_1_7_compiled_by_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_1_7_compiled_by_module = SimpleNamespace(RULE_INDEX=kpsc_validate_1_7_compiled_by_RULE_INDEX, RULE_TITLE=kpsc_validate_1_7_compiled_by_RULE_TITLE, TARGET_DOC=kpsc_validate_1_7_compiled_by_TARGET_DOC, load_rule=kpsc_validate_1_7_compiled_by_load_rule, load_data=kpsc_validate_1_7_compiled_by_load_data, extract_relevant_data=kpsc_validate_1_7_compiled_by_extract_relevant_data, build_prompt=kpsc_validate_1_7_compiled_by_build_prompt, call_llm=kpsc_validate_1_7_compiled_by_call_llm, save_result=kpsc_validate_1_7_compiled_by_save_result, log_step=kpsc_validate_1_7_compiled_by_log_step, main=kpsc_validate_1_7_compiled_by_main)

# END_SOURCE_KPSC_VALIDATE_1_7_COMPILED_BY

# START_SOURCE_KPSC_VALIDATE_2_1_PROBLEMS_COUNT
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_2_1_problems_count.py.
kpsc_validate_2_1_problems_count_RULE_INDEX = '2.1'

kpsc_validate_2_1_problems_count_RULE_TITLE = 'Количество проблем в таблице "Оцифровка потерь"'

kpsc_validate_2_1_problems_count_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_2_1_problems_count_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_2_1_problems_count_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    ocifrovka_file = parser_outputs_dir / 'ocifrovka_poteri_v2.json'
    if not ocifrovka_file.exists():
        return {'file_exists': False, 'data': None}
    with open(ocifrovka_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

def kpsc_validate_2_1_problems_count_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'problems_count': 0, 'problem_numbers': []}
    rows = data['data'].get('rows', [])
    problem_numbers = []
    for i, row in enumerate(rows):
        if i == 0:
            continue
        cells = row.get('cells', [])
        for cell in cells:
            if cell.get('col') in [1, 2]:
                problem_num_value = cell.get('value')
                if problem_num_value and str(problem_num_value).strip():
                    try:
                        problem_num = int(str(problem_num_value).strip())
                        problem_numbers.append(problem_num)
                    except:
                        pass
                    break
    return {'file_exists': True, 'problems_count': len(problem_numbers), 'problem_numbers': problem_numbers, 'min_number': min(problem_numbers) if problem_numbers else 0, 'max_number': max(problem_numbers) if problem_numbers else 0}

def kpsc_validate_2_1_problems_count_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nКоличество проблем: {extracted_data['problems_count']}\nНомера проблем: {extracted_data['problem_numbers']}\nДиапазон: с {extracted_data.get('min_number', 0)} по {extracted_data.get('max_number', 0)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\nЭто информационная проверка - нужно подтвердить, что найдены проблемы и извлечены номера.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_2_1_problems_count_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_2_1_problems_count_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_2_1_problems_count_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_2_1_problems_count_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_2_1_problems_count_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_2_1_problems_count_RULE_INDEX}: {kpsc_validate_2_1_problems_count_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_2_1_problems_count_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_2_1_problems_count_RULE_INDEX}: {kpsc_validate_2_1_problems_count_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_2_1_problems_count_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_2_1_problems_count_RULE_INDEX}')
    rule = kpsc_validate_2_1_problems_count_load_rule(kpsc_validate_2_1_problems_count_RULE_INDEX)
    if args.verbose:
        kpsc_validate_2_1_problems_count_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_2_1_problems_count_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_2_1_problems_count_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_2_1_problems_count_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_2_1_problems_count_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_2_1_problems_count_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_2_1_problems_count_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_2_1_problems_count_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_2_1_problems_count_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_2_1_problems_count_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_2_1_problems_count_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_2_1_problems_count_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_2_1_problems_count_module = SimpleNamespace(RULE_INDEX=kpsc_validate_2_1_problems_count_RULE_INDEX, RULE_TITLE=kpsc_validate_2_1_problems_count_RULE_TITLE, TARGET_DOC=kpsc_validate_2_1_problems_count_TARGET_DOC, load_rule=kpsc_validate_2_1_problems_count_load_rule, load_data=kpsc_validate_2_1_problems_count_load_data, extract_relevant_data=kpsc_validate_2_1_problems_count_extract_relevant_data, build_prompt=kpsc_validate_2_1_problems_count_build_prompt, call_llm=kpsc_validate_2_1_problems_count_call_llm, save_result=kpsc_validate_2_1_problems_count_save_result, log_step=kpsc_validate_2_1_problems_count_log_step, main=kpsc_validate_2_1_problems_count_main)

# END_SOURCE_KPSC_VALIDATE_2_1_PROBLEMS_COUNT

# START_SOURCE_KPSC_VALIDATE_2_2_PROBLEMS_SEQUENCE
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_2_2_problems_sequence.py.
kpsc_validate_2_2_problems_sequence_RULE_INDEX = '2.2'

kpsc_validate_2_2_problems_sequence_RULE_TITLE = 'Последовательность номеров проблем (пропуски в нумерации)'

kpsc_validate_2_2_problems_sequence_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_2_2_problems_sequence_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_2_2_problems_sequence_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    ocifrovka_file = parser_outputs_dir / 'ocifrovka_poteri_v2.json'
    if not ocifrovka_file.exists():
        return {'file_exists': False, 'data': None}
    with open(ocifrovka_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

def kpsc_validate_2_2_problems_sequence_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'problem_numbers': [], 'missing_numbers': []}
    rows = data['data'].get('rows', [])
    problem_numbers = []
    for i, row in enumerate(rows):
        if i == 0:
            continue
        cells = row.get('cells', [])
        for cell in cells:
            if cell.get('col') == 2:
                problem_num_value = cell.get('value')
                if problem_num_value and str(problem_num_value).strip():
                    try:
                        problem_num = int(str(problem_num_value).strip())
                        problem_numbers.append(problem_num)
                    except:
                        pass
                break
    if problem_numbers:
        min_num = min(problem_numbers)
        max_num = max(problem_numbers)
        expected_numbers = set(range(min_num, max_num + 1))
        actual_numbers = set(problem_numbers)
        missing_numbers = sorted(expected_numbers - actual_numbers)
    else:
        missing_numbers = []
    return {'file_exists': True, 'problem_numbers': sorted(problem_numbers), 'missing_numbers': missing_numbers, 'has_missing': len(missing_numbers) > 0}

def kpsc_validate_2_2_problems_sequence_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nНайденные номера проблем: {extracted_data['problem_numbers']}\nПропущенные номера: {extracted_data['missing_numbers']}\nЕсть пропуски: {extracted_data['has_missing']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_2_2_problems_sequence_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_2_2_problems_sequence_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_2_2_problems_sequence_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_2_2_problems_sequence_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_2_2_problems_sequence_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_2_2_problems_sequence_RULE_INDEX}: {kpsc_validate_2_2_problems_sequence_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_2_2_problems_sequence_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_2_2_problems_sequence_RULE_INDEX}: {kpsc_validate_2_2_problems_sequence_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_2_2_problems_sequence_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_2_2_problems_sequence_RULE_INDEX}')
    rule = kpsc_validate_2_2_problems_sequence_load_rule(kpsc_validate_2_2_problems_sequence_RULE_INDEX)
    if args.verbose:
        kpsc_validate_2_2_problems_sequence_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_2_2_problems_sequence_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_2_2_problems_sequence_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_2_2_problems_sequence_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_2_2_problems_sequence_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_2_2_problems_sequence_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_2_2_problems_sequence_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_2_2_problems_sequence_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_2_2_problems_sequence_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_2_2_problems_sequence_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_2_2_problems_sequence_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_2_2_problems_sequence_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_2_2_problems_sequence_module = SimpleNamespace(RULE_INDEX=kpsc_validate_2_2_problems_sequence_RULE_INDEX, RULE_TITLE=kpsc_validate_2_2_problems_sequence_RULE_TITLE, TARGET_DOC=kpsc_validate_2_2_problems_sequence_TARGET_DOC, load_rule=kpsc_validate_2_2_problems_sequence_load_rule, load_data=kpsc_validate_2_2_problems_sequence_load_data, extract_relevant_data=kpsc_validate_2_2_problems_sequence_extract_relevant_data, build_prompt=kpsc_validate_2_2_problems_sequence_build_prompt, call_llm=kpsc_validate_2_2_problems_sequence_call_llm, save_result=kpsc_validate_2_2_problems_sequence_save_result, log_step=kpsc_validate_2_2_problems_sequence_log_step, main=kpsc_validate_2_2_problems_sequence_main)

# END_SOURCE_KPSC_VALIDATE_2_2_PROBLEMS_SEQUENCE

# START_SOURCE_KPSC_VALIDATE_2_3_PROBLEMS_DESCRIPTION
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_2_3_problems_description.py.
kpsc_validate_2_3_problems_description_RULE_INDEX = '2.3'

kpsc_validate_2_3_problems_description_RULE_TITLE = 'Наличие описания для каждого номера проблемы'

kpsc_validate_2_3_problems_description_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_2_3_problems_description_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_2_3_problems_description_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    ocifrovka_file = parser_outputs_dir / 'ocifrovka_poteri_v2.json'
    if not ocifrovka_file.exists():
        return {'file_exists': False, 'data': None}
    with open(ocifrovka_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

def kpsc_validate_2_3_problems_description_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'problems_without_description': []}
    rows = data['data'].get('rows', [])
    problems_without_description = []
    for i, row in enumerate(rows):
        if i == 0:
            continue
        cells = row.get('cells', [])
        problem_num_value = None
        description = None
        for cell in cells:
            if cell.get('col') == 2:
                problem_num_value = cell.get('value', '')
            if cell.get('col') == 4:
                description = cell.get('value', '')
        if problem_num_value and str(problem_num_value).strip():
            if not description or not str(description).strip():
                try:
                    problem_num = int(str(problem_num_value).strip())
                    problems_without_description.append(problem_num)
                except:
                    pass
    return {'file_exists': True, 'problems_without_description': problems_without_description, 'has_missing_descriptions': len(problems_without_description) > 0, 'missing_count': len(problems_without_description)}

def kpsc_validate_2_3_problems_description_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nПроблемы без описания: {extracted_data['problems_without_description']}\nЕсть проблемы без описания: {extracted_data['has_missing_descriptions']}\nКоличество проблем без описания: {extracted_data['missing_count']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_2_3_problems_description_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_2_3_problems_description_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_2_3_problems_description_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_2_3_problems_description_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_2_3_problems_description_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_2_3_problems_description_RULE_INDEX}: {kpsc_validate_2_3_problems_description_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_2_3_problems_description_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_2_3_problems_description_RULE_INDEX}: {kpsc_validate_2_3_problems_description_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_2_3_problems_description_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_2_3_problems_description_RULE_INDEX}')
    rule = kpsc_validate_2_3_problems_description_load_rule(kpsc_validate_2_3_problems_description_RULE_INDEX)
    if args.verbose:
        kpsc_validate_2_3_problems_description_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_2_3_problems_description_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_2_3_problems_description_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_2_3_problems_description_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_2_3_problems_description_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_2_3_problems_description_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_2_3_problems_description_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_2_3_problems_description_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_2_3_problems_description_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_2_3_problems_description_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_2_3_problems_description_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_2_3_problems_description_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_2_3_problems_description_module = SimpleNamespace(RULE_INDEX=kpsc_validate_2_3_problems_description_RULE_INDEX, RULE_TITLE=kpsc_validate_2_3_problems_description_RULE_TITLE, TARGET_DOC=kpsc_validate_2_3_problems_description_TARGET_DOC, load_rule=kpsc_validate_2_3_problems_description_load_rule, load_data=kpsc_validate_2_3_problems_description_load_data, extract_relevant_data=kpsc_validate_2_3_problems_description_extract_relevant_data, build_prompt=kpsc_validate_2_3_problems_description_build_prompt, call_llm=kpsc_validate_2_3_problems_description_call_llm, save_result=kpsc_validate_2_3_problems_description_save_result, log_step=kpsc_validate_2_3_problems_description_log_step, main=kpsc_validate_2_3_problems_description_main)

# END_SOURCE_KPSC_VALIDATE_2_3_PROBLEMS_DESCRIPTION

# START_SOURCE_KPSC_VALIDATE_4_1_VPP_FILLED
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_4_1_vpp_filled.py.
kpsc_validate_4_1_vpp_filled_RULE_INDEX = '4.1'

kpsc_validate_4_1_vpp_filled_RULE_TITLE = 'Все ячейки строки ВПП заполнены'

kpsc_validate_4_1_vpp_filled_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_4_1_vpp_filled_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_4_1_vpp_filled_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    table1_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not table1_file.exists():
        return {'file_exists': False, 'data': None}
    with open(table1_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

def kpsc_validate_4_1_vpp_filled_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'empty_cells': []}
    rows = data['data'].get('rows', [])
    vpp_row_idx = None
    header_row_idx = None
    for i, row in enumerate(rows):
        cells = row.get('cells', [])
        for cell in cells:
            if cell.get('col') in [2, 3]:
                cell_value = str(cell.get('value', '')).upper()
                if 'ВПП' in cell_value:
                    vpp_row_idx = i
                    for j in range(max(0, i - 10), i):
                        row_cells = rows[j].get('cells', [])
                        for rc in row_cells:
                            if rc.get('col') in [2, 3]:
                                if str(rc.get('value', '')).lower().strip() == 'показатель':
                                    header_row_idx = j
                                    break
                        if header_row_idx is not None:
                            break
                    break
        if vpp_row_idx is not None:
            break
    if vpp_row_idx is None:
        return {'file_exists': True, 'empty_cells': [], 'error': 'Не найдена строка ВПП'}
    vpp_cells = rows[vpp_row_idx].get('cells', [])
    vpp_dict = {cell.get('col'): cell.get('value') for cell in vpp_cells}
    header_cells = rows[header_row_idx].get('cells', []) if header_row_idx is not None else []
    header_dict = {cell.get('col'): cell.get('value', '') for cell in header_cells}
    empty_cells = []
    all_cols = set(vpp_dict.keys()) | set(header_dict.keys())
    for col_num in sorted(all_cols):
        if col_num < 5:
            continue
        raw_header = header_dict.get(col_num)
        header_val = str(raw_header).strip() if raw_header is not None else ''
        if not header_val:
            continue
        if 'итого' in header_val.lower():
            continue
        vpp_value = vpp_dict.get(col_num)
        if vpp_value is None or str(vpp_value).strip() == '':
            empty_cells.append({'column': col_num, 'operation_name': header_val if header_val else f'Колонка {col_num}'})
    return {'file_exists': True, 'empty_cells': empty_cells, 'has_empty': len(empty_cells) > 0, 'empty_count': len(empty_cells)}

def kpsc_validate_4_1_vpp_filled_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nПустые ячейки в строке ВПП:\n{json.dumps(extracted_data.get('empty_cells', []), ensure_ascii=False, indent=2)}\n\nЕсть пустые ячейки: {extracted_data.get('has_empty', False)}\nКоличество пустых: {extracted_data.get('empty_count', 0)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_4_1_vpp_filled_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_4_1_vpp_filled_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_4_1_vpp_filled_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_4_1_vpp_filled_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_4_1_vpp_filled_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_4_1_vpp_filled_RULE_INDEX}: {kpsc_validate_4_1_vpp_filled_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_4_1_vpp_filled_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_4_1_vpp_filled_RULE_INDEX}: {kpsc_validate_4_1_vpp_filled_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_4_1_vpp_filled_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_4_1_vpp_filled_RULE_INDEX}')
    rule = kpsc_validate_4_1_vpp_filled_load_rule(kpsc_validate_4_1_vpp_filled_RULE_INDEX)
    if args.verbose:
        kpsc_validate_4_1_vpp_filled_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_4_1_vpp_filled_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_4_1_vpp_filled_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_4_1_vpp_filled_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_4_1_vpp_filled_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_4_1_vpp_filled_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_4_1_vpp_filled_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_4_1_vpp_filled_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_4_1_vpp_filled_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_4_1_vpp_filled_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_4_1_vpp_filled_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_4_1_vpp_filled_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_4_1_vpp_filled_module = SimpleNamespace(RULE_INDEX=kpsc_validate_4_1_vpp_filled_RULE_INDEX, RULE_TITLE=kpsc_validate_4_1_vpp_filled_RULE_TITLE, TARGET_DOC=kpsc_validate_4_1_vpp_filled_TARGET_DOC, load_rule=kpsc_validate_4_1_vpp_filled_load_rule, load_data=kpsc_validate_4_1_vpp_filled_load_data, extract_relevant_data=kpsc_validate_4_1_vpp_filled_extract_relevant_data, build_prompt=kpsc_validate_4_1_vpp_filled_build_prompt, call_llm=kpsc_validate_4_1_vpp_filled_call_llm, save_result=kpsc_validate_4_1_vpp_filled_save_result, log_step=kpsc_validate_4_1_vpp_filled_log_step, main=kpsc_validate_4_1_vpp_filled_main)

# END_SOURCE_KPSC_VALIDATE_4_1_VPP_FILLED

# START_SOURCE_KPSC_VALIDATE_4_2_VPP_SUM
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_4_2_vpp_sum.py.
kpsc_validate_4_2_vpp_sum_RULE_INDEX = '4.2'

kpsc_validate_4_2_vpp_sum_RULE_TITLE = 'Сумма значений строки ВПП совпадает с итоговым значением'

kpsc_validate_4_2_vpp_sum_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_4_2_vpp_sum_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_4_2_vpp_sum_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    table1_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not table1_file.exists():
        return {'file_exists': False, 'data': None}
    with open(table1_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

def kpsc_validate_4_2_vpp_sum_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'vpp_values': [], 'calculated_sum': 0, 'itogo_value': 0}
    rows = data['data'].get('rows', [])
    vpp_row_idx = None
    header_row_idx = None
    for i, row in enumerate(rows):
        cells = row.get('cells', [])
        for cell in cells:
            if cell.get('col') in [2, 3]:
                cell_value = str(cell.get('value', '')).upper()
                if 'ВПП' in cell_value:
                    vpp_row_idx = i
                    for j in range(max(0, i - 20), i):
                        row_cells = rows[j].get('cells', [])
                        for rc in row_cells:
                            if rc.get('col') in [2, 3]:
                                if str(rc.get('value', '')).lower().strip() == 'показатель':
                                    header_row_idx = j
                                    break
                        if header_row_idx is not None:
                            break
                    break
        if vpp_row_idx is not None:
            break
    if vpp_row_idx is None:
        return {'file_exists': True, 'error': 'Не найдена строка ВПП'}
    vpp_cells = rows[vpp_row_idx].get('cells', [])
    vpp_dict = {}
    vpp_cells_full = {}
    for cell in vpp_cells:
        col = cell.get('col')
        vpp_dict[col] = cell.get('value')
        vpp_cells_full[col] = cell
    header_cells = rows[header_row_idx].get('cells', []) if header_row_idx is not None else []
    header_dict = {cell.get('col'): str(cell.get('value', '')).strip() for cell in header_cells}
    vpp_values = []
    itogo_value = None
    all_cols = set(vpp_dict.keys()) | set(header_dict.keys())
    for col_num in sorted(all_cols):
        if col_num < 5:
            continue
        header_val = header_dict.get(col_num, '')
        if col_num in vpp_cells_full:
            cell_info = vpp_cells_full[col_num]
            is_merge_anchor = cell_info.get('merge_anchor', False)
            merge_range = cell_info.get('merge_range')
            if merge_range is not None and (not is_merge_anchor):
                continue
        cell_value = vpp_dict.get(col_num)
        if 'итого' in header_val.lower():
            try:
                itogo_value = float(str(cell_value).replace(',', '.'))
            except:
                itogo_value = None
            break
        if cell_value is not None and str(cell_value).strip():
            try:
                num_value = float(str(cell_value).replace(',', '.'))
                vpp_values.append({'column': col_num, 'operation': header_val if header_val else f'Колонка {col_num}', 'value': num_value})
            except:
                pass
    calculated_sum = sum((v['value'] for v in vpp_values))
    difference = None
    if itogo_value is not None:
        difference = abs(calculated_sum - itogo_value)
    return {'file_exists': True, 'vpp_values': vpp_values, 'calculated_sum': calculated_sum, 'itogo_value': itogo_value, 'difference': difference, 'matches': difference <= 0.01 if difference is not None else False}

def kpsc_validate_4_2_vpp_sum_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nЗначения ВПП по операциям:\n{json.dumps(extracted_data.get('vpp_values', []), ensure_ascii=False, indent=2)}\n\nВычисленная сумма: {extracted_data.get('calculated_sum', 0)}\nЗначение ИТОГО: {extracted_data.get('itogo_value', 0)}\nРасхождение: {extracted_data.get('difference', 0)}\nСуммы совпадают (±0.01): {extracted_data.get('matches', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_4_2_vpp_sum_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_4_2_vpp_sum_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_4_2_vpp_sum_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_4_2_vpp_sum_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_4_2_vpp_sum_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_4_2_vpp_sum_RULE_INDEX}: {kpsc_validate_4_2_vpp_sum_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_4_2_vpp_sum_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_4_2_vpp_sum_RULE_INDEX}: {kpsc_validate_4_2_vpp_sum_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_4_2_vpp_sum_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_4_2_vpp_sum_RULE_INDEX}')
    rule = kpsc_validate_4_2_vpp_sum_load_rule(kpsc_validate_4_2_vpp_sum_RULE_INDEX)
    if args.verbose:
        kpsc_validate_4_2_vpp_sum_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_4_2_vpp_sum_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_4_2_vpp_sum_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_4_2_vpp_sum_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_4_2_vpp_sum_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_4_2_vpp_sum_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_4_2_vpp_sum_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_4_2_vpp_sum_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_4_2_vpp_sum_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_4_2_vpp_sum_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_4_2_vpp_sum_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_4_2_vpp_sum_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_4_2_vpp_sum_module = SimpleNamespace(RULE_INDEX=kpsc_validate_4_2_vpp_sum_RULE_INDEX, RULE_TITLE=kpsc_validate_4_2_vpp_sum_RULE_TITLE, TARGET_DOC=kpsc_validate_4_2_vpp_sum_TARGET_DOC, load_rule=kpsc_validate_4_2_vpp_sum_load_rule, load_data=kpsc_validate_4_2_vpp_sum_load_data, extract_relevant_data=kpsc_validate_4_2_vpp_sum_extract_relevant_data, build_prompt=kpsc_validate_4_2_vpp_sum_build_prompt, call_llm=kpsc_validate_4_2_vpp_sum_call_llm, save_result=kpsc_validate_4_2_vpp_sum_save_result, log_step=kpsc_validate_4_2_vpp_sum_log_step, main=kpsc_validate_4_2_vpp_sum_main)

# END_SOURCE_KPSC_VALIDATE_4_2_VPP_SUM

# START_SOURCE_KPSC_VALIDATE_5_1_UNITS
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_5_1_units.py.
kpsc_validate_5_1_units_RULE_INDEX = '5.1'

kpsc_validate_5_1_units_RULE_TITLE = 'Заполнение всех ячеек в столбце "Единицы измерения"'

kpsc_validate_5_1_units_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_5_1_units_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_5_1_units_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    table1_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not table1_file.exists():
        return {'file_exists': False, 'data': None}
    with open(table1_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

def kpsc_validate_5_1_units_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'empty_units': []}
    rows = data['data'].get('rows', [])
    empty_units = []
    units_col = 4
    for i, row in enumerate(rows):
        if i == 0:
            continue
        cells = row.get('cells', [])
        row_num = row.get('row')
        cells_dict = {cell.get('col'): cell.get('value') for cell in cells}
        unit_value = cells_dict.get(units_col)
        if unit_value is None or str(unit_value).strip() == '':
            indicator_name = cells_dict.get(2, '')
            if not indicator_name or str(indicator_name).strip() == '':
                indicator_name = cells_dict.get(3, '')
            has_numeric_data = False
            for col_idx, val in cells_dict.items():
                if col_idx >= 5 and val is not None:
                    try:
                        float(str(val).replace(',', '.').strip())
                        has_numeric_data = True
                        break
                    except (ValueError, TypeError):
                        pass
            if not has_numeric_data:
                continue
            empty_units.append({'row': row_num, 'indicator': str(indicator_name).strip()})
    return {'file_exists': True, 'empty_units': empty_units, 'has_empty_units': len(empty_units) > 0, 'empty_count': len(empty_units)}

def kpsc_validate_5_1_units_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nСтроки с пустыми единицами измерения: {extracted_data['empty_units']}\nЕсть пустые ячейки: {extracted_data['has_empty_units']}\nКоличество пустых: {extracted_data['empty_count']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_5_1_units_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_5_1_units_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_5_1_units_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_5_1_units_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_5_1_units_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_5_1_units_RULE_INDEX}: {kpsc_validate_5_1_units_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_5_1_units_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_5_1_units_RULE_INDEX}: {kpsc_validate_5_1_units_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_5_1_units_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_5_1_units_RULE_INDEX}')
    rule = kpsc_validate_5_1_units_load_rule(kpsc_validate_5_1_units_RULE_INDEX)
    if args.verbose:
        kpsc_validate_5_1_units_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_5_1_units_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_5_1_units_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_5_1_units_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_5_1_units_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_5_1_units_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_5_1_units_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_5_1_units_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_5_1_units_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_5_1_units_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_5_1_units_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_5_1_units_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_5_1_units_module = SimpleNamespace(RULE_INDEX=kpsc_validate_5_1_units_RULE_INDEX, RULE_TITLE=kpsc_validate_5_1_units_RULE_TITLE, TARGET_DOC=kpsc_validate_5_1_units_TARGET_DOC, load_rule=kpsc_validate_5_1_units_load_rule, load_data=kpsc_validate_5_1_units_load_data, extract_relevant_data=kpsc_validate_5_1_units_extract_relevant_data, build_prompt=kpsc_validate_5_1_units_build_prompt, call_llm=kpsc_validate_5_1_units_call_llm, save_result=kpsc_validate_5_1_units_save_result, log_step=kpsc_validate_5_1_units_log_step, main=kpsc_validate_5_1_units_main)

# END_SOURCE_KPSC_VALIDATE_5_1_UNITS

# START_SOURCE_KPSC_VALIDATE_6_1_TRANSPORT_ROW
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_6_1_transport_row.py.
kpsc_validate_6_1_transport_row_RULE_INDEX = '6.1'

kpsc_validate_6_1_transport_row_RULE_TITLE = 'Наличие строки "Перемещения" в блоке "Расчет ВПП"'

kpsc_validate_6_1_transport_row_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_6_1_transport_row_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_6_1_transport_row_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    table1_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not table1_file.exists():
        return {'file_exists': False, 'data': None}
    with open(table1_file, 'r', encoding='utf-8') as f:
        return {'file_exists': True, 'data': json.load(f)}

def kpsc_validate_6_1_transport_row_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_transport_row': False}
    rows = data['data'].get('rows', [])
    has_transport_row = False
    transport_row_num = None
    for row in rows:
        cells = row.get('cells', [])
        row_num = row.get('row')
        for cell in cells:
            if cell.get('col') in [1, 2, 3]:
                cell_value = str(cell.get('value', '')).lower()
                if 'перемещени' in cell_value:
                    has_transport_row = True
                    transport_row_num = row_num
                    break
        if has_transport_row:
            break
    return {'file_exists': True, 'has_transport_row': has_transport_row, 'transport_row_index': transport_row_num}

def kpsc_validate_6_1_transport_row_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nСтрока "Перемещения" найдена: {extracted_data['has_transport_row']}\nИндекс строки: {extracted_data.get('transport_row_index', 'не найдена')}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_6_1_transport_row_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_6_1_transport_row_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_6_1_transport_row_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_6_1_transport_row_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_6_1_transport_row_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_6_1_transport_row_RULE_INDEX}: {kpsc_validate_6_1_transport_row_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_6_1_transport_row_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_6_1_transport_row_RULE_INDEX}: {kpsc_validate_6_1_transport_row_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_6_1_transport_row_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_6_1_transport_row_RULE_INDEX}')
    rule = kpsc_validate_6_1_transport_row_load_rule(kpsc_validate_6_1_transport_row_RULE_INDEX)
    if args.verbose:
        kpsc_validate_6_1_transport_row_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_6_1_transport_row_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_6_1_transport_row_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_6_1_transport_row_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_6_1_transport_row_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_6_1_transport_row_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_6_1_transport_row_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_6_1_transport_row_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_6_1_transport_row_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_6_1_transport_row_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_6_1_transport_row_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_6_1_transport_row_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_6_1_transport_row_module = SimpleNamespace(RULE_INDEX=kpsc_validate_6_1_transport_row_RULE_INDEX, RULE_TITLE=kpsc_validate_6_1_transport_row_RULE_TITLE, TARGET_DOC=kpsc_validate_6_1_transport_row_TARGET_DOC, load_rule=kpsc_validate_6_1_transport_row_load_rule, load_data=kpsc_validate_6_1_transport_row_load_data, extract_relevant_data=kpsc_validate_6_1_transport_row_extract_relevant_data, build_prompt=kpsc_validate_6_1_transport_row_build_prompt, call_llm=kpsc_validate_6_1_transport_row_call_llm, save_result=kpsc_validate_6_1_transport_row_save_result, log_step=kpsc_validate_6_1_transport_row_log_step, main=kpsc_validate_6_1_transport_row_main)

# END_SOURCE_KPSC_VALIDATE_6_1_TRANSPORT_ROW

# START_SOURCE_KPSC_VALIDATE_7_1_SHEET_KPSC
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_7_1_sheet_kpsc.py.
kpsc_validate_7_1_sheet_kpsc_RULE_INDEX = '7.1'

kpsc_validate_7_1_sheet_kpsc_RULE_TITLE = 'Наличие и заполнение листа "КПСЦ"'

kpsc_validate_7_1_sheet_kpsc_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_7_1_sheet_kpsc_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_7_1_sheet_kpsc_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    table1_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    files_exist = {'kpsc_header_v2.json': header_file.exists(), 'kpsc_table1_v2.json': table1_file.exists()}
    data_content = {}
    if header_file.exists():
        with open(header_file, 'r', encoding='utf-8') as f:
            data_content['header'] = json.load(f)
    if table1_file.exists():
        with open(table1_file, 'r', encoding='utf-8') as f:
            data_content['table1'] = json.load(f)
    return {'files_exist': files_exist, 'data_content': data_content}

def kpsc_validate_7_1_sheet_kpsc_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    files_exist = data['files_exist']
    data_content = data['data_content']
    has_data = False
    if 'header' in data_content:
        fields = data_content['header'].get('fields', {})
        if any((v for v in fields.values() if v)):
            has_data = True
    if 'table1' in data_content and (not has_data):
        rows = data_content['table1'].get('rows', [])
        if rows:
            has_data = True
    return {'files_status': str(files_exist), 'has_data': has_data, 'files_count': sum(files_exist.values())}

def kpsc_validate_7_1_sheet_kpsc_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nСтатус файлов парсинга: {extracted_data['files_status']}\nКоличество найденных файлов: {extracted_data['files_count']}\nЕсть ли данные в файлах: {extracted_data['has_data']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_7_1_sheet_kpsc_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_7_1_sheet_kpsc_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_7_1_sheet_kpsc_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_7_1_sheet_kpsc_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_7_1_sheet_kpsc_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_7_1_sheet_kpsc_RULE_INDEX}: {kpsc_validate_7_1_sheet_kpsc_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_7_1_sheet_kpsc_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_7_1_sheet_kpsc_RULE_INDEX}: {kpsc_validate_7_1_sheet_kpsc_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_7_1_sheet_kpsc_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_7_1_sheet_kpsc_RULE_INDEX}')
    rule = kpsc_validate_7_1_sheet_kpsc_load_rule(kpsc_validate_7_1_sheet_kpsc_RULE_INDEX)
    if args.verbose:
        kpsc_validate_7_1_sheet_kpsc_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_7_1_sheet_kpsc_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_7_1_sheet_kpsc_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_7_1_sheet_kpsc_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_7_1_sheet_kpsc_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_7_1_sheet_kpsc_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_7_1_sheet_kpsc_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_7_1_sheet_kpsc_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_7_1_sheet_kpsc_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_7_1_sheet_kpsc_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_7_1_sheet_kpsc_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_7_1_sheet_kpsc_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_7_1_sheet_kpsc_module = SimpleNamespace(RULE_INDEX=kpsc_validate_7_1_sheet_kpsc_RULE_INDEX, RULE_TITLE=kpsc_validate_7_1_sheet_kpsc_RULE_TITLE, TARGET_DOC=kpsc_validate_7_1_sheet_kpsc_TARGET_DOC, load_rule=kpsc_validate_7_1_sheet_kpsc_load_rule, load_data=kpsc_validate_7_1_sheet_kpsc_load_data, extract_relevant_data=kpsc_validate_7_1_sheet_kpsc_extract_relevant_data, build_prompt=kpsc_validate_7_1_sheet_kpsc_build_prompt, call_llm=kpsc_validate_7_1_sheet_kpsc_call_llm, save_result=kpsc_validate_7_1_sheet_kpsc_save_result, log_step=kpsc_validate_7_1_sheet_kpsc_log_step, main=kpsc_validate_7_1_sheet_kpsc_main)

# END_SOURCE_KPSC_VALIDATE_7_1_SHEET_KPSC

# START_SOURCE_KPSC_VALIDATE_7_2_SHEET_LEGEND
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_7_2_sheet_legend.py.
kpsc_validate_7_2_sheet_legend_RULE_INDEX = '7.2'

kpsc_validate_7_2_sheet_legend_RULE_TITLE = 'Наличие и заполнение листа "Условные обозначения"'

kpsc_validate_7_2_sheet_legend_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_7_2_sheet_legend_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_7_2_sheet_legend_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    legend_file = parser_outputs_dir / 'legend_v2.json'
    if not legend_file.exists():
        return {'file_exists': False, 'data': None}
    with open(legend_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'file_exists': True, 'data': data}

def kpsc_validate_7_2_sheet_legend_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_entries': False, 'entries_count': 0}
    entries = data['data'].get('entries', [])
    has_text = any((entry.get('text') for entry in entries))
    return {'file_exists': True, 'has_entries': len(entries) > 0, 'entries_count': len(entries), 'has_text': has_text}

def kpsc_validate_7_2_sheet_legend_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nЕсть записи (entries): {extracted_data['has_entries']}\nКоличество записей: {extracted_data['entries_count']}\nЕсть заполненный текст: {extracted_data.get('has_text', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_7_2_sheet_legend_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_7_2_sheet_legend_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_7_2_sheet_legend_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_7_2_sheet_legend_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_7_2_sheet_legend_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_7_2_sheet_legend_RULE_INDEX}: {kpsc_validate_7_2_sheet_legend_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_7_2_sheet_legend_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_7_2_sheet_legend_RULE_INDEX}: {kpsc_validate_7_2_sheet_legend_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_7_2_sheet_legend_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_7_2_sheet_legend_RULE_INDEX}')
    rule = kpsc_validate_7_2_sheet_legend_load_rule(kpsc_validate_7_2_sheet_legend_RULE_INDEX)
    if args.verbose:
        kpsc_validate_7_2_sheet_legend_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_7_2_sheet_legend_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_7_2_sheet_legend_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_7_2_sheet_legend_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_7_2_sheet_legend_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_7_2_sheet_legend_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_7_2_sheet_legend_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_7_2_sheet_legend_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_7_2_sheet_legend_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_7_2_sheet_legend_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_7_2_sheet_legend_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_7_2_sheet_legend_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_7_2_sheet_legend_module = SimpleNamespace(RULE_INDEX=kpsc_validate_7_2_sheet_legend_RULE_INDEX, RULE_TITLE=kpsc_validate_7_2_sheet_legend_RULE_TITLE, TARGET_DOC=kpsc_validate_7_2_sheet_legend_TARGET_DOC, load_rule=kpsc_validate_7_2_sheet_legend_load_rule, load_data=kpsc_validate_7_2_sheet_legend_load_data, extract_relevant_data=kpsc_validate_7_2_sheet_legend_extract_relevant_data, build_prompt=kpsc_validate_7_2_sheet_legend_build_prompt, call_llm=kpsc_validate_7_2_sheet_legend_call_llm, save_result=kpsc_validate_7_2_sheet_legend_save_result, log_step=kpsc_validate_7_2_sheet_legend_log_step, main=kpsc_validate_7_2_sheet_legend_main)

# END_SOURCE_KPSC_VALIDATE_7_2_SHEET_LEGEND

# START_SOURCE_KPSC_VALIDATE_7_3_SHEET_POKAZATELI
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_7_3_sheet_pokazateli.py.
kpsc_validate_7_3_sheet_pokazateli_RULE_INDEX = '7.3'

kpsc_validate_7_3_sheet_pokazateli_RULE_TITLE = 'Наличие и заполнение листа "Показатели"'

kpsc_validate_7_3_sheet_pokazateli_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_7_3_sheet_pokazateli_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_7_3_sheet_pokazateli_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    pokazateli_file = parser_outputs_dir / 'pokazateli_v3.json'
    if not pokazateli_file.exists():
        return {'file_exists': False, 'data': None}
    with open(pokazateli_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'file_exists': True, 'data': data}

def kpsc_validate_7_3_sheet_pokazateli_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    has_data = any((any((cell.get('value') for cell in row.get('cells', []))) for row in rows))
    return {'file_exists': True, 'has_rows': len(rows) > 0, 'rows_count': len(rows), 'has_data': has_data}

def kpsc_validate_7_3_sheet_pokazateli_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nЕсть строки (rows): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}\nЕсть данные в ячейках: {extracted_data.get('has_data', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_7_3_sheet_pokazateli_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_7_3_sheet_pokazateli_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_7_3_sheet_pokazateli_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_7_3_sheet_pokazateli_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_7_3_sheet_pokazateli_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_7_3_sheet_pokazateli_RULE_INDEX}: {kpsc_validate_7_3_sheet_pokazateli_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_7_3_sheet_pokazateli_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_7_3_sheet_pokazateli_RULE_INDEX}: {kpsc_validate_7_3_sheet_pokazateli_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_7_3_sheet_pokazateli_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_7_3_sheet_pokazateli_RULE_INDEX}')
    rule = kpsc_validate_7_3_sheet_pokazateli_load_rule(kpsc_validate_7_3_sheet_pokazateli_RULE_INDEX)
    if args.verbose:
        kpsc_validate_7_3_sheet_pokazateli_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_7_3_sheet_pokazateli_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_7_3_sheet_pokazateli_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_7_3_sheet_pokazateli_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_7_3_sheet_pokazateli_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_7_3_sheet_pokazateli_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_7_3_sheet_pokazateli_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_7_3_sheet_pokazateli_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_7_3_sheet_pokazateli_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_7_3_sheet_pokazateli_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_7_3_sheet_pokazateli_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_7_3_sheet_pokazateli_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_7_3_sheet_pokazateli_module = SimpleNamespace(RULE_INDEX=kpsc_validate_7_3_sheet_pokazateli_RULE_INDEX, RULE_TITLE=kpsc_validate_7_3_sheet_pokazateli_RULE_TITLE, TARGET_DOC=kpsc_validate_7_3_sheet_pokazateli_TARGET_DOC, load_rule=kpsc_validate_7_3_sheet_pokazateli_load_rule, load_data=kpsc_validate_7_3_sheet_pokazateli_load_data, extract_relevant_data=kpsc_validate_7_3_sheet_pokazateli_extract_relevant_data, build_prompt=kpsc_validate_7_3_sheet_pokazateli_build_prompt, call_llm=kpsc_validate_7_3_sheet_pokazateli_call_llm, save_result=kpsc_validate_7_3_sheet_pokazateli_save_result, log_step=kpsc_validate_7_3_sheet_pokazateli_log_step, main=kpsc_validate_7_3_sheet_pokazateli_main)

# END_SOURCE_KPSC_VALIDATE_7_3_SHEET_POKAZATELI

# START_SOURCE_KPSC_VALIDATE_7_4_SHEET_OCIFROVKA
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_7_4_sheet_ocifrovka.py.
kpsc_validate_7_4_sheet_ocifrovka_RULE_INDEX = '7.4'

kpsc_validate_7_4_sheet_ocifrovka_RULE_TITLE = 'Наличие и заполнение листа "Оцифровка потерь КПСЦ"'

kpsc_validate_7_4_sheet_ocifrovka_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_7_4_sheet_ocifrovka_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_7_4_sheet_ocifrovka_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    ocifrovka_file = parser_outputs_dir / 'ocifrovka_poteri_v2.json'
    if not ocifrovka_file.exists():
        return {'file_exists': False, 'data': None}
    with open(ocifrovka_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'file_exists': True, 'data': data}

def kpsc_validate_7_4_sheet_ocifrovka_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    return {'file_exists': True, 'has_rows': len(rows) >= 2, 'rows_count': len(rows)}

def kpsc_validate_7_4_sheet_ocifrovka_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nЕсть строки (минимум 2): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_7_4_sheet_ocifrovka_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_7_4_sheet_ocifrovka_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_7_4_sheet_ocifrovka_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_7_4_sheet_ocifrovka_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_7_4_sheet_ocifrovka_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_7_4_sheet_ocifrovka_RULE_INDEX}: {kpsc_validate_7_4_sheet_ocifrovka_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_7_4_sheet_ocifrovka_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_7_4_sheet_ocifrovka_RULE_INDEX}: {kpsc_validate_7_4_sheet_ocifrovka_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_7_4_sheet_ocifrovka_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_7_4_sheet_ocifrovka_RULE_INDEX}')
    rule = kpsc_validate_7_4_sheet_ocifrovka_load_rule(kpsc_validate_7_4_sheet_ocifrovka_RULE_INDEX)
    if args.verbose:
        kpsc_validate_7_4_sheet_ocifrovka_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_7_4_sheet_ocifrovka_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_7_4_sheet_ocifrovka_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_7_4_sheet_ocifrovka_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_7_4_sheet_ocifrovka_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_7_4_sheet_ocifrovka_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_7_4_sheet_ocifrovka_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_7_4_sheet_ocifrovka_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_7_4_sheet_ocifrovka_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_7_4_sheet_ocifrovka_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_7_4_sheet_ocifrovka_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_7_4_sheet_ocifrovka_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_7_4_sheet_ocifrovka_module = SimpleNamespace(RULE_INDEX=kpsc_validate_7_4_sheet_ocifrovka_RULE_INDEX, RULE_TITLE=kpsc_validate_7_4_sheet_ocifrovka_RULE_TITLE, TARGET_DOC=kpsc_validate_7_4_sheet_ocifrovka_TARGET_DOC, load_rule=kpsc_validate_7_4_sheet_ocifrovka_load_rule, load_data=kpsc_validate_7_4_sheet_ocifrovka_load_data, extract_relevant_data=kpsc_validate_7_4_sheet_ocifrovka_extract_relevant_data, build_prompt=kpsc_validate_7_4_sheet_ocifrovka_build_prompt, call_llm=kpsc_validate_7_4_sheet_ocifrovka_call_llm, save_result=kpsc_validate_7_4_sheet_ocifrovka_save_result, log_step=kpsc_validate_7_4_sheet_ocifrovka_log_step, main=kpsc_validate_7_4_sheet_ocifrovka_main)

# END_SOURCE_KPSC_VALIDATE_7_4_SHEET_OCIFROVKA

# START_SOURCE_KPSC_VALIDATE_7_5_SHEET_PA1
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_7_5_sheet_pa1.py.
kpsc_validate_7_5_sheet_pa1_RULE_INDEX = '7.5'

kpsc_validate_7_5_sheet_pa1_RULE_TITLE = 'Наличие и заполнение листа "ПА-1"'

kpsc_validate_7_5_sheet_pa1_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_7_5_sheet_pa1_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_7_5_sheet_pa1_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    pa1_table_file = parser_outputs_dir / 'pa1_table_v1.json'
    pa1_chart_file = parser_outputs_dir / 'pa1_chart_v3.json'
    files_exist = {'pa1_table_v1.json': pa1_table_file.exists(), 'pa1_chart_v3.json': pa1_chart_file.exists()}
    data_content = {}
    if pa1_table_file.exists():
        with open(pa1_table_file, 'r', encoding='utf-8') as f:
            data_content['table'] = json.load(f)
    if pa1_chart_file.exists():
        with open(pa1_chart_file, 'r', encoding='utf-8') as f:
            data_content['chart'] = json.load(f)
    return {'files_exist': files_exist, 'data_content': data_content}

def kpsc_validate_7_5_sheet_pa1_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    files_exist = data['files_exist']
    data_content = data['data_content']
    has_data = len(data_content) > 0
    return {'files_status': str(files_exist), 'files_count': sum(files_exist.values()), 'has_data': has_data}

def kpsc_validate_7_5_sheet_pa1_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nСтатус файлов: {extracted_data['files_status']}\nКоличество найденных файлов: {extracted_data['files_count']}\nЕсть данные: {extracted_data['has_data']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_7_5_sheet_pa1_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_7_5_sheet_pa1_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_7_5_sheet_pa1_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_7_5_sheet_pa1_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_7_5_sheet_pa1_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_7_5_sheet_pa1_RULE_INDEX}: {kpsc_validate_7_5_sheet_pa1_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_7_5_sheet_pa1_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_7_5_sheet_pa1_RULE_INDEX}: {kpsc_validate_7_5_sheet_pa1_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_7_5_sheet_pa1_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_7_5_sheet_pa1_RULE_INDEX}')
    rule = kpsc_validate_7_5_sheet_pa1_load_rule(kpsc_validate_7_5_sheet_pa1_RULE_INDEX)
    if args.verbose:
        kpsc_validate_7_5_sheet_pa1_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_7_5_sheet_pa1_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_7_5_sheet_pa1_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_7_5_sheet_pa1_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_7_5_sheet_pa1_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_7_5_sheet_pa1_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_7_5_sheet_pa1_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_7_5_sheet_pa1_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_7_5_sheet_pa1_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_7_5_sheet_pa1_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_7_5_sheet_pa1_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_7_5_sheet_pa1_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_7_5_sheet_pa1_module = SimpleNamespace(RULE_INDEX=kpsc_validate_7_5_sheet_pa1_RULE_INDEX, RULE_TITLE=kpsc_validate_7_5_sheet_pa1_RULE_TITLE, TARGET_DOC=kpsc_validate_7_5_sheet_pa1_TARGET_DOC, load_rule=kpsc_validate_7_5_sheet_pa1_load_rule, load_data=kpsc_validate_7_5_sheet_pa1_load_data, extract_relevant_data=kpsc_validate_7_5_sheet_pa1_extract_relevant_data, build_prompt=kpsc_validate_7_5_sheet_pa1_build_prompt, call_llm=kpsc_validate_7_5_sheet_pa1_call_llm, save_result=kpsc_validate_7_5_sheet_pa1_save_result, log_step=kpsc_validate_7_5_sheet_pa1_log_step, main=kpsc_validate_7_5_sheet_pa1_main)

# END_SOURCE_KPSC_VALIDATE_7_5_SHEET_PA1

# START_SOURCE_KPSC_VALIDATE_7_6_SHEET_SPAGHETTI
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_7_6_sheet_spaghetti.py.
kpsc_validate_7_6_sheet_spaghetti_RULE_INDEX = '7.6'

kpsc_validate_7_6_sheet_spaghetti_RULE_TITLE = 'Наличие и заполнение листа "Диаграмма Спагетти" или "ДС"'

kpsc_validate_7_6_sheet_spaghetti_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_7_6_sheet_spaghetti_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_7_6_sheet_spaghetti_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    spaghetti_file = parser_outputs_dir / 'spaghetti_sheet_v2.json'
    if not spaghetti_file.exists():
        return {'file_exists': False, 'data': None}
    with open(spaghetti_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'file_exists': True, 'data': data}

def kpsc_validate_7_6_sheet_spaghetti_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    return {'file_exists': True, 'has_rows': len(rows) > 0, 'rows_count': len(rows)}

def kpsc_validate_7_6_sheet_spaghetti_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nЕсть строки (rows): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_7_6_sheet_spaghetti_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_7_6_sheet_spaghetti_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_7_6_sheet_spaghetti_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_7_6_sheet_spaghetti_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_7_6_sheet_spaghetti_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_7_6_sheet_spaghetti_RULE_INDEX}: {kpsc_validate_7_6_sheet_spaghetti_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_7_6_sheet_spaghetti_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_7_6_sheet_spaghetti_RULE_INDEX}: {kpsc_validate_7_6_sheet_spaghetti_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_7_6_sheet_spaghetti_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_7_6_sheet_spaghetti_RULE_INDEX}')
    rule = kpsc_validate_7_6_sheet_spaghetti_load_rule(kpsc_validate_7_6_sheet_spaghetti_RULE_INDEX)
    if args.verbose:
        kpsc_validate_7_6_sheet_spaghetti_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_7_6_sheet_spaghetti_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_7_6_sheet_spaghetti_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_7_6_sheet_spaghetti_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_7_6_sheet_spaghetti_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_7_6_sheet_spaghetti_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_7_6_sheet_spaghetti_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_7_6_sheet_spaghetti_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_7_6_sheet_spaghetti_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_7_6_sheet_spaghetti_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_7_6_sheet_spaghetti_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_7_6_sheet_spaghetti_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_7_6_sheet_spaghetti_module = SimpleNamespace(RULE_INDEX=kpsc_validate_7_6_sheet_spaghetti_RULE_INDEX, RULE_TITLE=kpsc_validate_7_6_sheet_spaghetti_RULE_TITLE, TARGET_DOC=kpsc_validate_7_6_sheet_spaghetti_TARGET_DOC, load_rule=kpsc_validate_7_6_sheet_spaghetti_load_rule, load_data=kpsc_validate_7_6_sheet_spaghetti_load_data, extract_relevant_data=kpsc_validate_7_6_sheet_spaghetti_extract_relevant_data, build_prompt=kpsc_validate_7_6_sheet_spaghetti_build_prompt, call_llm=kpsc_validate_7_6_sheet_spaghetti_call_llm, save_result=kpsc_validate_7_6_sheet_spaghetti_save_result, log_step=kpsc_validate_7_6_sheet_spaghetti_log_step, main=kpsc_validate_7_6_sheet_spaghetti_main)

# END_SOURCE_KPSC_VALIDATE_7_6_SHEET_SPAGHETTI

# START_SOURCE_KPSC_VALIDATE_7_7_SHEET_SPAGHETTI_PROBLEMS
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_7_7_sheet_spaghetti_problems.py.
kpsc_validate_7_7_sheet_spaghetti_problems_RULE_INDEX = '7.7'

kpsc_validate_7_7_sheet_spaghetti_problems_RULE_TITLE = 'Наличие и заполнение листа "Перечень проблем по Диаграмме Спагетти"'

kpsc_validate_7_7_sheet_spaghetti_problems_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_7_7_sheet_spaghetti_problems_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_7_7_sheet_spaghetti_problems_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    spaghetti_problems_file = parser_outputs_dir / 'spaghetti_problems_v1.json'
    if not spaghetti_problems_file.exists():
        return {'file_exists': False, 'data': None}
    with open(spaghetti_problems_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'file_exists': True, 'data': data}

def kpsc_validate_7_7_sheet_spaghetti_problems_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['file_exists']:
        return {'file_exists': False, 'has_rows': False, 'rows_count': 0}
    rows = data['data'].get('rows', [])
    return {'file_exists': True, 'has_rows': len(rows) >= 2, 'rows_count': len(rows)}

def kpsc_validate_7_7_sheet_spaghetti_problems_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл существует: {extracted_data['file_exists']}\nЕсть строки (минимум 2): {extracted_data['has_rows']}\nКоличество строк: {extracted_data['rows_count']}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_7_7_sheet_spaghetti_problems_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_7_7_sheet_spaghetti_problems_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_7_7_sheet_spaghetti_problems_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_7_7_sheet_spaghetti_problems_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_7_7_sheet_spaghetti_problems_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_7_7_sheet_spaghetti_problems_RULE_INDEX}: {kpsc_validate_7_7_sheet_spaghetti_problems_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_7_7_sheet_spaghetti_problems_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_7_7_sheet_spaghetti_problems_RULE_INDEX}: {kpsc_validate_7_7_sheet_spaghetti_problems_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_7_7_sheet_spaghetti_problems_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_7_7_sheet_spaghetti_problems_RULE_INDEX}')
    rule = kpsc_validate_7_7_sheet_spaghetti_problems_load_rule(kpsc_validate_7_7_sheet_spaghetti_problems_RULE_INDEX)
    if args.verbose:
        kpsc_validate_7_7_sheet_spaghetti_problems_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_7_7_sheet_spaghetti_problems_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_7_7_sheet_spaghetti_problems_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_7_7_sheet_spaghetti_problems_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_7_7_sheet_spaghetti_problems_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_7_7_sheet_spaghetti_problems_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_7_7_sheet_spaghetti_problems_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_7_7_sheet_spaghetti_problems_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_7_7_sheet_spaghetti_problems_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_7_7_sheet_spaghetti_problems_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_7_7_sheet_spaghetti_problems_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_7_7_sheet_spaghetti_problems_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_7_7_sheet_spaghetti_problems_module = SimpleNamespace(RULE_INDEX=kpsc_validate_7_7_sheet_spaghetti_problems_RULE_INDEX, RULE_TITLE=kpsc_validate_7_7_sheet_spaghetti_problems_RULE_TITLE, TARGET_DOC=kpsc_validate_7_7_sheet_spaghetti_problems_TARGET_DOC, load_rule=kpsc_validate_7_7_sheet_spaghetti_problems_load_rule, load_data=kpsc_validate_7_7_sheet_spaghetti_problems_load_data, extract_relevant_data=kpsc_validate_7_7_sheet_spaghetti_problems_extract_relevant_data, build_prompt=kpsc_validate_7_7_sheet_spaghetti_problems_build_prompt, call_llm=kpsc_validate_7_7_sheet_spaghetti_problems_call_llm, save_result=kpsc_validate_7_7_sheet_spaghetti_problems_save_result, log_step=kpsc_validate_7_7_sheet_spaghetti_problems_log_step, main=kpsc_validate_7_7_sheet_spaghetti_problems_main)

# END_SOURCE_KPSC_VALIDATE_7_7_SHEET_SPAGHETTI_PROBLEMS

# START_SOURCE_KPSC_VALIDATE_7_8_SHEET_TAKT_TIME
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_7_8_sheet_takt_time.py.
kpsc_validate_7_8_sheet_takt_time_RULE_INDEX = '7.8'

kpsc_validate_7_8_sheet_takt_time_RULE_TITLE = 'Наличие и заполнение листа "Расчет такта" / "Расчет времени такта"'

kpsc_validate_7_8_sheet_takt_time_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_7_8_sheet_takt_time_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_7_8_sheet_takt_time_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    header_file = parser_outputs_dir / 'kpsc_header_v2.json'
    if not header_file.exists():
        return {'header_exists': False, 'data': None}
    with open(header_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return {'header_exists': True, 'data': data}

def kpsc_validate_7_8_sheet_takt_time_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['header_exists']:
        return {'header_exists': False, 'takt_time': None}
    takt_time = data['data'].get('fields', {}).get('takt_time', '')
    return {'header_exists': True, 'takt_time': takt_time, 'has_takt_time': bool(takt_time)}

def kpsc_validate_7_8_sheet_takt_time_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайл заголовка существует: {extracted_data['header_exists']}\nПоле takt_time: "{extracted_data.get('takt_time', '')}"\nПоле заполнено: {extracted_data.get('has_takt_time', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_7_8_sheet_takt_time_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_7_8_sheet_takt_time_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_7_8_sheet_takt_time_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_7_8_sheet_takt_time_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_7_8_sheet_takt_time_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_7_8_sheet_takt_time_RULE_INDEX}: {kpsc_validate_7_8_sheet_takt_time_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_7_8_sheet_takt_time_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_7_8_sheet_takt_time_RULE_INDEX}: {kpsc_validate_7_8_sheet_takt_time_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_7_8_sheet_takt_time_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_7_8_sheet_takt_time_RULE_INDEX}')
    rule = kpsc_validate_7_8_sheet_takt_time_load_rule(kpsc_validate_7_8_sheet_takt_time_RULE_INDEX)
    if args.verbose:
        kpsc_validate_7_8_sheet_takt_time_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_7_8_sheet_takt_time_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_7_8_sheet_takt_time_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded\nData: {json.dumps(data, ensure_ascii=False, indent=2)}')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_7_8_sheet_takt_time_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_7_8_sheet_takt_time_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_7_8_sheet_takt_time_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_7_8_sheet_takt_time_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_7_8_sheet_takt_time_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_7_8_sheet_takt_time_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_7_8_sheet_takt_time_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_7_8_sheet_takt_time_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_7_8_sheet_takt_time_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_7_8_sheet_takt_time_module = SimpleNamespace(RULE_INDEX=kpsc_validate_7_8_sheet_takt_time_RULE_INDEX, RULE_TITLE=kpsc_validate_7_8_sheet_takt_time_RULE_TITLE, TARGET_DOC=kpsc_validate_7_8_sheet_takt_time_TARGET_DOC, load_rule=kpsc_validate_7_8_sheet_takt_time_load_rule, load_data=kpsc_validate_7_8_sheet_takt_time_load_data, extract_relevant_data=kpsc_validate_7_8_sheet_takt_time_extract_relevant_data, build_prompt=kpsc_validate_7_8_sheet_takt_time_build_prompt, call_llm=kpsc_validate_7_8_sheet_takt_time_call_llm, save_result=kpsc_validate_7_8_sheet_takt_time_save_result, log_step=kpsc_validate_7_8_sheet_takt_time_log_step, main=kpsc_validate_7_8_sheet_takt_time_main)

# END_SOURCE_KPSC_VALIDATE_7_8_SHEET_TAKT_TIME

# START_SOURCE_KPSC_VALIDATE_8_1_UNITS_CROSS_CHECK
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_8_1_units_cross_check.py.
kpsc_validate_8_1_units_cross_check_RULE_INDEX = '8.1'

kpsc_validate_8_1_units_cross_check_RULE_TITLE = 'Соответствие единиц измерения между листами "Показатели" и "КПСЦ"'

kpsc_validate_8_1_units_cross_check_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_8_1_units_cross_check_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_8_1_units_cross_check_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    pokazateli_file = parser_outputs_dir / 'pokazateli_v3.json'
    kpsc_table_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not pokazateli_file.exists() or not kpsc_table_file.exists():
        return {'pokazateli_exists': pokazateli_file.exists(), 'kpsc_exists': kpsc_table_file.exists(), 'pokazateli_data': None, 'kpsc_data': None}
    with open(pokazateli_file, 'r', encoding='utf-8') as f:
        pokazateli_data = json.load(f)
    with open(kpsc_table_file, 'r', encoding='utf-8') as f:
        kpsc_data = json.load(f)
    return {'pokazateli_exists': True, 'kpsc_exists': True, 'pokazateli_data': pokazateli_data, 'kpsc_data': kpsc_data}

def kpsc_validate_8_1_units_cross_check_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['pokazateli_exists'] or not data['kpsc_exists']:
        return {'files_exist': False, 'mismatches': []}
    pokazateli_rows = data['pokazateli_data'].get('rows', [])
    pokazateli_units = {}
    for row in pokazateli_rows:
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(3, '')
        if not pokazatel_name or pokazatel_name == 'Показатель':
            continue
        unit = cells_dict.get(4, '')
        if unit:
            pokazateli_units[str(pokazatel_name).strip().lower()] = str(unit).strip()
    kpsc_rows = data['kpsc_data'].get('rows', [])
    kpsc_units = {}
    for i, row in enumerate(kpsc_rows):
        if i == 0:
            continue
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(2, '') or cells_dict.get(3, '')
        if not pokazatel_name:
            continue
        unit = cells_dict.get(4, '') or cells_dict.get(27, '')
        if unit and pokazatel_name:
            kpsc_units[str(pokazatel_name).strip().lower()] = str(unit).strip()
    mismatches = []
    for pokazatel, unit_pokazateli in pokazateli_units.items():
        if pokazatel in kpsc_units:
            unit_kpsc = kpsc_units[pokazatel]
            if unit_pokazateli.lower() != unit_kpsc.lower():
                mismatches.append({'pokazatel': pokazatel, 'unit_pokazateli': unit_pokazateli, 'unit_kpsc': unit_kpsc})
    return {'files_exist': True, 'pokazateli_count': len(pokazateli_units), 'kpsc_count': len(kpsc_units), 'pokazateli_units': pokazateli_units, 'kpsc_units': kpsc_units, 'mismatches': mismatches, 'has_mismatches': len(mismatches) > 0}

def kpsc_validate_8_1_units_cross_check_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайлы существуют: {extracted_data['files_exist']}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "ПОКАЗАТЕЛИ" (всего {extracted_data.get('pokazateli_count', 0)}):\n{json.dumps(extracted_data.get('pokazateli_units', {}), ensure_ascii=False, indent=2)}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "КПСЦ" БЛОК "РАСЧЕТ ВПП" (всего {extracted_data.get('kpsc_count', 0)}):\n{json.dumps(extracted_data.get('kpsc_units', {}), ensure_ascii=False, indent=2)}\n\nНЕСООТВЕТСТВИЯ ЕДИНИЦ ИЗМЕРЕНИЯ:\n{json.dumps(extracted_data.get('mismatches', []), ensure_ascii=False, indent=2)}\n\nЕсть несоответствия: {extracted_data.get('has_mismatches', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_8_1_units_cross_check_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_8_1_units_cross_check_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_8_1_units_cross_check_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_8_1_units_cross_check_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_8_1_units_cross_check_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_8_1_units_cross_check_RULE_INDEX}: {kpsc_validate_8_1_units_cross_check_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_8_1_units_cross_check_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_8_1_units_cross_check_RULE_INDEX}: {kpsc_validate_8_1_units_cross_check_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_8_1_units_cross_check_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_8_1_units_cross_check_RULE_INDEX}')
    rule = kpsc_validate_8_1_units_cross_check_load_rule(kpsc_validate_8_1_units_cross_check_RULE_INDEX)
    if args.verbose:
        kpsc_validate_8_1_units_cross_check_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_8_1_units_cross_check_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_8_1_units_cross_check_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_8_1_units_cross_check_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_8_1_units_cross_check_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_8_1_units_cross_check_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_8_1_units_cross_check_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_8_1_units_cross_check_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_8_1_units_cross_check_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_8_1_units_cross_check_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_8_1_units_cross_check_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_8_1_units_cross_check_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_8_1_units_cross_check_module = SimpleNamespace(RULE_INDEX=kpsc_validate_8_1_units_cross_check_RULE_INDEX, RULE_TITLE=kpsc_validate_8_1_units_cross_check_RULE_TITLE, TARGET_DOC=kpsc_validate_8_1_units_cross_check_TARGET_DOC, load_rule=kpsc_validate_8_1_units_cross_check_load_rule, load_data=kpsc_validate_8_1_units_cross_check_load_data, extract_relevant_data=kpsc_validate_8_1_units_cross_check_extract_relevant_data, build_prompt=kpsc_validate_8_1_units_cross_check_build_prompt, call_llm=kpsc_validate_8_1_units_cross_check_call_llm, save_result=kpsc_validate_8_1_units_cross_check_save_result, log_step=kpsc_validate_8_1_units_cross_check_log_step, main=kpsc_validate_8_1_units_cross_check_main)

# END_SOURCE_KPSC_VALIDATE_8_1_UNITS_CROSS_CHECK

# START_SOURCE_KPSC_VALIDATE_8_2_VALUES_CROSS_CHECK
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_8_2_values_cross_check.py.
kpsc_validate_8_2_values_cross_check_RULE_INDEX = '8.2'

kpsc_validate_8_2_values_cross_check_RULE_TITLE = 'Соответствие значений из "Показатели" с колонкой ИТОГО в "КПСЦ"'

kpsc_validate_8_2_values_cross_check_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_8_2_values_cross_check_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_8_2_values_cross_check_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    pokazateli_file = parser_outputs_dir / 'pokazateli_v3.json'
    kpsc_table_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not pokazateli_file.exists() or not kpsc_table_file.exists():
        return {'pokazateli_exists': pokazateli_file.exists(), 'kpsc_exists': kpsc_table_file.exists(), 'pokazateli_data': None, 'kpsc_data': None}
    with open(pokazateli_file, 'r', encoding='utf-8') as f:
        pokazateli_data = json.load(f)
    with open(kpsc_table_file, 'r', encoding='utf-8') as f:
        kpsc_data = json.load(f)
    return {'pokazateli_exists': True, 'kpsc_exists': True, 'pokazateli_data': pokazateli_data, 'kpsc_data': kpsc_data}

def kpsc_validate_8_2_values_cross_check_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['pokazateli_exists'] or not data['kpsc_exists']:
        return {'files_exist': False, 'mismatches': []}
    pokazateli_rows = data['pokazateli_data'].get('rows', [])
    pokazateli_values = {}
    for row in pokazateli_rows:
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(3, '')
        if not pokazatel_name or pokazatel_name == 'Показатель':
            continue
        value_e = cells_dict.get(5)
        value_g = cells_dict.get(7)
        value = value_e if value_e is not None else value_g
        if value is not None:
            try:
                pokazateli_values[str(pokazatel_name).strip().lower()] = float(value)
            except (ValueError, TypeError):
                pass
    kpsc_rows = data['kpsc_data'].get('rows', [])
    kpsc_itogo_values = {}
    itogo_col = None
    if kpsc_rows:
        header_cells = {cell.get('col'): cell.get('value') for cell in kpsc_rows[0].get('cells', [])}
        for col_idx, val in header_cells.items():
            if isinstance(val, str) and 'итого' in val.lower():
                itogo_col = col_idx
                break
    for i, row in enumerate(kpsc_rows):
        if i == 0:
            continue
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(2, '') or cells_dict.get(3, '')
        if not pokazatel_name:
            continue
        itogo_value = cells_dict.get(itogo_col) if itogo_col else None
        if itogo_value is not None and pokazatel_name:
            try:
                kpsc_itogo_values[str(pokazatel_name).strip().lower()] = float(itogo_value)
            except (ValueError, TypeError):
                pass
    mismatches = []
    tolerance = 0.01
    for pokazatel, value_pokazateli in pokazateli_values.items():
        if pokazatel in kpsc_itogo_values:
            value_kpsc = kpsc_itogo_values[pokazatel]
            if abs(value_pokazateli - value_kpsc) > tolerance:
                mismatches.append({'pokazatel': pokazatel, 'value_pokazateli': value_pokazateli, 'value_kpsc_itogo': value_kpsc, 'difference': abs(value_pokazateli - value_kpsc)})
    return {'files_exist': True, 'pokazateli_count': len(pokazateli_values), 'kpsc_count': len(kpsc_itogo_values), 'pokazateli_values': pokazateli_values, 'kpsc_itogo_values': kpsc_itogo_values, 'mismatches': mismatches, 'has_mismatches': len(mismatches) > 0}

def kpsc_validate_8_2_values_cross_check_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайлы существуют: {extracted_data['files_exist']}\n\nЗНАЧЕНИЯ ИЗ ЛИСТА "ПОКАЗАТЕЛИ" (всего {extracted_data.get('pokazateli_count', 0)}):\n{json.dumps(extracted_data.get('pokazateli_values', {}), ensure_ascii=False, indent=2)}\n\nЗНАЧЕНИЯ ИТОГО ИЗ ЛИСТА "КПСЦ" БЛОК "РАСЧЕТ ВПП" (всего {extracted_data.get('kpsc_count', 0)}):\n{json.dumps(extracted_data.get('kpsc_itogo_values', {}), ensure_ascii=False, indent=2)}\n\nНЕСООТВЕТСТВИЯ ЗНАЧЕНИЙ (погрешность > 0.01):\n{json.dumps(extracted_data.get('mismatches', []), ensure_ascii=False, indent=2)}\n\nЕсть несоответствия: {extracted_data.get('has_mismatches', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_8_2_values_cross_check_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_8_2_values_cross_check_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_8_2_values_cross_check_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_8_2_values_cross_check_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_8_2_values_cross_check_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_8_2_values_cross_check_RULE_INDEX}: {kpsc_validate_8_2_values_cross_check_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_8_2_values_cross_check_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_8_2_values_cross_check_RULE_INDEX}: {kpsc_validate_8_2_values_cross_check_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_8_2_values_cross_check_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_8_2_values_cross_check_RULE_INDEX}')
    rule = kpsc_validate_8_2_values_cross_check_load_rule(kpsc_validate_8_2_values_cross_check_RULE_INDEX)
    if args.verbose:
        kpsc_validate_8_2_values_cross_check_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_8_2_values_cross_check_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_8_2_values_cross_check_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_8_2_values_cross_check_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_8_2_values_cross_check_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_8_2_values_cross_check_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_8_2_values_cross_check_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_8_2_values_cross_check_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_8_2_values_cross_check_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_8_2_values_cross_check_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_8_2_values_cross_check_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_8_2_values_cross_check_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_8_2_values_cross_check_module = SimpleNamespace(RULE_INDEX=kpsc_validate_8_2_values_cross_check_RULE_INDEX, RULE_TITLE=kpsc_validate_8_2_values_cross_check_RULE_TITLE, TARGET_DOC=kpsc_validate_8_2_values_cross_check_TARGET_DOC, load_rule=kpsc_validate_8_2_values_cross_check_load_rule, load_data=kpsc_validate_8_2_values_cross_check_load_data, extract_relevant_data=kpsc_validate_8_2_values_cross_check_extract_relevant_data, build_prompt=kpsc_validate_8_2_values_cross_check_build_prompt, call_llm=kpsc_validate_8_2_values_cross_check_call_llm, save_result=kpsc_validate_8_2_values_cross_check_save_result, log_step=kpsc_validate_8_2_values_cross_check_log_step, main=kpsc_validate_8_2_values_cross_check_main)

# END_SOURCE_KPSC_VALIDATE_8_2_VALUES_CROSS_CHECK

# START_SOURCE_KPSC_VALIDATE_8_3_INDICATORS_CROSS_CHECK
# PURPOSE: Inlined source from audit_engine/kpsc/validation_scripts/validate_8_3_indicators_cross_check.py.
kpsc_validate_8_3_indicators_cross_check_RULE_INDEX = '8.3'

kpsc_validate_8_3_indicators_cross_check_RULE_TITLE = 'Соответствие названий показателей между листами "Показатели" и "КПСЦ"'

kpsc_validate_8_3_indicators_cross_check_TARGET_DOC = 'КПСЦ и Спагетти ТС_Предприятие.xlsx'

def kpsc_validate_8_3_indicators_cross_check_load_rule(rule_index: str) -> dict:
    """Загрузка правила проверки из JSON"""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).parent.parent / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено в validation_rules.json')

def kpsc_validate_8_3_indicators_cross_check_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка необходимых JSON данных"""
    pokazateli_file = parser_outputs_dir / 'pokazateli_v3.json'
    kpsc_table_file = parser_outputs_dir / 'kpsc_table1_v2.json'
    if not pokazateli_file.exists() or not kpsc_table_file.exists():
        return {'pokazateli_exists': pokazateli_file.exists(), 'kpsc_exists': kpsc_table_file.exists(), 'pokazateli_data': None, 'kpsc_data': None}
    with open(pokazateli_file, 'r', encoding='utf-8') as f:
        pokazateli_data = json.load(f)
    with open(kpsc_table_file, 'r', encoding='utf-8') as f:
        kpsc_data = json.load(f)
    return {'pokazateli_exists': True, 'kpsc_exists': True, 'pokazateli_data': pokazateli_data, 'kpsc_data': kpsc_data}

def kpsc_validate_8_3_indicators_cross_check_extract_relevant_data(data: dict) -> dict:
    """Извлечение данных для проверки"""
    if not data['pokazateli_exists'] or not data['kpsc_exists']:
        return {'files_exist': False, 'missing_indicators': []}
    pokazateli_rows = data['pokazateli_data'].get('rows', [])
    pokazateli_indicators = []
    for row in pokazateli_rows:
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(3, '')
        if not pokazatel_name or pokazatel_name == 'Показатель':
            continue
        pokazateli_indicators.append(str(pokazatel_name).strip())
    kpsc_rows = data['kpsc_data'].get('rows', [])
    kpsc_indicators = []
    kpsc_indicators_normalized = set()
    for i, row in enumerate(kpsc_rows):
        if i == 0:
            continue
        cells_dict = {cell.get('col'): cell.get('value') for cell in row.get('cells', [])}
        pokazatel_name = cells_dict.get(2, '') or cells_dict.get(3, '')
        if pokazatel_name:
            original_name = str(pokazatel_name).strip()
            kpsc_indicators.append(original_name)
            kpsc_indicators_normalized.add(original_name.lower())
    ALIASES = {'время протекания процесса': ['впп'], 'выработка': ['выпуск продукции'], 'объем выпускаемой продукции': ['выпуск продукции']}
    missing_indicators = []
    for indicator in pokazateli_indicators:
        ind_lower = indicator.lower()
        found = ind_lower in kpsc_indicators_normalized or any((ind_lower in kpsc_name for kpsc_name in kpsc_indicators_normalized))
        if not found:
            for alias in ALIASES.get(ind_lower, []):
                if any((alias in kpsc_name for kpsc_name in kpsc_indicators_normalized)):
                    found = True
                    break
        if not found:
            missing_indicators.append(indicator)
    return {'files_exist': True, 'pokazateli_count': len(pokazateli_indicators), 'kpsc_count': len(kpsc_indicators), 'pokazateli_indicators': pokazateli_indicators, 'kpsc_indicators': kpsc_indicators, 'missing_indicators': missing_indicators, 'has_missing': len(missing_indicators) > 0}

def kpsc_validate_8_3_indicators_cross_check_build_prompt(extracted_data: dict, rule: dict) -> str:
    """Формирование промпта для LLM с динамической подстановкой правила"""
    prompt = f'\nТы эксперт по проверке документов КПСЦ (Карта Потока Создания Ценности).\n\nТРЕБОВАНИЕ ЭКСПЕРТА:\n{rule['requirement_expert']}\n\nТЕХНИЧЕСКОЕ ОПИСАНИЕ:\n{rule['technical_description']}\n\nКРИТЕРИИ ПРОВЕРКИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nФайлы существуют: {extracted_data['files_exist']}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "ПОКАЗАТЕЛИ" (всего {extracted_data.get('pokazateli_count', 0)}):\n{json.dumps(extracted_data.get('pokazateli_indicators', []), ensure_ascii=False, indent=2)}\n\nПОКАЗАТЕЛИ ИЗ ЛИСТА "КПСЦ" БЛОК "РАСЧЕТ ВПП" (всего {extracted_data.get('kpsc_count', 0)}):\n{json.dumps(extracted_data.get('kpsc_indicators', []), ensure_ascii=False, indent=2)}\n\nОТСУТСТВУЮЩИЕ ПОКАЗАТЕЛИ (есть в "Показатели", но нет в "КПСЦ"):\n{json.dumps(extracted_data.get('missing_indicators', []), ensure_ascii=False, indent=2)}\n\nЕсть отсутствующие показатели: {extracted_data.get('has_missing', False)}\n\nЗАДАНИЕ:\nПроверь соответствие фактических данных требованию эксперта и критериям проверки.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{rule['rule_index']}",\n  "rule_title": "{rule['rule_title']}",\n  "target_document": "{kpsc_validate_8_3_indicators_cross_check_TARGET_DOC}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n\nВерни только JSON, без дополнительного текста.\n'
    return prompt

def kpsc_validate_8_3_indicators_cross_check_call_llm(prompt: str, api_key: str) -> dict:
    """Вызов LLM через единый клиент `src.llm.client.call_llm`.

    Использует `LLM_CONFIG` (Pydantic-конфиг) для URL и модели по умолчанию;
    `response_format={'type': 'json_object'}` форсит строгий JSON.
    """
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kpsc_validate_8_3_indicators_cross_check_save_result(result: dict, output_file: Path):
    """Сохранение результата"""
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f'✓ Результат сохранен: {output_file}')

def kpsc_validate_8_3_indicators_cross_check_log_step(log_file, step_num: int, step_title: str, content: str):
    """Запись шага в лог"""
    separator = '=' * 80
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f'\n{separator}\n')
        f.write(f'[STEP {step_num}/6] {step_title}\n')
        f.write(f'Timestamp: {timestamp}\n')
        f.write(f'{separator}\n')
        f.write(f'{content}\n')

def kpsc_validate_8_3_indicators_cross_check_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kpsc_validate_8_3_indicators_cross_check_RULE_INDEX}: {kpsc_validate_8_3_indicators_cross_check_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True, help='Путь к папке с результатами парсинга (parser_outputs)')
    parser.add_argument('--output', type=Path, required=True, help='Путь для сохранения результата валидации (JSON файл)')
    parser.add_argument('--verbose', action='store_true', help='Включить детальное логирование')
    parser.add_argument('--log-file', type=Path, help='Путь к файлу лога (если не указан, используется validation_logs/validate_X_Y.log)')
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    log_file = args.log_file
    if args.verbose and (not log_file):
        log_dir = Path(__file__).parent.parent / 'validation_logs'
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / f'validate_{kpsc_validate_8_3_indicators_cross_check_RULE_INDEX.replace('.', '_')}.log'
    if args.verbose and log_file:
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write(f'ВАЛИДАТОР {kpsc_validate_8_3_indicators_cross_check_RULE_INDEX}: {kpsc_validate_8_3_indicators_cross_check_RULE_TITLE}\n')
            f.write(f'Запуск: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n')
            f.write(f'Целевой документ: {kpsc_validate_8_3_indicators_cross_check_TARGET_DOC}\n')
    print(f'[1/6] Загрузка правила проверки {kpsc_validate_8_3_indicators_cross_check_RULE_INDEX}')
    rule = kpsc_validate_8_3_indicators_cross_check_load_rule(kpsc_validate_8_3_indicators_cross_check_RULE_INDEX)
    if args.verbose:
        kpsc_validate_8_3_indicators_cross_check_log_step(log_file, 1, 'Загрузка правила проверки', f'Rule loaded:\n{json.dumps(rule, ensure_ascii=False, indent=2)}')
    print(f'[2/6] Загрузка данных из {args.parser_outputs}')
    data = kpsc_validate_8_3_indicators_cross_check_load_data(args.parser_outputs)
    if args.verbose:
        kpsc_validate_8_3_indicators_cross_check_log_step(log_file, 2, 'Загрузка данных', f'Source files loaded')
    print('[3/6] Извлечение релевантных данных')
    extracted = kpsc_validate_8_3_indicators_cross_check_extract_relevant_data(data)
    if args.verbose:
        kpsc_validate_8_3_indicators_cross_check_log_step(log_file, 3, 'Извлечение релевантных данных', f'Extracted data:\n{json.dumps(extracted, ensure_ascii=False, indent=2)}')
    print('[4/6] Формирование промпта с правилом из ТЗ')
    prompt = kpsc_validate_8_3_indicators_cross_check_build_prompt(extracted, rule)
    if args.verbose:
        kpsc_validate_8_3_indicators_cross_check_log_step(log_file, 4, 'Сформированный промпт для LLM', f'FULL PROMPT:\n{prompt}')
    print('[5/6] Вызов LLM')
    result = kpsc_validate_8_3_indicators_cross_check_call_llm(prompt, api_key)
    if args.verbose:
        kpsc_validate_8_3_indicators_cross_check_log_step(log_file, 5, 'Ответ от LLM', f'LLM Response:\n{json.dumps(result, ensure_ascii=False, indent=2)}')
    print('[6/6] Сохранение результата')
    kpsc_validate_8_3_indicators_cross_check_save_result(result, args.output)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'\n{status_emoji} Проверка {kpsc_validate_8_3_indicators_cross_check_RULE_INDEX}: {result['status']}')
    if result['discrepancy']:
        print(f'   Проблема: {result['discrepancy']}')
    if args.verbose:
        kpsc_validate_8_3_indicators_cross_check_log_step(log_file, 6, 'Финальный результат', f'Status: {result['status']}\nDiscrepancy: {result.get('discrepancy', 'нет')}\nResult saved to: {args.output}')
        print(f'\n📄 Лог сохранен: {log_file}')

kpsc_validate_8_3_indicators_cross_check_module = SimpleNamespace(RULE_INDEX=kpsc_validate_8_3_indicators_cross_check_RULE_INDEX, RULE_TITLE=kpsc_validate_8_3_indicators_cross_check_RULE_TITLE, TARGET_DOC=kpsc_validate_8_3_indicators_cross_check_TARGET_DOC, load_rule=kpsc_validate_8_3_indicators_cross_check_load_rule, load_data=kpsc_validate_8_3_indicators_cross_check_load_data, extract_relevant_data=kpsc_validate_8_3_indicators_cross_check_extract_relevant_data, build_prompt=kpsc_validate_8_3_indicators_cross_check_build_prompt, call_llm=kpsc_validate_8_3_indicators_cross_check_call_llm, save_result=kpsc_validate_8_3_indicators_cross_check_save_result, log_step=kpsc_validate_8_3_indicators_cross_check_log_step, main=kpsc_validate_8_3_indicators_cross_check_main)

# END_SOURCE_KPSC_VALIDATE_8_3_INDICATORS_CROSS_CHECK

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_10_UNITS_METODIKA
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_10_units_metodika.py.
kartochka_proekta_validate_10_units_metodika_RULE_INDEX = '10'

kartochka_proekta_validate_10_units_metodika_RULE_TITLE = 'Единицы измерения (методика расчета)'

def kartochka_proekta_validate_10_units_metodika_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_10_units_metodika_load_data(parser_outputs_dir: Path) -> tuple:
    """Загрузка kartochka_main.json и metodika.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        kartochka = json.load(f)
    with open(parser_outputs_dir / 'metodika.json', 'r', encoding='utf-8') as f:
        metodika = json.load(f)
    return (kartochka, metodika)

def kartochka_proekta_validate_10_units_metodika__normalize_name(name: str) -> str:
    """Нормализация названия показателя для сопоставления."""
    import re
    name = re.sub('\\(.*?\\)', '', name)
    name = name.replace(':', '').strip().lower()
    return name

def kartochka_proekta_validate_10_units_metodika_validate(kartochka: dict, metodika: dict) -> dict:
    """Проверка совпадения единиц измерения между методикой и карточкой."""
    k_indicators = kartochka.get('indicators', [])
    m_indicators = metodika.get('indicators', [])
    errors = []
    k_units = {}
    for ind in k_indicators:
        name = ind.get('name', '')
        if name:
            k_units[kartochka_proekta_validate_10_units_metodika__normalize_name(name)] = ind.get('unit', '')
    for m_ind in m_indicators:
        m_name = m_ind.get('name', '')
        m_unit = m_ind.get('unit')
        if not m_unit:
            m_norm = kartochka_proekta_validate_10_units_metodika__normalize_name(m_name)
            if m_norm in k_units and k_units[m_norm]:
                errors.append(f"Показатель '{m_name}': единица в методике не заполнена, в карточке — '{k_units[m_norm]}'")
            continue
        m_norm = kartochka_proekta_validate_10_units_metodika__normalize_name(m_name)
        k_unit = k_units.get(m_norm)
        if k_unit is None:
            continue
        if m_unit.lower().strip() != k_unit.lower().strip():
            errors.append(f"Показатель '{m_name}': единица в методике '{m_unit}' ≠ единица в карточке '{k_unit}'")
    if errors:
        return {'rule_index': kartochka_proekta_validate_10_units_metodika_RULE_INDEX, 'rule_title': kartochka_proekta_validate_10_units_metodika_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_10_units_metodika_RULE_INDEX, 'rule_title': kartochka_proekta_validate_10_units_metodika_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_10_units_metodika_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_10_units_metodika_RULE_INDEX}: {kartochka_proekta_validate_10_units_metodika_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_10_units_metodika_load_rule(kartochka_proekta_validate_10_units_metodika_RULE_INDEX)
    kartochka, metodika = kartochka_proekta_validate_10_units_metodika_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_10_units_metodika_validate(kartochka, metodika)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_10_units_metodika_RULE_INDEX}] {kartochka_proekta_validate_10_units_metodika_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_10_units_metodika_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_10_units_metodika_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_10_units_metodika_RULE_TITLE, load_rule=kartochka_proekta_validate_10_units_metodika_load_rule, load_data=kartochka_proekta_validate_10_units_metodika_load_data, _normalize_name=kartochka_proekta_validate_10_units_metodika__normalize_name, validate=kartochka_proekta_validate_10_units_metodika_validate, main=kartochka_proekta_validate_10_units_metodika_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_10_UNITS_METODIKA

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_11_CALC_METHOD
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_11_calc_method.py.
kartochka_proekta_validate_11_calc_method_RULE_INDEX = '11'

kartochka_proekta_validate_11_calc_method_RULE_TITLE = 'Способ расчёта и источник данных'

def kartochka_proekta_validate_11_calc_method_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_11_calc_method_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка metodika.json."""
    with open(parser_outputs_dir / 'metodika.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_11_calc_method_validate(data: dict) -> dict:
    """Проверка способа расчёта и источника данных."""
    indicators = data.get('indicators', [])
    errors = []
    for ind in indicators:
        name = ind.get('name', '')
        unit = ind.get('unit')
        calc_method = ind.get('calc_method')
        data_source = ind.get('data_source')
        if not unit:
            continue
        missing = []
        if not calc_method or not str(calc_method).strip():
            missing.append('способ расчёта')
        if not data_source or not str(data_source).strip():
            missing.append('источник данных')
        if missing:
            errors.append(f"Показатель '{name}': не заполнено — {', '.join(missing)}")
    if errors:
        return {'rule_index': kartochka_proekta_validate_11_calc_method_RULE_INDEX, 'rule_title': kartochka_proekta_validate_11_calc_method_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_11_calc_method_RULE_INDEX, 'rule_title': kartochka_proekta_validate_11_calc_method_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_11_calc_method_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_11_calc_method_RULE_INDEX}: {kartochka_proekta_validate_11_calc_method_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_11_calc_method_load_rule(kartochka_proekta_validate_11_calc_method_RULE_INDEX)
    data = kartochka_proekta_validate_11_calc_method_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_11_calc_method_validate(data)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_11_calc_method_RULE_INDEX}] {kartochka_proekta_validate_11_calc_method_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_11_calc_method_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_11_calc_method_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_11_calc_method_RULE_TITLE, load_rule=kartochka_proekta_validate_11_calc_method_load_rule, load_data=kartochka_proekta_validate_11_calc_method_load_data, validate=kartochka_proekta_validate_11_calc_method_validate, main=kartochka_proekta_validate_11_calc_method_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_11_CALC_METHOD

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_1_FILENAME
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_1_filename.py.
kartochka_proekta_validate_1_filename_RULE_INDEX = '1'

kartochka_proekta_validate_1_filename_RULE_TITLE = 'Название файла'

def kartochka_proekta_validate_1_filename_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_1_filename_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_1_filename_validate(data: dict, rule: dict) -> dict:
    """Проверка названия файла.

    Проверяет наличие ключевых слов «Карточка» и «проект» в имени файла.
    Наименование предприятия НЕ проверяется — оно динамическое.
    """
    filename = data.get('meta', {}).get('workbook', '')
    filename_lower = filename.lower()
    keywords = ['карточка', 'проект']
    missing = [kw for kw in keywords if kw not in filename_lower]
    if missing:
        return {'rule_index': kartochka_proekta_validate_1_filename_RULE_INDEX, 'rule_title': kartochka_proekta_validate_1_filename_RULE_TITLE, 'status': 'FAIL', 'discrepancy': f"Имя файла '{filename}' не содержит ключевые слова: {missing}"}
    return {'rule_index': kartochka_proekta_validate_1_filename_RULE_INDEX, 'rule_title': kartochka_proekta_validate_1_filename_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_1_filename_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_1_filename_RULE_INDEX}: {kartochka_proekta_validate_1_filename_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    rule = kartochka_proekta_validate_1_filename_load_rule(kartochka_proekta_validate_1_filename_RULE_INDEX)
    data = kartochka_proekta_validate_1_filename_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_1_filename_validate(data, rule)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_1_filename_RULE_INDEX}] {kartochka_proekta_validate_1_filename_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_1_filename_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_1_filename_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_1_filename_RULE_TITLE, load_rule=kartochka_proekta_validate_1_filename_load_rule, load_data=kartochka_proekta_validate_1_filename_load_data, validate=kartochka_proekta_validate_1_filename_validate, main=kartochka_proekta_validate_1_filename_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_1_FILENAME

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_2_ORG_NAME
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_2_org_name.py.
kartochka_proekta_validate_2_org_name_RULE_INDEX = '2'

kartochka_proekta_validate_2_org_name_RULE_TITLE = 'Вид организации и название предприятия'

def kartochka_proekta_validate_2_org_name_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_2_org_name_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_2_org_name_validate(data: dict, rule: dict) -> dict:
    """Проверка юридической формы и названия предприятия."""
    org_name = data.get('header', {}).get('org_name', '')
    errors = []
    if not org_name:
        errors.append('Поле org_name (B2) пустое — нет названия организации')
    else:
        jur_form_match = re.search('\\b(ООО|ЗАО|АО|ПАО|ОАО|ИП)\\b', org_name)
        if not jur_form_match:
            errors.append(f"Юридическая форма (ООО/ЗАО/АО/ПАО/ОАО/ИП) не найдена в '{org_name}'")
        name_match = re.search('["\\«](.+?)["\\»]', org_name)
        if not name_match or not name_match.group(1).strip():
            after_form = re.sub('^(ООО|ЗАО|АО|ПАО|ОАО|ИП)\\s*', '', org_name).strip()
            after_form = after_form.strip('"«»\'" ')
            if not after_form:
                errors.append('Наименование предприятия пустое после юридической формы')
    if errors:
        return {'rule_index': kartochka_proekta_validate_2_org_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_2_org_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_2_org_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_2_org_name_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_2_org_name_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_2_org_name_RULE_INDEX}: {kartochka_proekta_validate_2_org_name_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    rule = kartochka_proekta_validate_2_org_name_load_rule(kartochka_proekta_validate_2_org_name_RULE_INDEX)
    data = kartochka_proekta_validate_2_org_name_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_2_org_name_validate(data, rule)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_2_org_name_RULE_INDEX}] {kartochka_proekta_validate_2_org_name_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_2_org_name_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_2_org_name_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_2_org_name_RULE_TITLE, load_rule=kartochka_proekta_validate_2_org_name_load_rule, load_data=kartochka_proekta_validate_2_org_name_load_data, validate=kartochka_proekta_validate_2_org_name_validate, main=kartochka_proekta_validate_2_org_name_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_2_ORG_NAME

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_3_FLOW_NAME
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_3_flow_name.py.
kartochka_proekta_validate_3_flow_name_RULE_INDEX = '3'

kartochka_proekta_validate_3_flow_name_RULE_TITLE = 'Название потока'

def kartochka_proekta_validate_3_flow_name_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_3_flow_name_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_3_flow_name_check_semantic(project_name: str, rule: dict, api_key: str) -> dict:
    """LLM-проверка осмысленности названия проекта/потока (через единый клиент)."""
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    prompt = f'Ты эксперт по проверке документов «Карточка проекта» в рамках бережливого производства.\n\nТРЕБОВАНИЕ:\n{rule['requirement_expert']}\n\nКРИТЕРИИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nНазвание проекта/потока: "{project_name}"\n\nЗАДАНИЕ:\nПроверь, является ли название проекта осмысленным текстом, описывающим реальный проект или поток.\nНе является осмысленным: placeholder ("Название проекта"), набор символов ("ааааа"), слишком общий текст ("тест").\nЯвляется осмысленным: конкретное описание проекта ("Оптимизация производства приборов учёта").\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{kartochka_proekta_validate_3_flow_name_RULE_INDEX}",\n  "rule_title": "{kartochka_proekta_validate_3_flow_name_RULE_TITLE}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n'
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kartochka_proekta_validate_3_flow_name_validate(data: dict, rule: dict, api_key: str) -> dict:
    """Проверка названия проекта/потока без сетевой зависимости."""
    project_name = data.get('header', {}).get('project_name', '')
    if not project_name:
        return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': 'Поле project_name (B4) пустое — нет названия проекта/потока'}
    project_name = project_name.strip()
    if len(project_name) <= 5:
        return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта слишком короткое ({len(project_name)} симв.): '{project_name}'"}
    normalized = ' '.join(project_name.lower().split())
    placeholder_values = {'название проекта', 'название потока', 'проект', 'поток', 'тест', 'test', 'aaaaa', 'aaaa', 'qwerty'}
    letters_only = ''.join((ch for ch in normalized if ch.isalpha()))
    if normalized in placeholder_values:
        return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта выглядит как placeholder: '{project_name}'"}
    if not letters_only:
        return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта не содержит осмысленного текста: '{project_name}'"}
    unique_letters = set(letters_only)
    if len(unique_letters) <= 2:
        return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'FAIL', 'discrepancy': f"Название проекта похоже на набор повторяющихся символов: '{project_name}'"}
    return {'rule_index': kartochka_proekta_validate_3_flow_name_RULE_INDEX, 'rule_title': kartochka_proekta_validate_3_flow_name_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_3_flow_name_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_3_flow_name_RULE_INDEX}: {kartochka_proekta_validate_3_flow_name_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY')
    if not api_key:
        raise ValueError('OPENAI_API_KEY не найден в переменных окружения')
    rule = kartochka_proekta_validate_3_flow_name_load_rule(kartochka_proekta_validate_3_flow_name_RULE_INDEX)
    data = kartochka_proekta_validate_3_flow_name_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_3_flow_name_validate(data, rule, api_key)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_3_flow_name_RULE_INDEX}] {kartochka_proekta_validate_3_flow_name_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_3_flow_name_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_3_flow_name_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_3_flow_name_RULE_TITLE, load_rule=kartochka_proekta_validate_3_flow_name_load_rule, load_data=kartochka_proekta_validate_3_flow_name_load_data, check_semantic=kartochka_proekta_validate_3_flow_name_check_semantic, validate=kartochka_proekta_validate_3_flow_name_validate, main=kartochka_proekta_validate_3_flow_name_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_3_FLOW_NAME

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_4_SIGNEE
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_4_signee.py.
kartochka_proekta_validate_4_signee_RULE_INDEX = '4'

kartochka_proekta_validate_4_signee_RULE_TITLE = 'Должность, ФИО подписанта, дата и подпись'

def kartochka_proekta_validate_4_signee_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_4_signee_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_4_signee_validate(data: dict) -> dict:
    """Проверка подписанта."""
    header = data.get('header', {})
    errors = []
    position = header.get('signee_position', '')
    if not position:
        errors.append('Должность подписанта (K4) не заполнена')
    name = header.get('signee_name', '')
    if not name:
        errors.append('ФИО подписанта (K7) не заполнено')
    else:
        words = re.findall('[А-Яа-яЁёA-Za-z]+\\.?', name)
        if len(words) < 2:
            errors.append(f"ФИО подписанта содержит менее 2 слов: '{name}'")
    date_val = header.get('signee_date', '')
    if not date_val:
        errors.append('Дата подписания (K8) не заполнена')
    elif '___' in date_val or '20__' in date_val:
        errors.append(f"Дата подписания содержит незаполненные placeholder: '{date_val}'")
    if errors:
        return {'rule_index': kartochka_proekta_validate_4_signee_RULE_INDEX, 'rule_title': kartochka_proekta_validate_4_signee_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_4_signee_RULE_INDEX, 'rule_title': kartochka_proekta_validate_4_signee_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_4_signee_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_4_signee_RULE_INDEX}: {kartochka_proekta_validate_4_signee_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_4_signee_load_rule(kartochka_proekta_validate_4_signee_RULE_INDEX)
    data = kartochka_proekta_validate_4_signee_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_4_signee_validate(data)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_4_signee_RULE_INDEX}] {kartochka_proekta_validate_4_signee_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_4_signee_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_4_signee_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_4_signee_RULE_TITLE, load_rule=kartochka_proekta_validate_4_signee_load_rule, load_data=kartochka_proekta_validate_4_signee_load_data, validate=kartochka_proekta_validate_4_signee_validate, main=kartochka_proekta_validate_4_signee_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_4_SIGNEE

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_5_REQUIRED_FIELDS
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_5_required_fields.py.
kartochka_proekta_validate_5_required_fields_RULE_INDEX = '5'

kartochka_proekta_validate_5_required_fields_RULE_TITLE = 'Обязательные поля секции 1'

kartochka_proekta_validate_5_required_fields_REQUIRED_FIELDS = {'clients': 'Клиенты процесса', 'perimeter': 'Периметр проекта', 'owner': 'Владелец процесса', 'boundaries': 'Границы процесса', 'leader': 'Руководитель проекта', 'team': 'Команда проекта'}

kartochka_proekta_validate_5_required_fields_FIO_FIELDS = ['owner', 'leader', 'team']

def kartochka_proekta_validate_5_required_fields_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_5_required_fields_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_5_required_fields__check_fio_format(text: str) -> bool:
    """Проверяет наличие паттерна 'ФИО - должность' (разделитель — тире)."""
    return bool(re.search('.+\\s*[-–—]\\s*.+', text))

def kartochka_proekta_validate_5_required_fields_validate(data: dict) -> dict:
    """Проверка обязательных полей секции 1."""
    section1 = data.get('section1', {})
    errors = []
    empty_fields = []
    for field_key, field_name in kartochka_proekta_validate_5_required_fields_REQUIRED_FIELDS.items():
        val = section1.get(field_key, '')
        if not val or not val.strip():
            empty_fields.append(field_name)
    if empty_fields:
        errors.append(f'Пустые поля: {', '.join(empty_fields)}')
    bad_format = []
    for field_key in kartochka_proekta_validate_5_required_fields_FIO_FIELDS:
        val = section1.get(field_key, '')
        if not val:
            continue
        if field_key == 'team':
            members = [m.strip() for m in val.split(',') if m.strip()]
            for member in members:
                if not kartochka_proekta_validate_5_required_fields__check_fio_format(member):
                    bad_format.append(f"team: '{member[:50]}'")
                    break
        elif not kartochka_proekta_validate_5_required_fields__check_fio_format(val):
            bad_format.append(f"{kartochka_proekta_validate_5_required_fields_REQUIRED_FIELDS[field_key]}: '{val[:50]}'")
    if bad_format:
        errors.append(f'Неверный формат ФИО-должность: {'; '.join(bad_format)}')
    if errors:
        return {'rule_index': kartochka_proekta_validate_5_required_fields_RULE_INDEX, 'rule_title': kartochka_proekta_validate_5_required_fields_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_5_required_fields_RULE_INDEX, 'rule_title': kartochka_proekta_validate_5_required_fields_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_5_required_fields_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_5_required_fields_RULE_INDEX}: {kartochka_proekta_validate_5_required_fields_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_5_required_fields_load_rule(kartochka_proekta_validate_5_required_fields_RULE_INDEX)
    data = kartochka_proekta_validate_5_required_fields_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_5_required_fields_validate(data)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_5_required_fields_RULE_INDEX}] {kartochka_proekta_validate_5_required_fields_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_5_required_fields_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_5_required_fields_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_5_required_fields_RULE_TITLE, REQUIRED_FIELDS=kartochka_proekta_validate_5_required_fields_REQUIRED_FIELDS, FIO_FIELDS=kartochka_proekta_validate_5_required_fields_FIO_FIELDS, load_rule=kartochka_proekta_validate_5_required_fields_load_rule, load_data=kartochka_proekta_validate_5_required_fields_load_data, _check_fio_format=kartochka_proekta_validate_5_required_fields__check_fio_format, validate=kartochka_proekta_validate_5_required_fields_validate, main=kartochka_proekta_validate_5_required_fields_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_5_REQUIRED_FIELDS

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_6_JUSTIFICATION
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_6_justification.py.
kartochka_proekta_validate_6_justification_RULE_INDEX = '6'

kartochka_proekta_validate_6_justification_RULE_TITLE = 'Обоснование и ключевой риск'

def kartochka_proekta_validate_6_justification_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_6_justification_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_6_justification_check_semantic(key_risk: str, justification: str, rule: dict, api_key: str) -> dict:
    """LLM-проверка осмысленности обоснования и ключевого риска (через единый клиент)."""
    del api_key  # src.llm.client сам формирует api_key для Spark-vLLM
    prompt = f'Ты эксперт по проверке документов «Карточка проекта» в рамках бережливого производства.\n\nТРЕБОВАНИЕ:\n{rule['requirement_expert']}\n\nКРИТЕРИИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nКлючевой риск (M11): "{key_risk}"\nОбоснование выбора потока (M13): "{justification}"\n\nЗАДАНИЕ:\nПроверь, содержат ли оба поля осмысленный текст:\n- Ключевой риск — должен описывать конкретный риск проекта (например: "Срыв сроков", "Потеря клиентов").\n  НЕ осмысленный: "-", "нет", "риск", набор символов.\n- Обоснование — должно содержать аргументацию выбора потока (например: "Наличие ожидания в потоке, несвоевременная подготовка").\n  НЕ осмысленный: "-", "обоснование", "тест", набор символов.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{kartochka_proekta_validate_6_justification_RULE_INDEX}",\n  "rule_title": "{kartochka_proekta_validate_6_justification_RULE_TITLE}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n'
    result_text = call_llm(
        messages=[
            {'role': 'system', 'content': 'Ты эксперт по валидации документов. Отвечаешь строго в формате JSON.'},
            {'role': 'user', 'content': prompt},
        ],
        model=LLM_CONFIG.default_model,
        base_url=LLM_CONFIG.base_url,
        temperature=0.0,
        response_format={'type': 'json_object'},
    )
    return json.loads(result_text)

def kartochka_proekta_validate_6_justification_validate(data: dict, rule: dict, api_key: str) -> dict:
    """Проверка обоснования и ключевого риска: non-LLM + LLM."""
    section2 = data.get('section2', {})
    errors = []
    key_risk = section2.get('key_risk', '')
    justification = section2.get('justification', '')
    if not key_risk or not key_risk.strip():
        errors.append('Ключевой риск (M11) не заполнен')
    if not justification or not justification.strip():
        errors.append('Обоснование выбора потока (M13) не заполнено')
    if errors:
        return {'rule_index': kartochka_proekta_validate_6_justification_RULE_INDEX, 'rule_title': kartochka_proekta_validate_6_justification_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return kartochka_proekta_validate_6_justification_check_semantic(key_risk, justification, rule, api_key)

def kartochka_proekta_validate_6_justification_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_6_justification_RULE_INDEX}: {kartochka_proekta_validate_6_justification_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    api_key = os.environ.get('OPENAI_API_KEY')
    if not api_key:
        raise ValueError('OPENAI_API_KEY не найден в переменных окружения')
    rule = kartochka_proekta_validate_6_justification_load_rule(kartochka_proekta_validate_6_justification_RULE_INDEX)
    data = kartochka_proekta_validate_6_justification_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_6_justification_validate(data, rule, api_key)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_6_justification_RULE_INDEX}] {kartochka_proekta_validate_6_justification_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_6_justification_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_6_justification_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_6_justification_RULE_TITLE, load_rule=kartochka_proekta_validate_6_justification_load_rule, load_data=kartochka_proekta_validate_6_justification_load_data, check_semantic=kartochka_proekta_validate_6_justification_check_semantic, validate=kartochka_proekta_validate_6_justification_validate, main=kartochka_proekta_validate_6_justification_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_6_JUSTIFICATION

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_7_EVENT_DATES
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_7_event_dates.py.
kartochka_proekta_validate_7_event_dates_RULE_INDEX = '7'

kartochka_proekta_validate_7_event_dates_RULE_TITLE = 'Даты мероприятий'

def kartochka_proekta_validate_7_event_dates_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_7_event_dates_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_7_event_dates__is_stage_event(name: str) -> bool:
    """Проверяет, является ли событие этапным (начинается с '1.', '2.', '3.', '4.')."""
    return bool(re.match('^\\d+\\.', name.strip()))

def kartochka_proekta_validate_7_event_dates_validate(data: dict) -> dict:
    """Проверка дат мероприятий."""
    events = data.get('events', [])
    errors = []
    if not events:
        return {'rule_index': kartochka_proekta_validate_7_event_dates_RULE_INDEX, 'rule_title': kartochka_proekta_validate_7_event_dates_RULE_TITLE, 'status': 'FAIL', 'discrepancy': 'Список мероприятий (events) пуст'}
    for i, event in enumerate(events):
        name = event.get('name', f'Событие #{i + 1}')
        start_date = event.get('start_date')
        end_date = event.get('end_date')
        if not start_date:
            errors.append(f"'{name[:40]}' — нет даты начала")
        if kartochka_proekta_validate_7_event_dates__is_stage_event(name) and (not end_date):
            errors.append(f"'{name[:40]}' — этапное событие без даты окончания")
    if errors:
        return {'rule_index': kartochka_proekta_validate_7_event_dates_RULE_INDEX, 'rule_title': kartochka_proekta_validate_7_event_dates_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_7_event_dates_RULE_INDEX, 'rule_title': kartochka_proekta_validate_7_event_dates_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_7_event_dates_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_7_event_dates_RULE_INDEX}: {kartochka_proekta_validate_7_event_dates_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_7_event_dates_load_rule(kartochka_proekta_validate_7_event_dates_RULE_INDEX)
    data = kartochka_proekta_validate_7_event_dates_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_7_event_dates_validate(data)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_7_event_dates_RULE_INDEX}] {kartochka_proekta_validate_7_event_dates_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_7_event_dates_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_7_event_dates_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_7_event_dates_RULE_TITLE, load_rule=kartochka_proekta_validate_7_event_dates_load_rule, load_data=kartochka_proekta_validate_7_event_dates_load_data, _is_stage_event=kartochka_proekta_validate_7_event_dates__is_stage_event, validate=kartochka_proekta_validate_7_event_dates_validate, main=kartochka_proekta_validate_7_event_dates_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_7_EVENT_DATES

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_8_INDICATORS
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_8_indicators.py.
kartochka_proekta_validate_8_indicators_RULE_INDEX = '8'

kartochka_proekta_validate_8_indicators_RULE_TITLE = 'Показатели и даты'

def kartochka_proekta_validate_8_indicators_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_8_indicators_load_data(parser_outputs_dir: Path) -> dict:
    """Загрузка kartochka_main.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def kartochka_proekta_validate_8_indicators_validate(data: dict) -> dict:
    """Проверка показателей и дат."""
    indicators = data.get('indicators', [])
    indicator_dates = data.get('indicator_dates', {})
    errors = []
    if not indicators:
        errors.append('Нет показателей (массив indicators пуст)')
    else:
        required_fields = ['name', 'unit', 'base_value', 'target_value', 'ideal_value']
        for ind in indicators:
            num = ind.get('number', '?')
            missing = []
            for field in required_fields:
                val = ind.get(field)
                if val is None or (isinstance(val, str) and (not val.strip())):
                    missing.append(field)
            if missing:
                errors.append(f'Показатель #{num}: пустые поля [{', '.join(missing)}]')
    date_fields = {'base_date': 'Дата базы', 'target_date': 'Дата цели', 'ideal_date': 'Дата идеала'}
    for field, label in date_fields.items():
        val = indicator_dates.get(field)
        if not val:
            errors.append(f'{label} не заполнена')
    if errors:
        return {'rule_index': kartochka_proekta_validate_8_indicators_RULE_INDEX, 'rule_title': kartochka_proekta_validate_8_indicators_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_8_indicators_RULE_INDEX, 'rule_title': kartochka_proekta_validate_8_indicators_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_8_indicators_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_8_indicators_RULE_INDEX}: {kartochka_proekta_validate_8_indicators_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_8_indicators_load_rule(kartochka_proekta_validate_8_indicators_RULE_INDEX)
    data = kartochka_proekta_validate_8_indicators_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_8_indicators_validate(data)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_8_indicators_RULE_INDEX}] {kartochka_proekta_validate_8_indicators_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_8_indicators_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_8_indicators_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_8_indicators_RULE_TITLE, load_rule=kartochka_proekta_validate_8_indicators_load_rule, load_data=kartochka_proekta_validate_8_indicators_load_data, validate=kartochka_proekta_validate_8_indicators_validate, main=kartochka_proekta_validate_8_indicators_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_8_INDICATORS

# START_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_9_UNITS_KARTOCHKA
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/validation_scripts/validate_9_units_kartochka.py.
kartochka_proekta_validate_9_units_kartochka_RULE_INDEX = '9'

kartochka_proekta_validate_9_units_kartochka_RULE_TITLE = 'Единицы измерения (карточка)'

kartochka_proekta_validate_9_units_kartochka_CATEGORY_KEYWORDS = {'Время протекания процесса': ['время', 'протекан'], 'Выработка': ['выработк'], 'Незавершенное производство': ['запас', 'нзп', 'незавершен']}

kartochka_proekta_validate_9_units_kartochka_DEFAULT_CATEGORY = 'Дополнительный показатель'

def kartochka_proekta_validate_9_units_kartochka_load_rule(rule_index: str) -> dict:
    """Загрузка правила из validation_rules.json."""
    rules_file = Path(os.environ.get('VALIDATION_RULES_PATH', str(Path(__file__).resolve().parents[2] / 'doc_configs' / 'kartochka_proekta' / 'validation_rules.json')))
    with open(rules_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    for rule in data['rules']:
        if rule['rule_index'] == rule_index:
            return rule
    raise ValueError(f'Правило {rule_index} не найдено')

def kartochka_proekta_validate_9_units_kartochka_load_data(parser_outputs_dir: Path) -> tuple:
    """Загрузка kartochka_main.json и dropdown_units.json."""
    with open(parser_outputs_dir / 'kartochka_main.json', 'r', encoding='utf-8') as f:
        kartochka = json.load(f)
    with open(parser_outputs_dir / 'dropdown_units.json', 'r', encoding='utf-8') as f:
        dropdown = json.load(f)
    return (kartochka, dropdown)

def kartochka_proekta_validate_9_units_kartochka__detect_category(indicator_name: str) -> str:
    """Определяет категорию показателя по ключевым словам в названии."""
    name_lower = indicator_name.lower()
    for category, keywords in kartochka_proekta_validate_9_units_kartochka_CATEGORY_KEYWORDS.items():
        if any((kw in name_lower for kw in keywords)):
            return category
    return kartochka_proekta_validate_9_units_kartochka_DEFAULT_CATEGORY

def kartochka_proekta_validate_9_units_kartochka_validate(kartochka: dict, dropdown: dict) -> dict:
    """Проверка единиц измерения показателей по справочнику."""
    indicators = kartochka.get('indicators', [])
    categories = dropdown.get('categories', {})
    errors = []
    all_units = set()
    all_units_display = []
    for units_list in categories.values():
        for u in units_list:
            normalized = u.lower().strip().rstrip('.')
            all_units.add(normalized)
            all_units_display.append(u)
    for ind in indicators:
        name = ind.get('name', '')
        unit = ind.get('unit', '')
        num = ind.get('number', '?')
        if not unit:
            errors.append(f"Показатель #{num} '{name}': единица измерения не заполнена")
            continue
        unit_normalized = unit.lower().strip().rstrip('.')
        if unit_normalized not in all_units:
            errors.append(f"Показатель #{num} '{name}': единица '{unit}' не найдена в справочнике (допустимые: {', '.join(all_units_display[:10])}...)")
    if errors:
        return {'rule_index': kartochka_proekta_validate_9_units_kartochka_RULE_INDEX, 'rule_title': kartochka_proekta_validate_9_units_kartochka_RULE_TITLE, 'status': 'FAIL', 'discrepancy': '; '.join(errors)}
    return {'rule_index': kartochka_proekta_validate_9_units_kartochka_RULE_INDEX, 'rule_title': kartochka_proekta_validate_9_units_kartochka_RULE_TITLE, 'status': 'PASS', 'discrepancy': ''}

def kartochka_proekta_validate_9_units_kartochka_main():
    parser = argparse.ArgumentParser(description=f'Валидатор {kartochka_proekta_validate_9_units_kartochka_RULE_INDEX}: {kartochka_proekta_validate_9_units_kartochka_RULE_TITLE}')
    parser.add_argument('--parser-outputs', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    kartochka_proekta_validate_9_units_kartochka_load_rule(kartochka_proekta_validate_9_units_kartochka_RULE_INDEX)
    kartochka, dropdown = kartochka_proekta_validate_9_units_kartochka_load_data(args.parser_outputs)
    result = kartochka_proekta_validate_9_units_kartochka_validate(kartochka, dropdown)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    status_emoji = '✅' if result['status'] == 'PASS' else '❌'
    print(f'{status_emoji} [{kartochka_proekta_validate_9_units_kartochka_RULE_INDEX}] {kartochka_proekta_validate_9_units_kartochka_RULE_TITLE}: {result['status']}')

kartochka_proekta_validate_9_units_kartochka_module = SimpleNamespace(RULE_INDEX=kartochka_proekta_validate_9_units_kartochka_RULE_INDEX, RULE_TITLE=kartochka_proekta_validate_9_units_kartochka_RULE_TITLE, CATEGORY_KEYWORDS=kartochka_proekta_validate_9_units_kartochka_CATEGORY_KEYWORDS, DEFAULT_CATEGORY=kartochka_proekta_validate_9_units_kartochka_DEFAULT_CATEGORY, load_rule=kartochka_proekta_validate_9_units_kartochka_load_rule, load_data=kartochka_proekta_validate_9_units_kartochka_load_data, _detect_category=kartochka_proekta_validate_9_units_kartochka__detect_category, validate=kartochka_proekta_validate_9_units_kartochka_validate, main=kartochka_proekta_validate_9_units_kartochka_main)

# END_SOURCE_KARTOCHKA_PROEKTA_VALIDATE_9_UNITS_KARTOCHKA
# END_VALIDATORS

# START_NON_LLM_CHECKS
# PURPOSE: Register and run deterministic rules that do not need an LLM call.
# INPUTS: parsed documents, audit config, rule identifiers.
# OUTPUTS: Violation lists for exact or heuristic checks.
# KEYWORDS: non-llm, registry, filename, shapka, prikaz.
# LINKS: audit_engine/non_llm_checks/*.py.
# RATIONALE: Deterministic rules stay explicit and inspectable in the monolith.

# START_SOURCE_NON_LLM_CHECKS_REGISTRY
# PURPOSE: Inlined source from audit_engine/non_llm_checks/registry.py.
non_llm_checks_registry__REGISTRY: Dict[Tuple[str, int], Callable] = {}

def register(doc_type: str, rule_index: int):
    """
    Декоратор для регистрации non-LLM проверки.

    Args:
        doc_type: тип документа (prikaz_ic, cheklist_eu, presentation_eu, ...)
        rule_index: номер правила

    Функция проверки должна принимать (target_doc: Dict, config: AuditConfig) -> List[Dict]
    """

    def decorator(fn: Callable):
        non_llm_checks_registry__REGISTRY[doc_type, rule_index] = fn
        return fn
    return decorator

def get_check(doc_type: str, rule_index: int) -> Optional[Callable]:
    """Возвращает зарегистрированную проверку или None."""
    return non_llm_checks_registry__REGISTRY.get((doc_type, rule_index))

def get_all_checks(doc_type: str) -> Dict[int, Callable]:
    """Возвращает все проверки для типа документа {rule_index: check_fn}."""
    return {idx: fn for (dt, idx), fn in non_llm_checks_registry__REGISTRY.items() if dt == doc_type}

non_llm_checks_registry_module = SimpleNamespace(_REGISTRY=non_llm_checks_registry__REGISTRY, register=register, get_check=get_check, get_all_checks=get_all_checks)

# END_SOURCE_NON_LLM_CHECKS_REGISTRY

# START_SOURCE_NON_LLM_CHECKS_CHEKLIST_SCORES
# PURPOSE: Inlined source from audit_engine/non_llm_checks/cheklist_scores.py.
def _extract_rows(html: str) -> List[List[str]]:
    """
    Извлекает строки таблицы из HTML.

    Args:
        html: HTML-текст с <tr>/<td> тегами
    Returns:
        Список строк, каждая — список значений ячеек (stripped).
    """
    rows = []
    for tr_match in re.finditer('<tr[^>]*>(.*?)</tr>', html, re.DOTALL):
        tr_content = tr_match.group(1)
        cells = [cell.strip() for cell in re.findall('<td[^>]*>(.*?)</td>', tr_content, re.DOTALL)]
        rows.append(cells)
    return rows

def _parse_criteria(table_text: str) -> Tuple[Dict[int, Optional[int]], Optional[int]]:
    """
    Извлекает оценки 7 критериев и итоговую оценку из HTML-таблицы.

    Структура строки критерия (6 ячеек):
      [номер(1-7), текст, оценка(0/1/2/пусто), описание_0, описание_1, описание_2]
    Строка итоговой оценки:
      [пусто, "Итоговая оценка", значение, ...]

    Args:
        table_text: HTML чанка таблица_критериев
    Returns:
        (criteria, total):
          criteria: {1: 2, 2: None, ...} — номер критерия → оценка (None = пустая)
          total: итоговая оценка (int) или None если не найдена
    """
    criteria = {}
    total = None
    for row in _extract_rows(table_text):
        if len(row) < 3:
            continue
        first_cell = row[0].strip()
        second_cell = row[1].strip()
        third_cell = row[2].strip()
        if first_cell in ('1', '2', '3', '4', '5', '6', '7'):
            num = int(first_cell)
            if third_cell in ('0', '1', '2'):
                criteria[num] = int(third_cell)
            else:
                criteria[num] = None
        if 'итоговая' in second_cell.lower() and 'оценка' in second_cell.lower():
            for cell in row[2:]:
                cleaned = re.sub('[^0-9]', '', cell)
                if cleaned and 0 <= int(cleaned) <= 14:
                    total = int(cleaned)
                    break
    return (criteria, total)

@register('cheklist_eu', 5)
def check_criteria_scores(target_doc: Dict[str, Any], config: Any) -> List[Dict[str, Any]]:
    """
    Правило #5: проверка заполнения оценок критериев.

    Проверяет что все 7 критериев имеют оценку (0, 1 или 2).
    Парсит HTML-таблицу из чанка 'таблица_критериев'.

    Args:
        target_doc: распарсенный документ {чанк: текст}
        config: конфиг аудита
    Returns:
        список нарушений (пустой = всё ок)
    """
    table_text = target_doc.get('таблица_критериев', '')
    if not table_text:
        return [{'rule_index': 5, 'rule_title': 'Проверка заполнения таблицы критериев', 'Целевой документ': 'отсутствует', 'Различие': "Не найден чанк 'таблица_критериев'"}]
    criteria, _ = _parse_criteria(table_text)
    violations = []
    for i in range(1, 8):
        if i not in criteria:
            violations.append({'rule_index': 5, 'rule_title': 'Проверка заполнения таблицы критериев', 'Целевой документ': f'Критерий {i} не найден в таблице', 'Различие': f'Критерий {i} должен присутствовать в таблице с оценкой 0, 1 или 2'})
        elif criteria[i] is None:
            violations.append({'rule_index': 5, 'rule_title': 'Проверка заполнения таблицы критериев', 'Целевой документ': f'Критерий {i}, оценка пустая', 'Различие': f'Для критерия {i} должна быть заполнена оценка: 0, 1 или 2'})
    return violations

@register('cheklist_eu', 6)
def check_total_score(target_doc: Dict[str, Any], config: Any) -> List[Dict[str, Any]]:
    """
    Правило #6: проверка итоговой оценки.

    Проверяет что указанная итоговая оценка равна сумме баллов по 7 критериям.
    Если оценки не заполнены — пропускаем (Rule 5 уже поймала).

    Args:
        target_doc: распарсенный документ {чанк: текст}
        config: конфиг аудита
    Returns:
        список нарушений (пустой = всё ок)
    """
    table_text = target_doc.get('таблица_критериев', '')
    if not table_text:
        return [{'rule_index': 6, 'rule_title': 'Проверка итоговой оценки', 'Целевой документ': 'отсутствует', 'Различие': "Не найден чанк 'таблица_критериев'"}]
    criteria, stated_total = _parse_criteria(table_text)
    scores = {k: v for k, v in criteria.items() if v is not None}
    if len(scores) != 7:
        return []
    calculated_sum = sum(scores.values())
    if stated_total is None:
        return [{'rule_index': 6, 'rule_title': 'Проверка итоговой оценки', 'Целевой документ': 'Строка «Итоговая оценка» не найдена или пустая', 'Различие': f'Должна быть указана итоговая оценка. Рассчитанная сумма: {calculated_sum}'}]
    if stated_total != calculated_sum:
        return [{'rule_index': 6, 'rule_title': 'Проверка итоговой оценки', 'Целевой документ': f'Итоговая оценка: {stated_total}', 'Различие': f'Указанная сумма ({stated_total}) не совпадает с рассчитанной ({calculated_sum})'}]
    return []

non_llm_checks_cheklist_scores_module = SimpleNamespace(_extract_rows=_extract_rows, _parse_criteria=_parse_criteria, check_criteria_scores=check_criteria_scores, check_total_score=check_total_score)

# END_SOURCE_NON_LLM_CHECKS_CHEKLIST_SCORES

# START_SOURCE_NON_LLM_CHECKS_FILENAME
# PURPOSE: Inlined source from audit_engine/non_llm_checks/filename.py.
def check_filename_universal(target_doc: Dict[str, Any], config: Any, rule_index: int=1, rule_title: str='Проверка названия файла') -> List[Dict[str, Any]]:
    """
    Универсальная проверка имени файла.

    Два режима:
    1. filename_keywords (приоритет) — все ключевые слова должны быть в имени файла
    2. filename_pattern (обратная совместимость) — подстрока должна быть в имени
    """
    actual = target_doc.get('filename', '')
    actual_clean = re.sub('\\.(docx?|pptx?|pdf)$', '', actual, flags=re.IGNORECASE)
    name_lower = unicodedata.normalize('NFC', actual_clean).lower()
    keywords = getattr(config, 'filename_keywords', None)
    if keywords:
        if all((kw.lower() in name_lower for kw in keywords)):
            return []
        missing = [kw for kw in keywords if kw.lower() not in name_lower]
        return [{'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': actual_clean, 'Различие': f'Ожидалось имя файла, содержащее все ключевые слова: {keywords}. Отсутствуют: {missing}'}]
    expected_pattern = getattr(config, 'filename_pattern', '')
    if not expected_pattern:
        return []
    if expected_pattern.lower() in name_lower:
        return []
    return [{'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': actual_clean, 'Различие': f'Ожидалось имя файла, содержащее: «{expected_pattern}»'}]

@register('presentation_eu', 1)
def check_filename_presentation(target_doc, config):
    """Правило #1: проверка имени файла для Презентации ЭУ."""
    return check_filename_universal(target_doc, config, 1, 'Проверка названия файла')

@register('cheklist_eu', 1)
def check_filename_cheklist(target_doc, config):
    """Правило #1: проверка имени файла для Чек-листа ЭУ."""
    return check_filename_universal(target_doc, config, 1, 'Проверка названия файла')

@register('prikaz_ic', 1)
def check_filename_prikaz(target_doc, config):
    """Правило #1: проверка имени файла для Приказа о создании ИЦ."""
    return check_filename_universal(target_doc, config, 1, 'Проверка названия файла')

@register('polozhenie_po', 1)
def check_filename_polozhenie_po(target_doc, config):
    """Правило #1: проверка имени файла для Положения о ПО."""
    return check_filename_universal(target_doc, config, 1, 'Проверка названия файла')

@register('prikaz_formirovanie_po', 1)
def check_filename_prikaz_formirovanie_po(target_doc, config):
    """Правило #1: проверка имени файла для Приказа о формировании ПО."""
    return check_filename_universal(target_doc, config, 1, 'Проверка названия файла')

@register('prikaz_ic_el', 1)
def check_filename_prikaz_ic_el(target_doc, config):
    """Правило #1: проверка имени файла для Приказа о создании ИЦ (эл. вид)."""
    return check_filename_universal(target_doc, config, 1, 'Проверка имени файла документа.')

@register('prikaz_ic_potoka', 2)
def check_filename_prikaz_ic_potoka(target_doc, config):
    """Правило #2: проверка имени файла для Приказа о создании ИЦ потока."""
    return check_filename_universal(target_doc, config, 2, 'Проверка имени файла документа.')

@register('prikaz_ic_potoka_el', 2)
def check_filename_prikaz_ic_potoka_el(target_doc, config):
    """Правило #2: проверка имени файла для Приказа о создании ИЦ потока (эл. вид)."""
    return check_filename_universal(target_doc, config, 2, 'Проверка имени файла документа.')

@register('prikaz_vyhod', 2)
def check_filename_prikaz_vyhod(target_doc, config):
    """Правило #2: проверка имени файла для Приказа о проведении выхода."""
    return check_filename_universal(target_doc, config, 2, 'Проверка названия файла документа')

@register('prikaz_comp_ppu', 2)
def check_filename_prikaz_comp_ppu(target_doc, config):
    """Правило #2: проверка имени файла для Приказа о конкурсах ППУ."""
    return check_filename_universal(target_doc, config, 2, 'Проверка имени файла документа.')

@register('prikaz_ppu', 2)
def check_filename_prikaz_ppu(target_doc, config):
    """Правило #2: проверка имени файла для Приказа о ППУ."""
    return check_filename_universal(target_doc, config, 2, 'Проверка имени файла документа.')

@register('polozhenie_comp_ppu', 2)
def check_filename_polozhenie_comp_ppu(target_doc, config):
    """Правило #2: проверка имени файла для Положения о конкурсах ППУ."""
    return check_filename_universal(target_doc, config, 2, 'Проверка имени файла документа.')

@register('polozhenie_ppu', 2)
def check_filename_polozhenie_ppu(target_doc, config):
    """Правило #2: проверка имени файла для Положения о ППУ."""
    return check_filename_universal(target_doc, config, 2, 'Проверка имени файла документа.')

@register('akt_nachala', 1)
def check_filename_akt_nachala(target_doc, config):
    """Non-LLM: проверка имени файла для Акта начала мероприятий"""
    return check_filename_universal(target_doc, config, 1, 'Проверка названия файла')

@register('prikaz_pa', 1)
def check_filename_prikaz_pa(target_doc, config):
    """Правило #1: проверка имени файла для Приказа о внедрении ПА."""
    return check_filename_universal(target_doc, config, 1, 'Проверка имени файла документа.')

@register('prikaz_tirazh', 1)
def check_filename_prikaz_tirazh(target_doc, config):
    """Правило #1: проверка имени файла для Приказа о тиражировании."""
    return check_filename_universal(target_doc, config, 1, 'Проверка имени файла документа.')

@register('otchet_rezultatov', 1)
def check_filename_otchet_rezultatov(target_doc, config):
    """Правило #1: проверка имени файла для Отчёта о результатах опроса."""
    return check_filename_universal(target_doc, config, 1, 'Проверка названия файла')

@register('prikaz_otvetstvennyh', 1)
def check_filename_prikaz_otvetstvennyh(target_doc, config):
    """Правило #1: проверка имени файла для Приказа о назначении ответственных."""
    return check_filename_universal(target_doc, config, 1, 'Проверка названия файла')

non_llm_checks_filename_module = SimpleNamespace(check_filename_universal=check_filename_universal, check_filename_presentation=check_filename_presentation, check_filename_cheklist=check_filename_cheklist, check_filename_prikaz=check_filename_prikaz, check_filename_polozhenie_po=check_filename_polozhenie_po, check_filename_prikaz_formirovanie_po=check_filename_prikaz_formirovanie_po, check_filename_prikaz_ic_el=check_filename_prikaz_ic_el, check_filename_prikaz_ic_potoka=check_filename_prikaz_ic_potoka, check_filename_prikaz_ic_potoka_el=check_filename_prikaz_ic_potoka_el, check_filename_prikaz_vyhod=check_filename_prikaz_vyhod, check_filename_prikaz_comp_ppu=check_filename_prikaz_comp_ppu, check_filename_prikaz_ppu=check_filename_prikaz_ppu, check_filename_polozhenie_comp_ppu=check_filename_polozhenie_comp_ppu, check_filename_polozhenie_ppu=check_filename_polozhenie_ppu, check_filename_akt_nachala=check_filename_akt_nachala, check_filename_prikaz_pa=check_filename_prikaz_pa, check_filename_prikaz_tirazh=check_filename_prikaz_tirazh, check_filename_otchet_rezultatov=check_filename_otchet_rezultatov, check_filename_prikaz_otvetstvennyh=check_filename_prikaz_otvetstvennyh)

# END_SOURCE_NON_LLM_CHECKS_FILENAME

# START_SOURCE_NON_LLM_CHECKS_PRIKAZ_CHECKS
# PURPOSE: Inlined source from audit_engine/non_llm_checks/prikaz_checks.py.
def _get_scopes_text(target_doc: Dict[str, Any], scopes: List[str]) -> str:
    """Собирает текст из нескольких scope в одну строку."""
    parts = []
    for scope in scopes:
        val = target_doc.get(scope, '')
        if val:
            parts.append(str(val))
    return '\n'.join(parts)

def check_prikaz_and_number(target_doc: Dict[str, Any], config: Any, rule_index: int=4, rule_title: str='Проверка наличия слова ПРИКАЗ и номера приказа.', scopes: List[str]=None) -> List[Dict[str, Any]]:
    """
    Проверяет наличие слова ПРИКАЗ и номера приказа.

    Допустимые формы: ПРИКАЗ, Приказ, П Р И К А З
    Номер: символ № + любой непустой текст (кроме подчёркиваний)
    scopes: список scope для извлечения текста (по умолчанию заголовок_город + номер_дата)
    """
    if scopes is None:
        scopes = ['заголовок_город', 'номер_дата']
    text = _get_scopes_text(target_doc, scopes)
    violations = []
    has_prikaz = bool(re.search('П\\s*Р\\s*И\\s*К\\s*А\\s*З|Приказ', text, re.IGNORECASE))
    if not has_prikaz:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует слово «ПРИКАЗ»'})
    number_match = re.search('№\\s*(.+)', text)
    if number_match:
        after_sign = number_match.group(1).strip()
        cleaned = re.sub('[_\\s]', '', after_sign)
        is_bn = bool(re.match('^б[/\\\\]?н$', cleaned, re.IGNORECASE))
        if not cleaned or is_bn:
            violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': f'№ {after_sign}', 'Различие': 'Номер приказа не заполнен' + (' (б/н = без номера)' if is_bn else ' (пустой или подчёркивания)')})
    else:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует символ № с номером приказа'})
    return violations

def check_city_and_date(target_doc: Dict[str, Any], config: Any, rule_index: int=5, rule_title: str='Проверка наличия города и даты приказа.', scopes: List[str]=None) -> List[Dict[str, Any]]:
    """
    Проверяет наличие города и полной даты приказа.

    Город: «г. Москва», полный адрес с городом, «Москва»
    Дата: числовой (ДД.ММ.ГГГГ) или словесный (01 сентября 2025 г.)
    scopes: список scope для извлечения текста (по умолчанию заголовок_город + номер_дата)
    """
    if scopes is None:
        scopes = ['заголовок_город', 'номер_дата']
    text = _get_scopes_text(target_doc, scopes)
    violations = []
    has_city = bool(re.search('г\\.\\s*[А-ЯЁ][а-яё]+|город\\s+[А-ЯЁ]|Москв[аеы]|Санкт-Петербург', text))
    if not has_city:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует город подписания (например «г. Москва»)'})
    months = 'январ[яь]|феврал[яь]|март[а]?|апрел[яь]|ма[йя]|июн[яь]|июл[яь]|август[а]?|сентябр[яь]|октябр[яь]|ноябр[яь]|декабр[яь]'
    has_date = bool(re.search('\\d{2}[.\\s]+\\d{2}[.\\s]+\\d{4}|[«"\\\'"]?\\d{1,2}[»"\\\'"]?\\s*(?:' + months + ')', text, re.IGNORECASE))
    if has_date:
        placeholder = re.search('_+\\._+\\._+', text)
        if placeholder:
            has_date = False
    if not has_date:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует или неполная', 'Различие': 'Отсутствует полная дата приказа'})
    return violations

def check_signatory(target_doc: Dict[str, Any], config: Any, rule_index: int=7, rule_title: str='Проверка должности и ФИО подписанта.', scopes: List[str]=None) -> List[Dict[str, Any]]:
    """
    Проверяет наличие должности и ФИО подписанта.

    Должность: любое словосочетание (директор, начальник, управляющий и др.)
    ФИО: слово + инициалы (Иванов А.В.) или инициалы + слово (А.В. Иванов)
    Переведено на non-LLM из-за систематической ошибки LLM с фамилиями на -ович
    """
    if scopes is None:
        scopes = ['должность_фио_подписанта']
    text = _get_scopes_text(target_doc, scopes)
    violations = []
    if not text.strip():
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Блок подписанта пуст'})
        return violations
    if re.search('ПРИКАЗЫВАЮ', text, re.IGNORECASE):
        paragraphs = [p.strip() for p in re.split('\\n\\s*\\n', text) if p.strip()]
        if paragraphs:
            last_block = paragraphs[-1]
            if re.match('\\d+\\.\\s', last_block):
                text = ''
            else:
                text = last_block
    fio_pattern = '[А-ЯЁа-яё]{2,}\\s+[А-ЯЁ]\\s*\\.\\s*[А-ЯЁ]\\s*\\.|[А-ЯЁ]\\s*\\.\\s*[А-ЯЁ]\\s*\\.\\s*[А-ЯЁа-яё]{2,}'
    has_fio = bool(re.search(fio_pattern, text))
    fio_placeholders = re.search('И\\.О\\.\\s*Фамилия|Фамилия\\s*И\\.О\\.|(?<![А-ЯЁа-яё])ФИО(?![А-ЯЁа-яё])', text)
    if fio_placeholders and (not has_fio):
        has_fio = False
    if not has_fio:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': text.strip()[:100], 'Различие': 'ФИО подписанта не заполнено или содержит плейсхолдер'})
    dolzhnost_placeholder = re.search('указать\\s+наименование\\s+должности|\\(должность\\)', text, re.IGNORECASE)
    text_no_fio = re.sub(fio_pattern, '', text)
    text_no_fio = re.sub('[_\\-—/\\\\«»"\\\'"().\\d]', ' ', text_no_fio)
    words = [w for w in text_no_fio.split() if len(w) >= 3 and w not in ('М.П.', 'МП')]
    has_dolzhnost = len(words) >= 1 and (not dolzhnost_placeholder)
    if not has_dolzhnost:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': text.strip()[:100], 'Различие': 'Должность подписанта не указана'})
    return violations

def check_fio_after_marker(target_doc: Dict[str, Any], config: Any, rule_index: int, rule_title: str, marker_phrase: str, scopes: List[str]=None) -> List[Dict[str, Any]]:
    """
    Проверяет наличие должности и ФИО после фразы-маркера в тексте приказа.

    Используется для проверки заполненности назначенных лиц (организатор, секретарь).
    Ищет marker_phrase → после неё должна быть должность + ФИО.
    """
    if scopes is None:
        scopes = ['текст_приказа']
    text = _get_scopes_text(target_doc, scopes)
    violations = []
    match = re.search(re.escape(marker_phrase), text, re.IGNORECASE)
    if not match:
        return violations
    after_marker = text[match.end():]
    next_punkt = re.search('\\n\\d+\\.', after_marker)
    if next_punkt:
        after_marker = after_marker[:next_punkt.start()]
    fio_pattern_short = '[А-ЯЁа-яё]{2,}\\s+[А-ЯЁ]\\s*\\.\\s*[А-ЯЁ]\\s*\\.|[А-ЯЁ]\\s*\\.\\s*[А-ЯЁ]\\s*\\.\\s*[А-ЯЁа-яё]{2,}'
    fio_pattern_full = '[А-ЯЁ][а-яё]{2,}\\s+[А-ЯЁ][а-яё]{2,}\\s+[А-ЯЁ][а-яё]{2,}'
    fio_pattern = fio_pattern_short + '|' + fio_pattern_full
    fio_match = re.search(fio_pattern, after_marker)
    has_fio = bool(fio_match)
    has_placeholder = bool(re.search('И\\.О\\.\\s*Фамилия|ФИО|_{4,}|\\(указать\\)', after_marker, re.IGNORECASE))
    has_dolzhnost = True
    if has_fio and fio_match:
        text_before_fio = after_marker[:fio_match.start()].strip()
        text_before_fio = re.sub('на\\s+(?:производственной\\s+)?площадке', '', text_before_fio, flags=re.IGNORECASE)
        words_before = [w for w in re.findall('[А-ЯЁа-яё]{3,}', text_before_fio) if w.lower() not in ('на', 'по', 'для', 'при', 'из', 'от', 'над')]
        has_dolzhnost = len(words_before) >= 1
    if has_placeholder or not has_fio:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': after_marker.strip()[:150], 'Различие': 'ФИО после маркера не заполнено' if not has_fio else 'Содержит плейсхолдер вместо реального ФИО'})
    elif not has_dolzhnost:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': after_marker.strip()[:150], 'Различие': 'Должность перед ФИО не указана'})
    return violations

@register('prikaz_comp_ppu', 4)
def check_prikaz_number_comp_ppu(target_doc, config):
    """Правило #4: ПРИКАЗ + номер для Приказа о конкурсах ППУ."""
    return check_prikaz_and_number(target_doc, config, 4, 'Проверка наличия слова ПРИКАЗ и номера приказа.')

@register('prikaz_comp_ppu', 5)
def check_city_date_comp_ppu(target_doc, config):
    """Правило #5: город + дата для Приказа о конкурсах ППУ."""
    return check_city_and_date(target_doc, config, 5, 'Проверка наличия города и даты приказа.')

@register('prikaz_comp_ppu', 7)
def check_signatory_comp_ppu(target_doc, config):
    """Правило #7: должность + ФИО подписанта для Приказа о конкурсах ППУ."""
    return check_signatory(target_doc, config, 7, 'Проверка должности и ФИО подписанта.')

@register('prikaz_ppu', 4)
def check_prikaz_number_ppu(target_doc, config):
    """Правило #4: ПРИКАЗ + номер для Приказа о ППУ."""
    return check_prikaz_and_number(target_doc, config, 4, 'Проверка наличия слова ПРИКАЗ и номера приказа.')

@register('prikaz_ppu', 5)
def check_city_date_ppu(target_doc, config):
    """Правило #5: город + дата для Приказа о ППУ."""
    return check_city_and_date(target_doc, config, 5, 'Проверка наличия города и даты приказа.')

@register('prikaz_ppu', 7)
def check_signatory_ppu(target_doc, config):
    """Правило #7: должность + ФИО подписанта для Приказа о ППУ."""
    return check_signatory(target_doc, config, 7, 'Проверка должности и ФИО подписанта.')

@register('prikaz_vyhod', 4)
def check_prikaz_number_vyhod(target_doc, config):
    """Правило #4: ПРИКАЗ + номер для Приказа о проведении выхода."""
    return check_prikaz_and_number(target_doc, config, 4, 'Проверка наличия слова ПРИКАЗ и номера приказа')

@register('prikaz_vyhod', 5)
def check_city_date_vyhod(target_doc, config):
    """Правило #5: город + дата для Приказа о проведении выхода."""
    return check_city_and_date(target_doc, config, 5, 'Проверка наличия города и даты приказа')

@register('prikaz_vyhod', 7)
def check_signatory_vyhod(target_doc, config):
    """Правило #7: должность + ФИО подписанта для Приказа о проведении выхода."""
    return check_signatory(target_doc, config, 7, 'Проверка наличия должности и ФИО подписанта')

@register('prikaz_vyhod', 8)
def check_organizer_vyhod(target_doc, config):
    """Правило #8: ФИО организатора (п.1) для Приказа о проведении выхода."""
    return check_fio_after_marker(target_doc, config, 8, 'Проверка заполнения ФИО организатора (п.1)', 'назначить организатором проведения обхода')

@register('prikaz_vyhod', 9)
def check_secretary_vyhod(target_doc, config):
    """Правило #9: ФИО секретаря (п.2) для Приказа о проведении выхода."""
    return check_fio_after_marker(target_doc, config, 9, 'Проверка заполнения ФИО секретаря (п.2)', 'назначить секретарем проведения обхода')

def _extract_paragraph(text: str, start_re: str, end_re: str) -> str:
    """Извлекает текст параграфа между start_re и end_re (regex-маркеры)."""
    start = re.search(start_re, text)
    if not start:
        return ''
    rest = text[start.end():]
    end = re.search(end_re, rest)
    if end:
        return rest[:end.start()].strip()
    return rest.strip()

def _trim_to_section_2(reglament: str) -> str:
    """Обрезает регламент до раздела 2 — чтобы «1.2.1.» не совпадал с «2.1.» в regex."""
    section_2 = re.search('(?:^|\\n)\\s*2\\.\\s+[А-ЯЁA-Z]', reglament)
    if section_2:
        return reglament[section_2.start():]
    return reglament

def _check_reglament_times(reglament: str, rule_index: int, rule_title: str) -> List[Dict[str, Any]]:
    """
    Проверяет формат времени в п.2.7 (12-00) и п.2.8 (17-00) регламента.

    Возвращает violations если время обрезано/отсутствует.
    """
    violations = []
    reglament = _trim_to_section_2(reglament)
    text_27 = _extract_paragraph(reglament, '2\\.7\\.?\\s', '\\n\\s*2\\.8')
    if text_27 and (not re.search('12[\\-:\\.]\\s*00', text_27)):
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': f'п.2.7: {text_27[:200]}', 'Различие': 'Время «12-00» не указано или неполное'})
    text_28 = _extract_paragraph(reglament, '2\\.8\\.?\\s', '\\n\\s*2\\.9')
    if text_28 and (not re.search('17[\\-:\\.]\\s*00', text_28)):
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': f'п.2.8: {text_28[:200]}', 'Различие': 'Время «17-00» не указано или неполное'})
    return violations

def _normalize_org_form(text: str) -> str:
    """Нормализует юрформу: 000/OOO → ООО (OCR-толерантно)."""
    return text.replace('000', 'ООО').replace('OOO', 'ООО').replace('0OO', 'ООО').replace('OO0', 'ООО')

def _extract_org_names(text: str) -> list:
    """
    Извлекает все упоминания юрлиц из текста: [(позиция, raw, normalized), ...].

    OCR-толерантно: ООО = 000 = OOO.
    """
    org_forms = '(ООО|000|OOO|ЗАО|АО|ПАО|ОАО)'
    quotes = '[«"\\\'\\u201c\\u201e]([^»"\\\'\\u201d\\u201f]{1,100})[»"\\\'\\u201d\\u201f]'
    results = []
    for m in re.finditer(org_forms + '\\s*' + quotes, text):
        form = m.group(1)
        name = m.group(2).strip()
        norm_form = _normalize_org_form(form)
        norm = f'{norm_form} "{name}"'
        results.append((m.start(), m.group(0), norm))
    return results

def check_company_name_cross(target_doc: Dict[str, Any], config: Any, rule_index: int=10, rule_title: str='Сверка наименования компании.', header_scope: str='шапка', body_scope: str='приложение_2_к_приказу') -> List[Dict[str, Any]]:
    """
    Проверяет что наименование компании из шапки совпадает с упоминаниями в приложении_2.

    Извлекает эталонное наименование юрлица из шапки (первое упоминание),
    затем ищет все упоминания юрлиц в body_scope и сравнивает с эталоном.
    OCR-толерантно: ООО = 000 = OOO (кириллица/латиница/нули).
    """
    violations = []
    header = target_doc.get(header_scope, '')
    body = target_doc.get(body_scope, '')
    if not header or not body:
        return violations
    header_orgs = _extract_org_names(header)
    body_orgs = _extract_org_names(body)
    if not header_orgs:
        return violations
    _, etalon_raw, etalon_norm = header_orgs[0]
    for pos, raw, norm in body_orgs:
        if norm != etalon_norm:
            before_text = body[:pos]
            punkt_matches = list(re.finditer('(\\d+\\.\\d+(?:\\.\\d+)?)', before_text))
            punkt = punkt_matches[-1].group(1) if punkt_matches else '?'
            violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': f'{etalon_raw} (п.{punkt}: {raw})', 'Различие': f'Наименование в п.{punkt} отличается от эталонного'})
    return violations

def check_org_name_in_header(target_doc: Dict[str, Any], config: Any, rule_index: int=6, rule_title: str='Проверка юр.формы и наименования организации.', scopes: List[str]=None) -> List[Dict[str, Any]]:
    """
    Проверяет наличие наименования организации с юр. формой в шапке.

    Ищет паттерн: юр.форма (ООО/ЗАО/АО/ПАО/ОАО) + наименование.
    OCR-толерантно: ○○○ (circles) / 000 (нули) / OOO (латиница) → ООО.
    """
    if scopes is None:
        scopes = ['шапка']
    text = _get_scopes_text(target_doc, scopes)
    violations = []
    if not text.strip():
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Текст шапки пуст'})
        return violations
    normalized = text.replace('○○○', 'ООО')
    normalized = _normalize_org_form(normalized)
    has_org = bool(re.search('(ООО|ОАО|ЗАО|АО|ПАО|НАО)\\s*[«"\\\']?.+', normalized))
    if not has_org:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует наименование организации с юридической формой (ООО, ЗАО, АО и т.п.) в шапке'})
    return violations

@register('prikaz_ic_potoka', 6)
def check_org_name_ic_potoka(target_doc, config):
    """Правило #6: юр.форма + наименование в шапке для Приказа о создании ИЦ потока."""
    return check_org_name_in_header(target_doc, config, 6, 'Проверка юр.формы и наименования организации.', scopes=['шапка'])

@register('prikaz_ic_potoka', 4)
def check_prikaz_number_ic_potoka(target_doc, config):
    """Правило #4: ПРИКАЗ + номер для Приказа о создании ИЦ потока."""
    return check_prikaz_and_number(target_doc, config, 4, 'Проверка слова ПРИКАЗ и номера.', scopes=['шапка'])

@register('prikaz_ic_potoka', 5)
def check_city_date_ic_potoka(target_doc, config):
    """Правило #5: город + дата для Приказа о создании ИЦ потока."""
    return check_city_and_date(target_doc, config, 5, 'Проверка города и даты приказа.', scopes=['шапка'])

@register('prikaz_ic_potoka', 7)
def check_signatory_ic_potoka(target_doc, config):
    """Правило #7: должность + ФИО подписанта для Приказа о создании ИЦ потока."""
    return check_signatory(target_doc, config, 7, 'Проверка подписанта.', scopes=['подписант'])

@register('prikaz_ic_potoka', 8)
def check_responsible_fio_ic_potoka(target_doc, config):
    """
    Правило #8: ФИО ответственных лиц для Приказа о создании ИЦ потока.

    Проверяет 4 позиции:
    A) текст_приказа — пункт с «ознакомить»/«ознакомление» → должность+ФИО
    B) приложение_2_регламент п.2.1 — ответственный за ИЦ → ФИО
    C) приложение_2_регламент п.2.2 — исполняющий обязанности → ФИО
    D) приложение_2_регламент п.2.3 — администратор → ФИО

    Переведено с LLM на non-LLM из-за систематической ошибки:
    LLM не распознаёт ФИО после «исполняющему обязанности» (позиция C).
    Также: context_filter LLM жёстко привязан к нумерации пунктов,
    non-LLM ищет семантически — устойчив к сдвигу нумерации.
    """
    violations = []
    rule_index = 8
    rule_title = 'Проверка ФИО ответственных лиц.'
    fio_pattern = '[А-ЯЁа-яё]{2,}\\s+[А-ЯЁ]\\s*\\.\\s*[А-ЯЁ]\\s*\\.|[А-ЯЁ]\\s*\\.\\s*[А-ЯЁ]\\s*\\.\\s*[А-ЯЁа-яё]{2,}'
    text_prikaz = target_doc.get('текст_приказа', '')
    pos_a_paragraph = ''
    for line in text_prikaz.split('\n'):
        if re.search('ознакомл|ознакомит', line, re.IGNORECASE):
            pos_a_paragraph = line.strip()
            break
    if not pos_a_paragraph:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': text_prikaz.strip()[:200], 'Различие': 'Позиция A: пункт с «ознакомлением» отсутствует в тексте приказа'})
    elif not re.search(fio_pattern, pos_a_paragraph):
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': pos_a_paragraph[:200], 'Различие': 'Позиция A: в пункте с «ознакомлением» не указано ФИО ответственного'})
    reglament = target_doc.get('приложение_2_регламент', '')
    if not reglament or not reglament.strip() or '[НЕТ СТРАНИЦ' in reglament or ('[Фильтр: не найдено]' in reglament):
        return violations
    reglament = _trim_to_section_2(reglament)
    text_21 = _extract_paragraph(reglament, '2\\.1\\.?\\s', '\\n\\s*2\\.2')
    if text_21 and (not re.search(fio_pattern, text_21)):
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': f'п.2.1: {text_21[:200]}', 'Различие': 'Позиция B: в п.2.1 регламента не указано ФИО ответственного за ИЦ'})
    text_22 = _extract_paragraph(reglament, '2\\.2\\.?\\s', '\\n\\s*2\\.3')
    if text_22:
        io_match = re.search('исполняющ\\w*\\s+обязанност\\w*', text_22, re.IGNORECASE)
        if io_match:
            after_io = text_22[io_match.end():]
            if not re.search(fio_pattern, after_io):
                violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': f'п.2.2: {text_22[:200]}', 'Различие': 'Позиция C: после «исполняющему обязанности» не указано ФИО'})
    text_23 = _extract_paragraph(reglament, '2\\.3\\.?\\s', '\\n\\s*2\\.4')
    if text_23:
        for line in text_23.split('\n'):
            if re.search('[Аа]дминистратор', line):
                if not re.search(fio_pattern, line):
                    violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': f'п.2.3: {line.strip()[:200]}', 'Различие': 'Позиция D: для администратора ИЦ не указано ФИО'})
                break
    return violations

def check_signatory_with_stamp(target_doc: Dict[str, Any], config: Any, rule_index: int=9, rule_title: str='Проверка подписи директора и М.П.', scopes: List[str]=None) -> List[Dict[str, Any]]:
    """
    Проверяет наличие должности, ФИО и печати (М.П./МП/ПЕЧАТЬ) в подписанте.

    Переведено с LLM на non-LLM: gpt-4.1-mini не видит «М.П.» в собственном
    контексте (hallucination) и не распознаёт «МП» без точек как аналог «М.П.».
    """
    if scopes is None:
        scopes = ['подписант']
    text = _get_scopes_text(target_doc, scopes)
    violations = []
    if not text or len(text.strip()) < 3:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': '(текст подписанта отсутствует)', 'Различие': 'Отсутствует блок подписанта'})
        return violations
    position_words = ['директор', 'руководитель', 'начальник', 'управляющ', 'заместител', 'президент', 'председател']
    text_lower = text.lower()
    has_position = any((w in text_lower for w in position_words))
    if not has_position:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': text.strip()[:200], 'Различие': 'Отсутствует должность подписанта (Генеральный директор и т.п.)'})
    fio_pattern = '[А-ЯЁ][а-яё]+\\s+[А-ЯЁ]\\.[А-ЯЁ]\\.|[А-ЯЁ]\\.[А-ЯЁ]\\.\\s*[А-ЯЁ][а-яё]+'
    placeholder_pattern = '_{4,}|И\\.О\\.\\s*Фамилия|Фамилия\\s*И\\.О\\.'
    has_fio = bool(re.search(fio_pattern, text))
    has_placeholder = bool(re.search(placeholder_pattern, text))
    if not has_fio or has_placeholder:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': text.strip()[:200], 'Различие': 'ФИО подписанта не заполнено (плейсхолдер или отсутствует)'})
    stamp_pattern = 'М\\.?\\s*П\\.?|ПЕЧАТЬ|печать'
    has_stamp = bool(re.search(stamp_pattern, text))
    if not has_stamp:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': text.strip()[:200], 'Различие': 'Отсутствует указание на печать (М.П., МП или ПЕЧАТЬ)'})
    return violations

@register('prikaz_formirovanie_po', 3)
def check_rekvizity_formirovanie_po(target_doc, config):
    """
    Правило #3: реквизиты приказа — юрлицо, дата, номер, город.

    Проверяет наличие 4 обязательных реквизитов в шапке документа:
    1) Юрлицо: ООО/ЗАО/АО/ПАО + наименование
    2) Дата: числовой (ДД.ММ.ГГГГ) или словесный («ДД» месяц ГГГГ г.)
    3) Номер приказа: символ № + непустое значение
    4) Город: г. + название
    """
    scopes = ['шапка']
    text = _get_scopes_text(target_doc, scopes)
    violations = []
    rule_index = 3
    rule_title = 'Проверка реквизитов приказа'
    has_org = bool(re.search('(ООО|ЗАО|АО|ПАО)\\s*[«"\\\']?.+', text))
    if not has_org:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует наименование юрлица (ООО/ЗАО/АО/ПАО)'})
    months = 'январ[яь]|феврал[яь]|март[а]?|апрел[яь]|ма[йя]|июн[яь]|июл[яь]|август[а]?|сентябр[яь]|октябр[яь]|ноябр[яь]|декабр[яь]'
    has_date = bool(re.search('\\d{2}[.\\s]+\\d{2}[.\\s]+\\d{4}|[«"\\\'"]?\\d{1,2}[»"\\\'"]?\\s*(?:' + months + ')', text, re.IGNORECASE))
    if has_date and re.search('_+\\._+\\._+', text):
        has_date = False
    if not has_date:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует дата приказа'})
    number_match = re.search('№\\s*(.+)', text)
    if number_match:
        after_sign = number_match.group(1).strip()
        cleaned = re.sub('[_\\s]', '', after_sign)
        is_bn = bool(re.match('^б[/\\\\]?н$', cleaned, re.IGNORECASE))
        if not cleaned or is_bn:
            violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': f'№ {after_sign}', 'Различие': 'Номер приказа не заполнен'})
    else:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует символ № с номером приказа'})
    has_city = bool(re.search('г\\.\\s*[А-ЯЁ][а-яё]+|город\\s+[А-ЯЁ]|Москв[аеы]|Санкт-Петербург', text))
    if not has_city:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует город подписания (например «г. Москва»)'})
    return violations

@register('prikaz_formirovanie_po', 4)
def check_stamp_formirovanie_po(target_doc, config):
    """Правило #4: должность + ФИО + М.П./ПЕЧАТЬ для Приказа о формировании ПО."""
    return check_signatory_with_stamp(target_doc, config, 4, 'Проверка подписанта и М.П.')

@register('prikaz_ic', 2)
def check_rekvizity_ic(target_doc, config):
    """
    Правило #2: реквизиты приказа (OCR-толерантный) для Приказа о создании ИЦ.

    Проверяет 5 обязательных реквизитов в шапке:
    1) Юрлицо: ООО/ЗАО/АО/ПАО (+ OCR: 000=ООО)
    2) ПРИКАЗ
    3) Дата
    4) Номер
    5) Город
    """
    scopes = ['шапка']
    text = _get_scopes_text(target_doc, scopes)
    violations = []
    rule_index = 2
    rule_title = 'Проверка реквизитов приказа'
    has_org = bool(re.search('(ООО|000|OOO|ЗАО|АО|ПАО)\\s*[«"\\\']?.+', text))
    if not has_org:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует наименование юрлица (ООО/ЗАО/АО/ПАО)'})
    has_prikaz = bool(re.search('(?:П\\s*Р\\s*И\\s*К\\s*А\\s*З|Приказ)(?![А-ЯЁа-яё])', text, re.IGNORECASE))
    if not has_prikaz:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует слово «ПРИКАЗ»'})
    months = 'январ[яь]|феврал[яь]|март[а]?|апрел[яь]|ма[йя]|июн[яь]|июл[яь]|август[а]?|сентябр[яь]|октябр[яь]|ноябр[яь]|декабр[яь]'
    has_date = bool(re.search('\\d{2}[.\\s]+\\d{2}[.\\s]+\\d{4}|[«"\\\'"]?\\d{1,2}[»"\\\'"]?\\s*(?:' + months + ')', text, re.IGNORECASE))
    if has_date and re.search('_+\\._+\\._+', text):
        has_date = False
    if not has_date:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует дата приказа'})
    number_match = re.search('№\\s*(.+)', text)
    if number_match:
        after_sign = number_match.group(1).strip()
        cleaned = re.sub('[_\\s]', '', after_sign)
        is_bn = bool(re.match('^б[/\\\\]?н$', cleaned, re.IGNORECASE))
        if not cleaned or is_bn:
            violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': f'№ {after_sign}', 'Различие': 'Номер приказа не заполнен'})
    else:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует символ № с номером приказа'})
    has_city = bool(re.search('г\\.\\s*[А-ЯЁ][а-яё]+|город\\s+[А-ЯЁ]|Москв[аеы]|Санкт-Петербург', text))
    if not has_city:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует город подписания (например «г. Москва»)'})
    return violations

@register('prikaz_ic', 4)
def check_signatory_ic(target_doc, config):
    """Правило #4: должность + ФИО подписанта для Приказа о создании ИЦ."""
    return check_signatory(target_doc, config, 4, 'Проверка подписанта.', scopes=['шапка'])

@register('prikaz_ic', 6)
def check_responsible_fio_ic(target_doc, config):
    """
    Правило #6: ФИО ответственных лиц в регламенте для Приказа о создании ИЦ.

    Переиспользует логику check_responsible_fio_ic_potoka с маппингом scope:
    prikaz_ic использует "приложение_2_к_приказу", а не "приложение_2_регламент".
    """
    adapted_doc = dict(target_doc)
    adapted_doc['приложение_2_регламент'] = target_doc.get('приложение_2_к_приказу', '')
    violations = check_responsible_fio_ic_potoka(adapted_doc, config)
    for v in violations:
        v['rule_index'] = 6
        v['rule_title'] = 'Проверка ФИО ответственных лиц в регламенте.'
    return violations

@register('prikaz_ic_el', 2)
def check_rekvizity_ic_el(target_doc, config):
    """Правило #2: реквизиты приказа (OCR-толерантный) для Приказа о создании ИЦ (эл.)."""
    return check_rekvizity_ic(target_doc, config)

@register('prikaz_ic_el', 4)
def check_signatory_ic_el(target_doc, config):
    """Правило #4: должность + ФИО подписанта для Приказа о создании ИЦ (эл.)."""
    return check_signatory(target_doc, config, 4, 'Проверка подписанта.', scopes=['шапка'])

@register('prikaz_ic_el', 6)
def check_responsible_fio_ic_el(target_doc, config):
    """
    Правило #6: заполненность регламента для Приказа о создании ИЦ (эл.).

    Проверяет:
    - ФИО в 4 позициях (A-D) — через check_responsible_fio_ic_potoka
    - Время в п.2.7 (12-00) и п.2.8 (17-00) — через _check_reglament_times
    Маппинг scope: prikaz_ic_el использует "приложение_2_к_приказу".
    """
    adapted_doc = dict(target_doc)
    reglament = target_doc.get('приложение_2_к_приказу', '')
    adapted_doc['приложение_2_регламент'] = reglament
    violations = check_responsible_fio_ic_potoka(adapted_doc, config)
    if reglament and reglament.strip():
        violations.extend(_check_reglament_times(reglament, 6, 'Проверка заполненности регламента.'))
    for v in violations:
        v['rule_index'] = 6
        v['rule_title'] = 'Проверка заполненности регламента.'
    return violations

@register('prikaz_ic_el', 10)
def check_company_cross_ic_el(target_doc, config):
    """Правило #10: сверка наименования компании шапка↔приложение_2 для ИЦ (эл.)."""
    return check_company_name_cross(target_doc, config, 10, 'Сверка наименования компании.')

non_llm_checks_prikaz_checks_module = SimpleNamespace(_get_scopes_text=_get_scopes_text, check_prikaz_and_number=check_prikaz_and_number, check_city_and_date=check_city_and_date, check_signatory=check_signatory, check_fio_after_marker=check_fio_after_marker, check_prikaz_number_comp_ppu=check_prikaz_number_comp_ppu, check_city_date_comp_ppu=check_city_date_comp_ppu, check_signatory_comp_ppu=check_signatory_comp_ppu, check_prikaz_number_ppu=check_prikaz_number_ppu, check_city_date_ppu=check_city_date_ppu, check_signatory_ppu=check_signatory_ppu, check_prikaz_number_vyhod=check_prikaz_number_vyhod, check_city_date_vyhod=check_city_date_vyhod, check_signatory_vyhod=check_signatory_vyhod, check_organizer_vyhod=check_organizer_vyhod, check_secretary_vyhod=check_secretary_vyhod, _extract_paragraph=_extract_paragraph, _trim_to_section_2=_trim_to_section_2, _check_reglament_times=_check_reglament_times, _normalize_org_form=_normalize_org_form, _extract_org_names=_extract_org_names, check_company_name_cross=check_company_name_cross, check_org_name_in_header=check_org_name_in_header, check_org_name_ic_potoka=check_org_name_ic_potoka, check_prikaz_number_ic_potoka=check_prikaz_number_ic_potoka, check_city_date_ic_potoka=check_city_date_ic_potoka, check_signatory_ic_potoka=check_signatory_ic_potoka, check_responsible_fio_ic_potoka=check_responsible_fio_ic_potoka, check_signatory_with_stamp=check_signatory_with_stamp, check_rekvizity_formirovanie_po=check_rekvizity_formirovanie_po, check_stamp_formirovanie_po=check_stamp_formirovanie_po, check_rekvizity_ic=check_rekvizity_ic, check_signatory_ic=check_signatory_ic, check_responsible_fio_ic=check_responsible_fio_ic, check_rekvizity_ic_el=check_rekvizity_ic_el, check_signatory_ic_el=check_signatory_ic_el, check_responsible_fio_ic_el=check_responsible_fio_ic_el, check_company_cross_ic_el=check_company_cross_ic_el)

# END_SOURCE_NON_LLM_CHECKS_PRIKAZ_CHECKS

# START_SOURCE_NON_LLM_CHECKS_PRIKAZ_TIRAZH_CHECKS
# PURPOSE: Inlined source from audit_engine/non_llm_checks/prikaz_tirazh_checks.py.
def _check_number_and_date(text: str, rule_index: int, rule_title: str) -> List[Dict[str, Any]]:
    """
    Проверяет наличие номера приказа и полной даты в тексте.

    Args:
        text: OCR-текст приказа (страницы 1-2)
        rule_index: индекс правила
        rule_title: заголовок правила

    Returns:
        List[Dict] — список нарушений (пустой = норма)
    """
    violations = []
    prikaz_split = re.split('ПРИКАЗЫВАЮ', text, maxsplit=1, flags=re.IGNORECASE)
    header_text = prikaz_split[0] if prikaz_split else text[:500]
    number_match = re.search('(?:№|No|Ne|N[еo°⁰])\\s*(.+)', header_text, re.IGNORECASE)
    if number_match:
        after_sign = number_match.group(1).strip()
        after_sign = re.split('\\n|от\\s', after_sign)[0].strip()
        cleaned = re.sub('[_\\s]', '', after_sign)
        if not cleaned:
            violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': f'№ {after_sign}', 'Различие': 'Номер приказа не заполнен (пустой или подчёркивания после №)'})
    else:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Отсутствует символ № с номером приказа'})
    date_match = re.search('(\\d{2})[.\\s]+(\\d{2})[.\\s]+(\\d{4})', header_text)
    months = 'январ[яь]|феврал[яь]|март[а]?|апрел[яь]|ма[йя]|июн[яь]|июл[яь]|август[а]?|сентябр[яь]|октябр[яь]|ноябр[яь]|декабр[яь]'
    verbal_date = re.search('\\d{1,2}\\s*(?:' + months + ')\\s*\\d{4}', header_text, re.IGNORECASE)
    has_date = bool(date_match) or bool(verbal_date)
    if has_date and re.search('_+[.\\s]*_+[.\\s]*_+', header_text):
        if not date_match and (not verbal_date):
            has_date = False
    if not has_date:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует или неполная', 'Различие': 'Отсутствует полная дата приказа (ожидается ДД.ММ.ГГГГ)'})
    return violations

def _check_date_in_punkt_11(text: str, rule_index: int, rule_title: str) -> List[Dict[str, Any]]:
    """
    Проверяет наличие полной даты в пункте 11 «Отменить действие приказа от...».

    Args:
        text: OCR-текст приказа (страницы 1-2)
        rule_index: индекс правила
        rule_title: заголовок правила

    Returns:
        List[Dict] — список нарушений (пустой = норма)
    """
    violations = []
    punkt_match = re.search('11\\s*[.)\\s]\\s*(Отменить[^\\n]*(?:\\n(?!\\d+\\s*[.)])[^\\n]*)*)', text, re.IGNORECASE)
    if not punkt_match:
        punkt_match = re.search('(Отменить\\s+действие\\s+приказа[^\\n]*)', text, re.IGNORECASE)
    if not punkt_match:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'пункт 11 не найден', 'Различие': 'Не удалось найти пункт 11 «Отменить действие приказа» в тексте'})
        return violations
    punkt_text = punkt_match.group(1)
    date_after_ot = re.search('от\\s+(\\d{2})[.\\s]+(\\d{2})[.\\s]+(\\d{4})', punkt_text)
    date_anywhere = re.search('(\\d{2})[.\\s]+(\\d{2})[.\\s]+(\\d{4})', punkt_text)
    has_date = bool(date_after_ot) or bool(date_anywhere)
    has_placeholder = bool(re.search('_+[.\\s]*_+[.\\s]*_+|_+\\s*202_', punkt_text))
    if has_placeholder and (not date_after_ot):
        has_date = False
    if not has_date:
        violations.append({'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': punkt_text.strip()[:150], 'Различие': 'Дата отменяемого приказа не заполнена или неполная (ожидается ДД.ММ.ГГГГ)'})
    return violations

@register('prikaz_tirazh', 7)
def check_date_punkt_11_tirazh(target_doc: Dict[str, Any], config: Any) -> List[Dict[str, Any]]:
    """Правило #7: дата отменяемого приказа в п.11 для Приказа о тиражировании."""
    text = target_doc.get('текст_приказа', '')
    return _check_date_in_punkt_11(text, 7, 'Проверка даты отменяемого приказа в пункте 11.')

non_llm_checks_prikaz_tirazh_checks_module = SimpleNamespace(_check_number_and_date=_check_number_and_date, _check_date_in_punkt_11=_check_date_in_punkt_11, check_date_punkt_11_tirazh=check_date_punkt_11_tirazh)

# END_SOURCE_NON_LLM_CHECKS_PRIKAZ_TIRAZH_CHECKS

# START_SOURCE_NON_LLM_CHECKS_SHAPKA_ELEMENTS
# PURPOSE: Inlined source from audit_engine/non_llm_checks/shapka_elements.py.
def check_shapka_elements_universal(target_doc: Dict[str, Any], config: Any, rule_index: int, rule_title: str) -> List[Dict[str, Any]]:
    """
    Проверяет наличие обязательных элементов в шапке документа.

    Обязательные элементы:
    1. Слово «Приложение» (регистронезависимо)
    2. Номер приложения — «№» + цифры/буквы
    3. Ссылка на приказ — «приказ» или «приказу» (регистронезависимо)
    """
    shapka = target_doc.get('шапка', '')
    if not shapka:
        return [{'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': 'отсутствует', 'Различие': 'Чанк «шапка» не найден в документе'}]
    shapka_lower = shapka.lower()
    missing = []
    if 'приложение' not in shapka_lower:
        missing.append("слово 'Приложение'")
    if not re.search('№\\s*\\S+', shapka):
        missing.append('номер приложения (№...)')
    if not re.search('приказ[а-яё]*', shapka_lower):
        missing.append("ссылка на приказ (слово 'приказ'/'приказу')")
    if missing:
        return [{'rule_index': rule_index, 'rule_title': rule_title, 'Целевой документ': shapka.strip()[:200], 'Различие': f'Отсутствуют элементы: {', '.join(missing)}'}]
    return []

@register('polozhenie_comp_ppu', 6)
def check_shapka_polozhenie_comp_ppu(target_doc, config):
    """Правило #6: проверка обязательных элементов шапки для Положения о конкурсах ППУ."""
    return check_shapka_elements_universal(target_doc, config, 6, 'Дополнительная сверка шапки с шаблоном.')

non_llm_checks_shapka_elements_module = SimpleNamespace(check_shapka_elements_universal=check_shapka_elements_universal, check_shapka_polozhenie_comp_ppu=check_shapka_polozhenie_comp_ppu)

# END_SOURCE_NON_LLM_CHECKS_SHAPKA_ELEMENTS
# END_NON_LLM_CHECKS

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

# START_SOURCE_DRIVERS_ANALYZER
# PURPOSE: Inlined source from audit_engine/drivers/analyzer.py.
@dataclass
class DriverEntry:
    number: str
    driver_name: str
    average_score: Optional[float]
    problem_comment: str
    notes: List[Dict[str, str]]

def _default_prompt_path() -> Path:
    return Path(__file__).resolve().parents[2] / 'doc_configs' / 'drivers' / 'driver_check_prompt.txt'

def load_driver_prompt() -> str:
    env_path = os.environ.get('DRIVERS_PROMPT_PATH')
    path = Path(env_path) if env_path else _default_prompt_path()
    return path.read_text(encoding='utf-8')

def parse_average(value: Optional[str]) -> Optional[float]:
    if value is None:
        return None
    cleaned = value.replace(',', '.').strip()
    if not cleaned:
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None

def extract_summary_driver_map(summary_text: str) -> Dict[str, str]:
    blocks: Dict[str, str] = {}
    current_key: Optional[str] = None
    current_lines: List[str] = []
    bullet_pattern = re.compile('^\\s*(\\d+)\\)\\s*(.*)$')
    code_prefix = re.compile('^[-–]?\\s*\\d+[?.]?\\s*')
    for raw_line in summary_text.splitlines():
        line = raw_line.strip()
        bullet = bullet_pattern.match(line)
        if bullet:
            if current_key is not None:
                blocks[current_key] = ' '.join(filter(None, current_lines)).strip()
            remainder = bullet.group(2).strip()
            remainder = code_prefix.sub('', remainder, count=1).strip()
            current_key = remainder.split(':', 1)[0].strip() if remainder else f'Driver{bullet.group(1)}'
            first_text = remainder.split(':', 1)[1].strip() if ':' in remainder else ''
            current_lines = [first_text] if first_text else []
        elif current_key is not None and line:
            current_lines.append(line)
    if current_key is not None:
        blocks[current_key] = ' '.join(filter(None, current_lines)).strip()
    return blocks

def to_driver_entries(section: Dict[str, object]) -> List[DriverEntry]:
    entries: List[DriverEntry] = []
    for q in section.get('questions', []):
        avg = parse_average((q.get('scores') or {}).get('Средняя'))
        entries.append(DriverEntry(number=(q.get('number') or '').strip(), driver_name=(q.get('question') or '').strip(), average_score=avg, problem_comment=(q.get('problem_comment') or '').strip(), notes=q.get('notes', [])))
    return entries

def prepare_section_context(section: Dict[str, object], primary_threshold: float=9.0, fallback_threshold: float=7.0) -> Dict[str, object]:
    drivers = to_driver_entries(section)
    primary = [d for d in drivers if d.average_score is not None and d.average_score >= primary_threshold]
    selected_threshold = primary_threshold
    selected = primary
    if not selected:
        selected = [d for d in drivers if d.average_score is not None and d.average_score >= fallback_threshold]
        selected_threshold = fallback_threshold
    summary_text = section.get('summary', '') or ''
    summary_driver_map = extract_summary_driver_map(summary_text)
    return {'section_title': section.get('title'), 'score_thresholds': {'primary': primary_threshold, 'fallback': fallback_threshold}, 'selected_threshold': selected_threshold, 'eligible_drivers': [{'number': d.number, 'driver_name': d.driver_name, 'average_score': d.average_score, 'problem_comment': d.problem_comment, 'notes': d.notes} for d in selected], 'all_drivers': [{'number': d.number, 'driver_name': d.driver_name, 'average_score': d.average_score, 'problem_comment': d.problem_comment, 'notes': d.notes} for d in drivers], 'summary_text': summary_text, 'summary_driver_map': summary_driver_map}

def call_driver_llm(client: OpenAI, section_payload: Dict[str, object], model: str, temperature: float, system_prompt: str) -> Tuple[str, Any]:
    user_payload = json.dumps(section_payload, ensure_ascii=False, indent=2)
    try:
        response = client.chat.completions.create(model=model, temperature=temperature, response_format={'type': 'json_object'}, messages=[{'role': 'system', 'content': system_prompt}, {'role': 'user', 'content': user_payload}])
    except Exception:
        response = client.chat.completions.create(model=model, temperature=temperature, messages=[{'role': 'system', 'content': system_prompt}, {'role': 'user', 'content': user_payload}])
    return (response.choices[0].message.content or '', response)

def analyze_sections(parsed_data: Dict[str, Any], client: OpenAI, primary_threshold: float, fallback_threshold: float, model: str, temperature: float, progress_callback: Optional[Callable[[str, int, int], None]]=None, system_prompt: Optional[str]=None) -> List[Dict[str, Any]]:
    prompt = system_prompt or load_driver_prompt()
    section_results: List[Dict[str, Any]] = []
    total_sections = len(parsed_data.get('sections') or [])
    for idx, section in enumerate(parsed_data.get('sections') or [], 1):
        if progress_callback:
            title = section.get('title') or f'Секция {idx}'
            progress_callback(f'Анализ: {title}', idx, total_sections)
        payload = prepare_section_context(section, primary_threshold, fallback_threshold)
        answer_text, _ = call_driver_llm(client, payload, model, temperature, prompt)
        section_json = json.loads(answer_text)
        section_results.append(section_json)
    return section_results

def collect_remarks_and_summaries(section_jsons: Iterable[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], Dict[str, str]]:
    aggregated_remarks: List[Dict[str, Any]] = []
    driver_summary_map: Dict[str, str] = {}
    for section in section_jsons:
        section_title = section.get('section_title')
        for remark in section.get('remarks') or []:
            merged = dict(remark)
            if section_title and 'section_title' not in merged:
                merged['section_title'] = section_title
            aggregated_remarks.append(merged)
        for driver_name, summary_text in (section.get('summary_driver_map') or {}).items():
            normalized_name = (driver_name or '').strip()
            if not normalized_name:
                continue
            summary_text = (summary_text or '').strip()
            existing = driver_summary_map.get(normalized_name)
            if existing and existing != summary_text:
                driver_summary_map[normalized_name] = f'{existing}\n---\n{summary_text}'
            else:
                driver_summary_map[normalized_name] = summary_text
    return (aggregated_remarks, driver_summary_map)

def build_question_lookup(parsed_sections: Iterable[Dict[str, Any]]) -> Dict[Tuple[str, str], Dict[str, Any]]:
    lookup: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for section in parsed_sections:
        title = (section.get('title') or '').strip()
        for question in section.get('questions', []):
            key = (title, (question.get('number') or '').strip())
            lookup[key] = {'driver_name': (question.get('question') or '').strip(), 'problem_comment': (question.get('problem_comment') or '').strip(), 'notes': '\n'.join((note.get('text', '').strip() for note in question.get('notes', [])))}
    return lookup

def export_missing_driver_report(parsed_data: Dict[str, Any], section_results: Iterable[Dict[str, Any]], output_path: Path, sheet_name: str='Итог') -> None:
    question_lookup = build_question_lookup(parsed_data.get('sections', []))
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(['Блок', 'Номер драйвера', 'Наименование драйвера', 'Средняя оценка', 'Комментарий (проблема)', 'Дополнительные примечания', 'Замечание'])
    for section in section_results:
        title = (section.get('section_title') or '').strip()
        remark_lookup = {(remark.get('number'), remark.get('driver_name')): remark for remark in section.get('remarks') or []}
        for check in section.get('driver_checks') or []:
            if check.get('found_in_summary'):
                continue
            number = (check.get('number') or '').strip()
            driver_name = (check.get('driver_name') or '').strip()
            key = (title, number)
            parsed_info = question_lookup.get(key, {})
            remark = remark_lookup.get((number, driver_name), {})
            ws.append([title, number, driver_name or parsed_info.get('driver_name', ''), check.get('average_score'), parsed_info.get('problem_comment', ''), parsed_info.get('notes', ''), remark.get('issue', 'Нет в выводах')])
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

drivers_analyzer_module = SimpleNamespace(DriverEntry=DriverEntry, _default_prompt_path=_default_prompt_path, load_driver_prompt=load_driver_prompt, parse_average=parse_average, extract_summary_driver_map=extract_summary_driver_map, to_driver_entries=to_driver_entries, prepare_section_context=prepare_section_context, call_driver_llm=call_driver_llm, analyze_sections=analyze_sections, collect_remarks_and_summaries=collect_remarks_and_summaries, build_question_lookup=build_question_lookup, export_missing_driver_report=export_missing_driver_report)

# END_SOURCE_DRIVERS_ANALYZER

# START_SOURCE_DRIVERS
# PURPOSE: Inlined source from audit_engine/drivers/__init__.py.
drivers_logger = logging.getLogger(__name__)

def drivers__load_config() -> Dict[str, Any]:
    """Загрузка конфига из doc_configs/drivers/config.json."""
    config_path = Path(__file__).resolve().parents[2] / 'doc_configs' / 'drivers' / 'config.json'
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def drivers_run(args) -> AuditResult:
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
    config = drivers__load_config()
    model = args.model or config.get('model', 'gpt-4.1-mini')
    temperature = args.temperature if args.temperature is not None else config.get('temperature', 0.0)
    primary_threshold = config.get('primary_threshold', 9.0)
    fallback_threshold = config.get('fallback_threshold', 7.0)
    if args.session_dir:
        session_dir = Path(args.session_dir)
    else:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        session_dir = Path(__file__).resolve().parents[2] / 'logs_result' / 'drivers' / f'session_{timestamp}'
    session_dir.mkdir(parents=True, exist_ok=True)
    print(f'\n{'=' * 60}')
    print(f'  Аудит драйверов производительности')
    print(f'{'=' * 60}')
    print(f'  Файл: {target_path.name}')
    print(f'  Модель: {model}')
    print(f'  Пороги: primary={primary_threshold}, fallback={fallback_threshold}')
    print(f'  Сессия: {session_dir}')
    print()
    try:
        return drivers__run_pipeline(args, config, model, temperature, primary_threshold, fallback_threshold, session_dir, start_time)
    except Exception as e:
        duration = time.time() - start_time
        error_msg = f'{type(e).__name__}: {e}\n{traceback.format_exc()}'
        error_path = session_dir / 'ERROR.txt'
        error_path.write_text(error_msg, encoding='utf-8')
        drivers_logger.error('drivers: ошибка аудита → %s: %s', error_path, e)
        print(f'\n  ОШИБКА: {e}')
        print(f'  Лог ошибки: {error_path}')
        raise

def drivers__run_pipeline(args, config: Dict[str, Any], model: str, temperature: float, primary_threshold: float, fallback_threshold: float, session_dir: Path, start_time: float) -> AuditResult:
    """Основной пайплайн аудита драйверов."""
    target_path = Path(args.target)
    print('[1/4] Парсинг Excel...')
    parsed = parse_excel_to_json(target_path)
    parsed_path = session_dir / '01_parsed.json'
    with open(parsed_path, 'w', encoding='utf-8') as f:
        json.dump(parsed, f, ensure_ascii=False, indent=2)
    print(f'  Секций: {len(parsed.get('sections', []))}')
    print(f'  Сохранено: {parsed_path}')
    if args.parse_only:
        print(f'\n{'=' * 60}')
        print('TARGET (parsed):')
        print(json.dumps(parsed, ensure_ascii=False, indent=2))
        return AuditResult(doc_type='drivers', session_dir=session_dir, target_path=str(target_path), duration_sec=time.time() - start_time)
    print('\n[2/4] LLM-анализ секций...')
    api_key = os.environ.get('OPENAI_API_KEY', 'dummy')
    base_url = config.get('llm_base_url', LLM_CONFIG.base_url)
    client = OpenAI(api_key=api_key, base_url=base_url)
    model = resolve_runtime_llm_model(model)
    system_prompt = load_driver_prompt()

    def progress_cb(msg: str, idx: int, total: int):
        print(f'  [{idx}/{total}] {msg}')
    section_results = analyze_sections(parsed_data=parsed, client=client, primary_threshold=primary_threshold, fallback_threshold=fallback_threshold, model=model, temperature=temperature, progress_callback=progress_cb, system_prompt=system_prompt)
    analysis_path = session_dir / '02_llm_analysis.json'
    with open(analysis_path, 'w', encoding='utf-8') as f:
        json.dump(section_results, f, ensure_ascii=False, indent=2)
    print(f'  Сохранено: {analysis_path}')
    print('\n[3/4] Сбор замечаний...')
    remarks, driver_summary_map = collect_remarks_and_summaries(section_results)
    print(f'  Замечаний: {len(remarks)}')
    print('\n[4/4] Генерация Excel-отчёта...')
    report_path = session_dir / '03_report.xlsx'
    export_missing_driver_report(parsed, section_results, report_path)
    print(f'  Отчёт: {report_path}')
    duration = time.time() - start_time
    violations = [{'rule_index': f'driver_{r.get('section_title', 'x')}_{r.get('number', '?')}', 'rule_title': r.get('driver_name', ''), 'section_title': r.get('section_title', ''), 'issue': r.get('issue', '')} for r in remarks]
    print(f'\n{'=' * 60}')
    print(f'  Итого нарушений: {len(violations)}')
    if violations:
        for v in violations:
            print(f'  - [{v['rule_index']}] {v['rule_title']}: {v['issue']}')
    print(f'  Время: {duration:.1f} сек')
    print(f'  Сессия: {session_dir}')
    print(f'{'=' * 60}')
    return AuditResult(violations=violations, doc_type='drivers', session_dir=session_dir, duration_sec=duration, rules_checked=len(parsed.get('sections', [])), target_path=str(target_path))

drivers_module = SimpleNamespace(logger=drivers_logger, _load_config=drivers__load_config, run=drivers_run, _run_pipeline=drivers__run_pipeline)

# END_SOURCE_DRIVERS

# START_SOURCE_KPSC_RUN_VALIDATIONS
# PURPOSE: Inlined source from audit_engine/kpsc/run_validations.py.
kpsc_run_validations_PARSER_MODULES = ['audit_engine.kpsc.parser_scripts.parse_kpsc_header', 'audit_engine.kpsc.parser_scripts.parse_kpsc_table1', 'audit_engine.kpsc.parser_scripts.parse_legend', 'audit_engine.kpsc.parser_scripts.parse_loss_digitization', 'audit_engine.kpsc.parser_scripts.parse_pa1_chart', 'audit_engine.kpsc.parser_scripts.parse_pa1_table', 'audit_engine.kpsc.parser_scripts.parse_pokazateli', 'audit_engine.kpsc.parser_scripts.parse_spaghetti_sheet', 'audit_engine.kpsc.parser_scripts.parse_spaghetti_problems']

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

def kpsc_run_validations_run_parsers(xlsx_path: Path, output_dir: Path) -> Dict[str, Any]:
    """
    Запускает все 9 парсеров.

    Каждый парсер вызывается через import → parse(xlsx_path, output_dir).
    Ошибки отдельных парсеров не останавливают остальные.

    Возвращает: {module_name: {"status": "ok"/"error", "error": "..."}}
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    results = {}
    for module_name in kpsc_run_validations_PARSER_MODULES:
        short_name = module_name.rsplit('.', 1)[-1]
        try:
            mod = importlib.import_module(module_name)
            mod.parse(xlsx_path, output_dir)
            results[short_name] = {'status': 'ok'}
            print(f'  [OK] {short_name}')
        except Exception as e:
            results[short_name] = {'status': 'error', 'error': str(e)}
            print(f'  [ERR] {short_name}: {e}')
    return results

def kpsc_run_validations_run_single_validator(rule: Dict, parser_outputs_dir: Path, output_dir: Path, rules_path: Path) -> Tuple[Dict, Dict, float]:
    """
    Запускает один валидатор через импорт.

    Валидаторы ожидают:
    - OPENAI_API_KEY в env
    - VALIDATION_RULES_PATH в env (путь к rules JSON)
    - --parser-outputs и --output как аргументы CLI

    Вместо subprocess мы подменяем sys.argv и вызываем main().
    Но безопаснее: каждый валидатор имеет load_rule / load_data / extract / build_prompt / call_llm.
    Мы вызываем их main() через subprocess для изоляции (argparse конфликтов).
    """
    import subprocess
    import sys
    rule_index = rule['rule_index']
    start = time.time()
    try:
        prefix = rule_index.replace('.', '_')
        scripts_dir = Path(__file__).parent / 'validation_scripts'
        pattern = f'validate_{prefix}_*.py'
        matches = list(scripts_dir.glob(pattern))
        if not matches:
            return (rule, {'rule_index': rule_index, 'status': 'MISSING', 'discrepancy': f'Валидатор не найден (паттерн: {pattern})'}, 0.0)
        script_path = matches[0]
        output_file = output_dir / f'validate_{prefix}.json'
        env = os.environ.copy()
        env['VALIDATION_RULES_PATH'] = str(rules_path)
        result = subprocess.run([sys.executable, str(script_path), '--parser-outputs', str(parser_outputs_dir), '--output', str(output_file)], capture_output=True, text=True, timeout=300, env=env)
        duration = time.time() - start
        if output_file.exists():
            with open(output_file, 'r', encoding='utf-8') as f:
                result_data = json.load(f)
        else:
            result_data = {'rule_index': rule_index, 'status': 'ERROR', 'discrepancy': f'Выходной файл не создан. STDERR: {result.stderr[:300]}'}
        return (rule, result_data, duration)
    except Exception as e:
        duration = time.time() - start
        return (rule, {'rule_index': rule_index, 'status': 'ERROR', 'discrepancy': f'Исключение: {str(e)}'}, duration)

def kpsc_run_validations_run_validators_parallel(rules: List[Dict], parser_outputs_dir: Path, output_dir: Path, rules_path: Path, max_workers: int=5) -> List[Tuple[Dict, Dict, float]]:
    """Запуск всех валидаторов параллельно."""
    output_dir.mkdir(parents=True, exist_ok=True)
    results = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {}
        for rule in rules:
            future = executor.submit(kpsc_run_validations_run_single_validator, rule, parser_outputs_dir, output_dir, rules_path)
            futures[future] = rule
        for future in as_completed(futures):
            rule = futures[future]
            try:
                rule_data, result_data, duration = future.result()
                results.append((rule_data, result_data, duration))
                status = result_data.get('status', '?')
                emoji = '+' if status == 'PASS' else '-' if status == 'FAIL' else '!'
                idx = result_data.get('rule_index', '?')
                print(f'  [{emoji}] [{len(results)}/{len(rules)}] {idx}: {status} ({duration:.1f}s)')
            except Exception as e:
                results.append((rule, {'rule_index': rule['rule_index'], 'status': 'ERROR', 'discrepancy': f'Future exception: {str(e)}'}, 0.0))
    return results

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

kpsc_run_validations_module = SimpleNamespace(PARSER_MODULES=kpsc_run_validations_PARSER_MODULES, load_rules=kpsc_run_validations_load_rules, get_validator_module_name=kpsc_run_validations_get_validator_module_name, run_parsers=kpsc_run_validations_run_parsers, run_single_validator=kpsc_run_validations_run_single_validator, run_validators_parallel=kpsc_run_validations_run_validators_parallel, create_excel_report=kpsc_run_validations_create_excel_report)

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
kartochka_proekta_run_validations_PARSER_MODULES = ['audit_engine.kartochka_proekta.parser_scripts.parse_kartochka_main', 'audit_engine.kartochka_proekta.parser_scripts.parse_metodika', 'audit_engine.kartochka_proekta.parser_scripts.parse_dropdown']

def kartochka_proekta_run_validations_load_rules(rules_path: Path) -> List[Dict]:
    """Загрузка правил из validation_rules.json."""
    with open(rules_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data['rules']

def kartochka_proekta_run_validations_run_parsers(xlsx_path: Path, output_dir: Path) -> Dict[str, Any]:
    """
    Запускает все 3 парсера.

    Каждый парсер: import → parse(xlsx_path, output_dir).
    Ошибки отдельных парсеров не останавливают остальные.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    results = {}
    for module_name in kartochka_proekta_run_validations_PARSER_MODULES:
        short_name = module_name.rsplit('.', 1)[-1]
        try:
            mod = importlib.import_module(module_name)
            mod.parse(xlsx_path, output_dir)
            results[short_name] = {'status': 'ok'}
            print(f'  [OK] {short_name}')
        except Exception as e:
            results[short_name] = {'status': 'error', 'error': str(e)}
            print(f'  [ERR] {short_name}: {e}')
    return results

def kartochka_proekta_run_validations_run_single_validator(rule: Dict, parser_outputs_dir: Path, output_dir: Path, rules_path: Path) -> Tuple[Dict, Dict, float]:
    """
    Запускает один валидатор как subprocess.

    Валидатор — скрипт validate_N_*.py в validation_scripts/.
    Принимает --parser-outputs и --output, пишет JSON-результат.
    """
    rule_index = rule['rule_index']
    start = time.time()
    try:
        prefix = rule_index.replace('.', '_')
        scripts_dir = Path(__file__).parent / 'validation_scripts'
        pattern = f'validate_{prefix}_*.py'
        matches = list(scripts_dir.glob(pattern))
        if not matches:
            return (rule, {'rule_index': rule_index, 'status': 'MISSING', 'discrepancy': f'Валидатор не найден (паттерн: {pattern})'}, 0.0)
        script_path = matches[0]
        output_file = output_dir / f'validate_{prefix}.json'
        env = os.environ.copy()
        env['VALIDATION_RULES_PATH'] = str(rules_path)
        result = subprocess.run([sys.executable, str(script_path), '--parser-outputs', str(parser_outputs_dir), '--output', str(output_file)], capture_output=True, text=True, timeout=300, env=env)
        duration = time.time() - start
        if output_file.exists():
            with open(output_file, 'r', encoding='utf-8') as f:
                result_data = json.load(f)
        else:
            result_data = {'rule_index': rule_index, 'status': 'ERROR', 'discrepancy': f'Выходной файл не создан. STDERR: {result.stderr[:500]}'}
        return (rule, result_data, duration)
    except Exception as e:
        duration = time.time() - start
        return (rule, {'rule_index': rule_index, 'status': 'ERROR', 'discrepancy': f'Исключение: {str(e)}'}, duration)

def kartochka_proekta_run_validations_run_validators_parallel(rules: List[Dict], parser_outputs_dir: Path, output_dir: Path, rules_path: Path, max_workers: int=5) -> List[Tuple[Dict, Dict, float]]:
    """Запуск всех валидаторов параллельно через ThreadPoolExecutor."""
    output_dir.mkdir(parents=True, exist_ok=True)
    results = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {}
        for rule in rules:
            future = executor.submit(kartochka_proekta_run_validations_run_single_validator, rule, parser_outputs_dir, output_dir, rules_path)
            futures[future] = rule
        for future in as_completed(futures):
            rule = futures[future]
            try:
                rule_data, result_data, duration = future.result()
                results.append((rule_data, result_data, duration))
                status = result_data.get('status', '?')
                emoji = '+' if status == 'PASS' else '-' if status == 'FAIL' else '!'
                idx = result_data.get('rule_index', '?')
                print(f'  [{emoji}] [{len(results)}/{len(rules)}] {idx}: {status} ({duration:.1f}s)')
            except Exception as e:
                results.append((rule, {'rule_index': rule['rule_index'], 'status': 'ERROR', 'discrepancy': f'Future exception: {str(e)}'}, 0.0))
    return results

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

kartochka_proekta_run_validations_module = SimpleNamespace(PARSER_MODULES=kartochka_proekta_run_validations_PARSER_MODULES, load_rules=kartochka_proekta_run_validations_load_rules, run_parsers=kartochka_proekta_run_validations_run_parsers, run_single_validator=kartochka_proekta_run_validations_run_single_validator, run_validators_parallel=kartochka_proekta_run_validations_run_validators_parallel, create_excel_report=kartochka_proekta_run_validations_create_excel_report)

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

# START_SOURCE_PLAN_GRAFIK
# PURPOSE: Inlined source from audit_engine/plan_grafik/__init__.py.
plan_grafik_logger = logging.getLogger(__name__)

def plan_grafik_run(args) -> AuditResult:
    """
    Запуск аудита план-графика.

    Args:
        args: Namespace с полями target, session_dir, parse_only, rule_filter

    Returns:
        AuditResult с нарушениями
    """
    start_time = time.time()
    target_path = args.target
    session_dir = Path(args.session_dir) if args.session_dir else Path(f'logs_result/plan_grafik/session_{datetime.now().strftime('%Y%m%d_%H%M%S')}')
    session_dir.mkdir(parents=True, exist_ok=True)
    print(f'[plan_grafik] Старт аудита: {target_path}')
    print(f'[plan_grafik] Сессия: {session_dir}')
    try:
        parsed = parse_plan_grafik(target_path)
    except Exception as e:
        error_msg = f'Ошибка парсинга: {e}'
        print(f'[plan_grafik] {error_msg}', file=sys.stderr)
        (session_dir / 'ERROR.txt').write_text(error_msg, encoding='utf-8')
        return AuditResult(violations=[], doc_type='plan_grafik', session_dir=session_dir, duration_sec=time.time() - start_time, rules_checked=0, target_path=target_path)
    parsed_path = session_dir / 'parsed.json'
    with open(parsed_path, 'w', encoding='utf-8') as f:
        json.dump(parsed, f, ensure_ascii=False, indent=2, default=str)
    print(f'[plan_grafik] Парсинг сохранён: {parsed_path}')
    if getattr(args, 'parse_only', False):
        return AuditResult(doc_type='plan_grafik', session_dir=session_dir, duration_sec=time.time() - start_time, target_path=target_path)
    violations = run_all_validators(parsed, target_path)
    xlsx_path = str(session_dir / 'validation_report.xlsx')
    save_to_excel(violations, xlsx_path)
    print(f'[plan_grafik] Отчёт: {xlsx_path}')
    duration = time.time() - start_time
    print(f'[plan_grafik] Завершён за {duration:.1f} сек. Нарушений: {len(violations)}')
    return AuditResult(violations=violations, doc_type='plan_grafik', session_dir=session_dir, duration_sec=duration, rules_checked=9, target_path=target_path)

plan_grafik_module = SimpleNamespace(logger=plan_grafik_logger, run=plan_grafik_run)

# END_SOURCE_PLAN_GRAFIK

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

api_server_module = SimpleNamespace(app=app, _audit_lock=_audit_lock, _queue_counter=_queue_counter, _queue_counter_lock=_queue_counter_lock, AUTH_LOGIN=AUTH_LOGIN, AUTH_PASSWORD=AUTH_PASSWORD, AUTH_TOKEN=AUTH_TOKEN, UPLOAD_DIR=UPLOAD_DIR, sessions=sessions, LoginRequest=LoginRequest, _check_auth=_check_auth, health_check=health_check, login=login, get_types=get_types, start_audit=start_audit, audit_events=audit_events, get_result=get_result, download_report=download_report, _detect_engine=api_server__detect_engine, _STANDARD_EXTENSIONS=_STANDARD_EXTENSIONS, _XLSX_EXTENSIONS=_XLSX_EXTENSIONS, _PPTX_EXTENSIONS=_PPTX_EXTENSIONS, _get_allowed_extensions=_get_allowed_extensions, SPECIAL_ENGINES=api_server_SPECIAL_ENGINES, _run_special_engine=_run_special_engine, _run_audit_thread=_run_audit_thread)

# END_SOURCE_API_SERVER
# END_API

# START_CLI
# PURPOSE: Provide the repository CLI for document audit execution and doc-type discovery.
# INPUTS: argparse flags, doc_type, target path, optional overrides.
# OUTPUTS: terminal progress, session artifacts, process exit status.
# KEYWORDS: cli, argparse, audit, list-types.
# LINKS: run_audit.py.
# RATIONALE: `python main.py ...` must become the only active CLI path.

# START_SOURCE_RUN_AUDIT
# PURPOSE: Inlined source from run_audit.py.
run_audit_SPECIAL_ENGINES = {'drivers': 'audit_engine.drivers', 'kpsc': 'audit_engine.kpsc', 'kartochka_proekta': 'audit_engine.kartochka_proekta'}

def run_audit__detect_engine(doc_type: str) -> str:
    """
    Определяет тип движка по config.json.

    Если в config.json есть поле "engine" — возвращает его значение.
    Иначе возвращает "vision" (стандартный pipeline).
    """
    config_path = Path(__file__).parent / 'doc_configs' / doc_type / 'config.json'
    if config_path.exists():
        with open(config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data.get('engine', 'vision')
    return 'vision'

def main():
    """Точка входа CLI."""
    parser = argparse.ArgumentParser(description='Единый аудит документов — Vision Pipeline + LLM')
    parser.add_argument('--doc-type', default=None, help='Тип документа (имя папки в doc_configs/)')
    parser.add_argument('--target', default=None, help='Путь к целевому документу')
    parser.add_argument('--model', default=None, help='Модель OpenAI (переопределяет конфиг)')
    parser.add_argument('--temperature', type=float, default=None, help='Температура генерации')
    parser.add_argument('--rule-filter', default=None, help='Проверить только указанное правило (номер: 3 для Vision, 1.1 для КПСЦ)')
    parser.add_argument('--parse-only', action='store_true', help='Только Vision-парсинг, без проверки правил')
    parser.add_argument('--print-prompts', action='store_true', help='Режим отладки: только промпты без вызова LLM')
    parser.add_argument('--session-dir', default=None, help='Директория для логов сессии')
    parser.add_argument('--chunk-filter', default=None, help='Парсить только указанный чанк')
    parser.add_argument('--out-xlsx', default=None, help='Путь для сохранения Excel')
    parser.add_argument('--secondary', default=None, help='Путь к вторичному файлу (XLSX для multi-file аудитов)')
    parser.add_argument('--list-types', action='store_true', help='Показать список доступных типов документов')
    args = parser.parse_args()
    if args.list_types:
        doc_types = AuditEngine.list_doc_types()
        if not doc_types:
            print('Нет доступных типов документов в doc_configs/')
            print('Создайте папку doc_configs/<doc_type>/ с config.json, rules.json и chunks_vision.json')
            sys.exit(0)
        print(f'\n📋 Доступные типы документов ({len(doc_types)}):')
        print(f'{'─' * 50}')
        for dt in doc_types:
            title = f' — {dt['doc_title']}' if dt['doc_title'] else ''
            print(f'  {dt['doc_type']}{title}')
        print(f'\nИспользование: python run_audit.py --doc-type <doc_type> --target <file>')
        sys.exit(0)
    if not args.doc_type:
        parser.error('--doc-type обязателен (или используйте --list-types)')
    if not args.target:
        parser.error('--target обязателен')
    engine_type = run_audit__detect_engine(args.doc_type)
    if engine_type in run_audit_SPECIAL_ENGINES:
        module = importlib.import_module(run_audit_SPECIAL_ENGINES[engine_type])
        result = module.run(args)
    else:
        rule_filter = int(args.rule_filter) if args.rule_filter is not None else None
        engine = AuditEngine(args.doc_type)
        result = engine.run(target_path=args.target, model=args.model, temperature=args.temperature, rule_filter=rule_filter, parse_only=args.parse_only, print_prompts=args.print_prompts, session_dir=args.session_dir, chunk_filter=args.chunk_filter, out_xlsx=args.out_xlsx, secondary_path=args.secondary)
    sys.exit(1 if result.violations else 0)

run_audit_module = SimpleNamespace(SPECIAL_ENGINES=run_audit_SPECIAL_ENGINES, _detect_engine=run_audit__detect_engine, main=main)

# END_SOURCE_RUN_AUDIT
# END_CLI

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

KPSC_VALIDATOR_MODULE_DISPATCH = {
    "1.1": kpsc_validate_1_1_kpsc_text_module,
    "1.2": kpsc_validate_1_2_company_name_module,
    "1.3": kpsc_validate_1_3_flow_name_module,
    "1.4": kpsc_validate_1_4_responsible_module,
    "1.5": kpsc_validate_1_5_date_developed_module,
    "1.6": kpsc_validate_1_6_date_implementation_module,
    "1.7": kpsc_validate_1_7_compiled_by_module,
    "2.1": kpsc_validate_2_1_problems_count_module,
    "2.2": kpsc_validate_2_2_problems_sequence_module,
    "2.3": kpsc_validate_2_3_problems_description_module,
    "4.1": kpsc_validate_4_1_vpp_filled_module,
    "4.2": kpsc_validate_4_2_vpp_sum_module,
    "5.1": kpsc_validate_5_1_units_module,
    "6.1": kpsc_validate_6_1_transport_row_module,
    "7.1": kpsc_validate_7_1_sheet_kpsc_module,
    "7.2": kpsc_validate_7_2_sheet_legend_module,
    "7.3": kpsc_validate_7_3_sheet_pokazateli_module,
    "7.4": kpsc_validate_7_4_sheet_ocifrovka_module,
    "7.5": kpsc_validate_7_5_sheet_pa1_module,
    "7.6": kpsc_validate_7_6_sheet_spaghetti_module,
    "7.7": kpsc_validate_7_7_sheet_spaghetti_problems_module,
    "7.8": kpsc_validate_7_8_sheet_takt_time_module,
    "8.1": kpsc_validate_8_1_units_cross_check_module,
    "8.2": kpsc_validate_8_2_values_cross_check_module,
    "8.3": kpsc_validate_8_3_indicators_cross_check_module,
}

KARTOCHKA_VALIDATOR_MODULE_DISPATCH = {
    "1": kartochka_proekta_validate_1_filename_module,
    "2": kartochka_proekta_validate_2_org_name_module,
    "3": kartochka_proekta_validate_3_flow_name_module,
    "4": kartochka_proekta_validate_4_signee_module,
    "5": kartochka_proekta_validate_5_required_fields_module,
    "6": kartochka_proekta_validate_6_justification_module,
    "7": kartochka_proekta_validate_7_event_dates_module,
    "8": kartochka_proekta_validate_8_indicators_module,
    "9": kartochka_proekta_validate_9_units_kartochka_module,
    "10": kartochka_proekta_validate_10_units_metodika_module,
    "11": kartochka_proekta_validate_11_calc_method_module,
}

kpsc_run_validations_PARSER_MODULES = list(KPSC_PARSER_MODULE_DISPATCH)
kartochka_proekta_run_validations_PARSER_MODULES = list(KARTOCHKA_PARSER_MODULE_DISPATCH)


def get_parser(name: str) -> Callable[[str], Dict[str, str]]:
    if name not in SECONDARY_PARSER_DISPATCH:
        raise KeyError(f"Парсер '{name}' не зарегистрирован. Доступные: {sorted(SECONDARY_PARSER_DISPATCH)}")
    return SECONDARY_PARSER_DISPATCH[name]


def _run_kpsc_validator_module(module_ns: SimpleNamespace, parser_outputs_dir: Path, output_file: Path) -> Dict[str, Any]:
    api_key = os.environ.get("OPENAI_API_KEY", "dummy")
    rule = module_ns.load_rule(module_ns.RULE_INDEX)
    data = module_ns.load_data(parser_outputs_dir)
    extracted = module_ns.extract_relevant_data(data)
    prompt = module_ns.build_prompt(extracted, rule)
    result = module_ns.call_llm(prompt, api_key)
    module_ns.save_result(result, output_file)
    return result


def _run_kartochka_validator_module(module_ns: SimpleNamespace, parser_outputs_dir: Path, output_file: Path) -> Dict[str, Any]:
    import inspect

    rule = module_ns.load_rule(module_ns.RULE_INDEX)
    loaded_data = module_ns.load_data(parser_outputs_dir)
    pending_data = list(loaded_data) if isinstance(loaded_data, tuple) else [loaded_data]
    call_args: List[Any] = []
    api_key = os.environ.get("OPENAI_API_KEY", "dummy")

    for param in inspect.signature(module_ns.validate).parameters.values():
        if param.name == "rule":
            call_args.append(rule)
        elif param.name == "api_key":
            call_args.append(api_key)
        else:
            if not pending_data:
                raise ValueError(f"Недостаточно данных для валидатора {module_ns.RULE_INDEX}")
            call_args.append(pending_data.pop(0))

    result = module_ns.validate(*call_args)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return result


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
        result_data = _run_kpsc_validator_module(module_ns, parser_outputs_dir, output_file)
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
        result_data = _run_kartochka_validator_module(module_ns, parser_outputs_dir, output_file)
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


def _default_prompt_path() -> Path:
    return DOC_CONFIGS_DIR / "drivers" / "driver_check_prompt.txt"


def drivers__load_config() -> Dict[str, Any]:
    with open(DOC_CONFIGS_DIR / "drivers" / "config.json", "r", encoding="utf-8") as f:
        return json.load(f)


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


def run_drivers_special(args) -> AuditResult:
    start_time = time.time()
    config = drivers__load_config()
    model = args.model or config.get("model", "gpt-4.1-mini")
    temperature = args.temperature if args.temperature is not None else config.get("temperature", 0.0)
    primary_threshold = config.get("primary_threshold", 9.0)
    fallback_threshold = config.get("fallback_threshold", 7.0)
    session_dir = Path(args.session_dir) if args.session_dir else LOGS_RESULT_DIR / "drivers" / f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    session_dir.mkdir(parents=True, exist_ok=True)
    return drivers__run_pipeline(args, config, model, temperature, primary_threshold, fallback_threshold, session_dir, start_time)


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


api_server_module = SimpleNamespace(
    app=app,
    AUTH_LOGIN=AUTH_LOGIN,
    AUTH_PASSWORD=AUTH_PASSWORD,
    AUTH_TOKEN=AUTH_TOKEN,
    UPLOAD_DIR=UPLOAD_DIR,
    sessions=sessions,
    LoginRequest=LoginRequest,
    _check_auth=_check_auth,
    health_check=health_check,
    login=login,
    get_types=get_types,
    start_audit=start_audit,
    audit_events=audit_events,
    get_result=get_result,
    download_report=download_report,
    _detect_engine=api_server__detect_engine,
    SPECIAL_ENGINES=api_server_SPECIAL_ENGINES,
    _run_special_engine=_run_special_engine,
    _run_audit_thread=_run_audit_thread,
)


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


run_audit_module = SimpleNamespace(
    SPECIAL_ENGINES=run_audit_SPECIAL_ENGINES,
    _detect_engine=run_audit__detect_engine,
    main=main,
)


build_kpsc_header_payload = kpsc_parse_kpsc_header_build_payload
run_kpsc_parsers = kpsc_run_validations_run_parsers

# END_RUNTIME_INTEGRATION

if __name__ == "__main__":
    main()
