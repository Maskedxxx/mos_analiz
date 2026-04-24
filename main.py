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
SYSTEM_PROMPTS_DIR = PROJECT_ROOT / 'audit_engine' / 'system_prompts'

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
        model — модель OpenAI для LLM-проверок
        system_prompt — имя файла системного промпта (из system_prompts/)
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
        config_dir — путь к папке с конфигами (автоматически)
        rules_path — путь к rules.json (автоматически)
        chunks_vision_path — путь к chunks_vision.json (автоматически)
        template_path — путь к файлу шаблона (автоматически)
        template_cached_path — путь к кэшу Vision-парсинга шаблона
    """
    doc_type: str
    doc_title: str = ''
    model: str = 'gpt-4.1-mini'
    system_prompt: str = 'default.txt'
    filename_pattern: str = ''
    filename_keywords: Optional[List[str]] = None
    max_workers: int = 1
    temperature: float = 0.0
    secondary_file: Optional[SecondaryFileConfig] = None
    # Выбор парсера по расширению файла для generic-пути. Для special-движков — пусто.
    parser_by_ext: Dict[str, str] = field(default_factory=dict)
    # Имя special-движка (kpsc, kartochka_proekta, ...). Взаимоисключающе с parser_by_ext.
    engine: Optional[str] = None
    ocr_model: Optional[str] = None
    ocr_base_url: str = 'http://localhost:8010/v1/'
    ocr_prompt: str = '提取文档图片中正文的所有信息用markdown格式表示，忽略页眉页脚。表格用html格式表达，公式用LaTeX格式表示，按照阅读顺序组织进行解析。特别注意：保留表格上方和下方的所有独立标题行和文本，不要将标题合并到表格中。'
    ocr_dpi: int = 200
    llm_base_url: Optional[str] = 'http://localhost:8001/v1/'
    llm_max_tokens: int = 4096
    reasoning_effort: Optional[str] = None
    llm_seed: Optional[int] = None
    paddle_layout_model: Optional[str] = None
    paddle_layout_device: str = 'cuda:0'
    paddle_layout_base_url: Optional[str] = None
    paddle_vlm_model: Optional[str] = None
    strip_annotations_chunks: List[str] = field(default_factory=list)
    engine_mode: str = 'legacy'
    config_dir: Path = field(default_factory=Path)
    rules_path: Path = field(default_factory=Path)
    chunks_vision_path: Path = field(default_factory=Path)
    template_path: Optional[Path] = None
    template_cached_path: Optional[Path] = None

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

    Ищет config.json, rules.json, chunks_vision.json, template/*.

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
    config = AuditConfig(doc_type=data['doc_type'], doc_title=data.get('doc_title', data['doc_type']), model=data.get('model', 'gpt-4.1-mini'), system_prompt=data.get('system_prompt', 'default.txt'), filename_pattern=data.get('filename_pattern', ''), filename_keywords=data.get('filename_keywords', None), max_workers=data.get('max_workers', 1), temperature=data.get('temperature', 0.0), parser_by_ext=parser_by_ext if has_parser_map else {}, engine=engine if has_engine else None, ocr_model=data.get('ocr_model', None), ocr_base_url=data.get('ocr_base_url', 'http://localhost:8010/v1/'), ocr_prompt=data.get('ocr_prompt', AuditConfig.ocr_prompt), ocr_dpi=data.get('ocr_dpi', 200), paddle_layout_model=data.get('paddle_layout_model', None), paddle_layout_device=data.get('paddle_layout_device', 'cuda:0'), paddle_layout_base_url=data.get('paddle_layout_base_url', None), paddle_vlm_model=data.get('paddle_vlm_model', None), llm_base_url=data.get('llm_base_url', 'http://localhost:8001/v1/'), llm_max_tokens=data.get('llm_max_tokens', 4096), reasoning_effort=data.get('reasoning_effort', None), llm_seed=data.get('llm_seed', None), strip_annotations_chunks=data.get('strip_annotations_chunks', []), engine_mode=data.get('engine_mode', 'legacy'), config_dir=config_dir, rules_path=config_dir / 'rules.json', chunks_vision_path=config_dir / 'chunks_vision.json')
    if 'secondary_file' in data:
        sf = data['secondary_file']
        config.secondary_file = SecondaryFileConfig(type=sf['type'], parser=sf['parser'], chunk_prefix=sf.get('chunk_prefix', 'xlsx_'))
    template_dir = config_dir / 'template'
    if template_dir.exists():
        for ext in ['*.docx', '*.pptx']:
            templates = list(template_dir.glob(ext))
            if templates:
                config.template_path = templates[0]
                break
        cached = template_dir / 'template_cached.json'
        if cached.exists():
            config.template_cached_path = cached
    return config

models_module = SimpleNamespace(RuleSpec=RuleSpec, SecondaryFileConfig=SecondaryFileConfig, AuditConfig=AuditConfig, AuditResult=AuditResult, load_rules=load_rules, load_audit_config=load_audit_config)

# END_SOURCE_MODELS

# START_SOURCE_VISION_PARSER_CONFIG
# PURPOSE: Inlined source from audit_engine/vision_parser/config.py.
@dataclass
class ChunkConfig:
    """
    Конфигурация одного чанка документа.

    Attributes:
        name: Имя чанка (ключ в результирующем словаре)
        pages: Страницы для обработки:
               - List[int]: конкретные номера страниц [1, 2, 3]
               - str "all": все страницы
               - str "1-3": диапазон страниц
               - List с отрицательными: [-1] = последняя страница
        prompt: Промпт для Vision LLM с инструкциями по извлечению
    """
    name: str
    pages: Union[List[int], str]
    prompt: str

    def get_page_indices(self, total_pages: int) -> List[int]:
        """
        Преобразует pages в список индексов страниц (0-based).

        Args:
            total_pages: Общее количество страниц в документе

        Returns:
            Список индексов страниц (0-based)
        """
        if self.pages == 'all':
            return list(range(total_pages))
        if isinstance(self.pages, str):
            if '-' in self.pages:
                parts = self.pages.split('-')
                start = int(parts[0]) - 1
                end = int(parts[1])
                return list(range(start, end))
            else:
                return [int(self.pages) - 1]
        if isinstance(self.pages, list):
            result = []
            for p in self.pages:
                if p < 0:
                    result.append(total_pages + p)
                else:
                    result.append(p - 1)
            return result
        return []

@dataclass
class VisionConfig:
    """
    Конфигурация Vision Pipeline.

    Attributes:
        model: Модель OpenAI для Vision API
        dpi: Разрешение для рендеринга PDF (200 = хороший баланс качества/размера)
        max_workers: Максимум параллельных запросов к API
        timeout: Таймаут одного запроса в секундах
        max_retries: Максимум повторных попыток при ошибке
        detail: Уровень детализации для Vision API ("low", "high", "auto")
    """
    model: str = 'gpt-4.1-mini'
    dpi: int = 200
    max_workers: int = 4
    timeout: int = 60
    max_retries: int = 3
    detail: str = 'high'

def load_chunks_config(config_path: str) -> List[ChunkConfig]:
    """
    Загружает конфигурацию чанков из JSON-файла.

    Args:
        config_path: Путь к JSON-файлу с конфигурацией

    Returns:
        Список ChunkConfig

    Raises:
        FileNotFoundError: Если файл не найден
        json.JSONDecodeError: Если JSON невалидный
    """
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f'Конфигурация чанков не найдена: {config_path}')
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    chunks = []
    for chunk_data in data.get('chunks', []):
        chunk = ChunkConfig(name=chunk_data['name'], pages=chunk_data['pages'], prompt=chunk_data['prompt'])
        chunks.append(chunk)
    return chunks

def load_vision_config(config_path: str=None) -> VisionConfig:
    """
    Загружает конфигурацию Vision Pipeline.

    Args:
        config_path: Путь к JSON-файлу (опционально).
                    Если не указан, возвращает дефолтную конфигурацию.

    Returns:
        VisionConfig с параметрами
    """
    if config_path is None:
        return VisionConfig()
    path = Path(config_path)
    if not path.exists():
        return VisionConfig()
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    vision_data = data.get('vision', {})
    return VisionConfig(model=vision_data.get('model', 'gpt-4.1-mini'), dpi=vision_data.get('dpi', 200), max_workers=vision_data.get('max_workers', 4), timeout=vision_data.get('timeout', 60), max_retries=vision_data.get('max_retries', 3), detail=vision_data.get('detail', 'high'))

vision_parser_config_module = SimpleNamespace(ChunkConfig=ChunkConfig, VisionConfig=VisionConfig, load_chunks_config=load_chunks_config, load_vision_config=load_vision_config)

# END_SOURCE_VISION_PARSER_CONFIG

# START_SOURCE_PADDLE_PARSER_CONFIG
# PURPOSE: Inlined source from audit_engine/paddle_parser/config.py.
LAYOUT_MODEL_REPO = 'docling-project/docling-layout-heron-101'

LAYOUT_CONF_THRESHOLD = 0.3

LAYOUT_POST_CONF_MIN = 0.1

LAYOUT_DEDUP_IOU = 0.9

LAYOUT_CONTAINMENT_THR = 0.7

CROP_PADDING_PX = 15

VLM_API_KEY = 'none'

VLM_MAX_TOKENS = 2000

VLM_TEMPERATURE = 0.0

CLASS_PROMPTS = {'Text': 'OCR:', 'Title': 'OCR:', 'Section-header': 'OCR:', 'Caption': 'OCR:', 'Footnote': 'OCR:', 'List-item': 'OCR:', 'Page-header': 'OCR:', 'Page-footer': 'OCR:', 'Document Index': 'OCR:', 'Code': 'OCR:', 'Form': 'OCR:', 'Key-Value Region': 'OCR:', 'Checkbox-Selected': 'OCR:', 'Checkbox-Unselected': 'OCR:', 'Table': 'Table Recognition:', 'Formula': 'Formula Recognition:', 'Picture': 'OCR:'}

SMALL_ELEMENT_MIN_AREA_RATIO = 0.001

SMALL_ELEMENT_CLASSES = {'Page-header', 'Page-footer'}

CLASS_MD_WRAPPER = {'Text': '{text}', 'Title': '## {text}', 'Section-header': '### {text}', 'Caption': '*{text}*', 'Footnote': '*{text}*', 'List-item': '{text}', 'Page-header': '{text}', 'Page-footer': '{text}', 'Document Index': '{text}', 'Code': '```\n{text}\n```', 'Form': '{text}', 'Key-Value Region': '{text}', 'Checkbox-Selected': '[x] {text}', 'Checkbox-Unselected': '[ ] {text}', 'Table': '{text}', 'Formula': '$$\n{text}\n$$', 'Picture': '{text}'}

TABLE_OUTPUT_FORMAT = 'html'

TABLE_MASK_PADDING_PX = 5

paddle_parser_config_module = SimpleNamespace(LAYOUT_MODEL_REPO=LAYOUT_MODEL_REPO, LAYOUT_CONF_THRESHOLD=LAYOUT_CONF_THRESHOLD, LAYOUT_POST_CONF_MIN=LAYOUT_POST_CONF_MIN, LAYOUT_DEDUP_IOU=LAYOUT_DEDUP_IOU, LAYOUT_CONTAINMENT_THR=LAYOUT_CONTAINMENT_THR, CROP_PADDING_PX=CROP_PADDING_PX, VLM_API_KEY=VLM_API_KEY, VLM_MAX_TOKENS=VLM_MAX_TOKENS, VLM_TEMPERATURE=VLM_TEMPERATURE, CLASS_PROMPTS=CLASS_PROMPTS, SMALL_ELEMENT_MIN_AREA_RATIO=SMALL_ELEMENT_MIN_AREA_RATIO, SMALL_ELEMENT_CLASSES=SMALL_ELEMENT_CLASSES, CLASS_MD_WRAPPER=CLASS_MD_WRAPPER, TABLE_OUTPUT_FORMAT=TABLE_OUTPUT_FORMAT, TABLE_MASK_PADDING_PX=TABLE_MASK_PADDING_PX)

# END_SOURCE_PADDLE_PARSER_CONFIG
# END_MODELS

# START_PREPROCESSORS
# PURPOSE: Normalize document chunks before LLM comparison or semantic filtering.
# INPUTS: doc_type, scope name, raw chunk text.
# OUTPUTS: Preprocessed text registered by doc_type and scope.
# KEYWORDS: preprocessors, registry, normalization.
# LINKS: audit_engine/preprocessors/*.py.
# RATIONALE: Preprocessors remain data-shaping business logic even inside one file.

# START_SOURCE_PREPROCESSORS_REGISTRY
# PURPOSE: Inlined source from audit_engine/preprocessors/registry.py.
preprocessors_registry__REGISTRY: Dict[Tuple[str, str], Callable[[str], str]] = {}

def register_preprocessor(doc_type: str, scope: str):
    """
    Декоратор для регистрации препроцессора.

    Args:
        doc_type: тип документа
        scope: чанк, к которому применяется препроцессор

    Функция должна принимать str → str.
    """

    def decorator(fn: Callable[[str], str]):
        preprocessors_registry__REGISTRY[doc_type, scope] = fn
        return fn
    return decorator

def get_preprocessors(doc_type: str) -> Dict[str, Callable[[str], str]]:
    """
    Возвращает все препроцессоры для типа документа.

    Returns:
        Словарь {scope: preprocess_fn}
    """
    return {scope: fn for (dt, scope), fn in preprocessors_registry__REGISTRY.items() if dt == doc_type}

preprocessors_registry_module = SimpleNamespace(_REGISTRY=preprocessors_registry__REGISTRY, register_preprocessor=register_preprocessor, get_preprocessors=get_preprocessors)

# END_SOURCE_PREPROCESSORS_REGISTRY

# START_SOURCE_PREPROCESSORS_CHEKLIST
# PURPOSE: Inlined source from audit_engine/preprocessors/cheklist.py.
def preprocessors_cheklist_normalize_text_for_comparison(text: str) -> str:
    """
    Нормализует текст перед сравнением, заменяя плейсхолдеры на унифицированные метки.
    Применяется и к целевому документу и к шаблону.
    """
    text = re.sub('\\s+-\\s+(заголовок|дата|номер|форма|место|подпись).*$', '', text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub('[ \\t]+', ' ', text)
    text = re.sub('\\n\\s*\\n', '\n\n', text)
    text = re.sub('\\d{1,2}\\.\\d{1,2}\\.\\d{4}\\s*г?\\.?', '[ДАТА]', text)
    text = re.sub('«?\\d{1,2}»?\\s*[а-яё]+\\s*\\d{4}\\s*г?\\.?', '[ДАТА]', text, flags=re.IGNORECASE)
    text = re.sub('_{2,}\\.\\s*_{2,}\\.\\s*\\d{4}|_{2,}\\.\\s*_{2,}\\.\\s*202_?', '[ДАТА]', text)
    text = re.sub('[А-ЯЁ][а-яё]+\\s+[А-ЯЁ]\\.[А-ЯЁ]\\.', '[ФИО]', text)
    text = re.sub('[А-ЯЁ]\\.[А-ЯЁ]\\.\\s*[А-ЯЁ][а-яё]+', '[ФИО]', text)
    text = re.sub('И\\.О\\.\\s*Фамилия', '[ФИО]', text)
    text = re.sub('\\(Фамилия И\\.О\\.\\)', '', text)
    text = re.sub('(ООО|ЗАО|АО|ПАО)\\s*[«"][\\w\\s]+[»"]', '[ОРГАНИЗАЦИЯ]', text)
    text = re.sub('(ООО|ЗАО|АО|ПАО)\\s*[«"]_+[»"]', '[ОРГАНИЗАЦИЯ]', text)
    text = re.sub('№\\s*\\d+[а-яА-Я]*', '№ [НОМЕР]', text)
    text = re.sub('№\\s*_+', '№ [НОМЕР]', text)
    text = re.sub('_+', '[ПЛЕЙСХОЛДЕР]', text)
    text = re.sub('\\[ДАТА\\]\\s*\\[ДАТА\\]', '[ДАТА]', text)
    text = re.sub('\\[ФИО\\]\\s*\\[ФИО\\]', '[ФИО]', text)
    return text.strip()

def normalize_table_for_comparison(table_text: str) -> str:
    """
    Нормализует таблицу критериев для сравнения TARGET и TEMPLATE.

    1. Заменяет числовые оценки (0, 1, 2) на [ОЦЕНКА]
    2. Добавляет номера критериев если отсутствуют
    3. Убирает структурные различия
    """
    if not table_text:
        return table_text
    criteria_keywords = {'Влияние результатов работы эталонного участка': '1', 'Руководство предприятия выделяет этот участок': '2', 'На участке выявлены резервы повышения производительности': '3', 'Применение обязательных инструментов БП': '4', 'На участке есть проблемы, которые возможно исключить': '5', 'На какие потоки предприятия влияет': '6', 'Оцените потенциал тиражирования': '7'}
    lines = table_text.split('\n')
    normalized_lines = []
    for line in lines:
        if line.strip().startswith('[') or 'Итоговая оценка' in line:
            normalized_lines.append(line)
            continue
        if '|' in line:
            parts = line.split('|')
            if len(parts) > 5:
                line = '|'.join(parts[:4]) + '|'
            for keyword, num in criteria_keywords.items():
                if keyword in line:
                    pattern = '^\\|\\s*\\|\\s*(' + re.escape(keyword[:20]) + ')'
                    if re.search(pattern, line):
                        line = re.sub(pattern, '| ' + num + ' | \\1', line)
                    break
            line = re.sub('\\|\\s*([012])\\s*\\|', '| [ОЦЕНКА] |', line)
            if re.match('^\\|[\\s\\-|]+$', line):
                dashes = line.split('|')
                if len(dashes) > 5:
                    line = '|'.join(dashes[:4]) + '|'
        normalized_lines.append(line)
    return '\n'.join(normalized_lines)

@register_preprocessor('cheklist_eu', 'шапка')
def preprocess_cheklist_shapa(text: str) -> str:
    """Нормализация шапки чек-листа."""
    return preprocessors_cheklist_normalize_text_for_comparison(text)

@register_preprocessor('cheklist_eu', 'таблица_критериев')
def preprocess_cheklist_table(text: str) -> str:
    """Нормализация таблицы критериев."""
    text = preprocessors_cheklist_normalize_text_for_comparison(text)
    text = normalize_table_for_comparison(text)
    return text

@register_preprocessor('cheklist_eu', 'подвал')
def preprocess_cheklist_podval(text: str) -> str:
    """Нормализация подвала чек-листа."""
    return preprocessors_cheklist_normalize_text_for_comparison(text)

preprocessors_cheklist_module = SimpleNamespace(normalize_text_for_comparison=preprocessors_cheklist_normalize_text_for_comparison, normalize_table_for_comparison=normalize_table_for_comparison, preprocess_cheklist_shapa=preprocess_cheklist_shapa, preprocess_cheklist_table=preprocess_cheklist_table, preprocess_cheklist_podval=preprocess_cheklist_podval)

# END_SOURCE_PREPROCESSORS_CHEKLIST

# START_SOURCE_PREPROCESSORS_ITER8_NORMALIZE
# PURPOSE: Inlined source from audit_engine/preprocessors/iter8_normalize.py.
def preprocessors_iter8_normalize_normalize_text_for_comparison(text: str) -> str:
    """
    Нормализует текст перед сравнением с шаблоном.

    Заменяет переменные данные (даты, ФИО, организации) на метки [ДАТА], [ФИО] и т.д.
    Применяется и к целевому документу, и к шаблону.
    """
    text = re.sub('\\s*-\\s*[а-яё\\s]+$', '', text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub('\\s*-\\s*(заголовок|дата|номер|форма|место|подпись).*$', '', text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub('[ \\t]+', ' ', text)
    text = re.sub('\\n\\s*\\n', '\n\n', text)
    text = re.sub('\\d{1,2}\\.\\d{1,2}\\.\\d{4}\\s*г?\\.?', '[ДАТА]', text)
    text = re.sub('«?\\d{1,2}»?\\s*[а-яё]+\\s*\\d{4}\\s*г?\\.?', '[ДАТА]', text, flags=re.IGNORECASE)
    text = re.sub('_{2,}\\.\\s*_{2,}\\.\\s*\\d{4}|_{2,}\\.\\s*_{2,}\\.\\s*202_?', '[ДАТА]', text)
    text = re.sub('в срок до\\s*_+\\.?', 'в срок до [ДАТА]', text)
    text = re.sub('в срок до\\s*\\[ДАТА\\]\\.?', 'в срок до [ДАТА]', text)
    text = re.sub('срок\\s*_*\\d{1,2}\\.\\d{1,2}\\.\\d{4}\\.?', 'срок [ДАТА]', text)
    text = re.sub('срок\\s*_+\\.?', 'срок [ДАТА]', text)
    text = re.sub('срок\\s*\\[ДАТА\\]\\.?', 'срок [ДАТА]', text)
    text = re.sub('Ответственный\\s*_+', 'Ответственный [ДОЛЖНОСТЬ]', text)
    text = re.sub('Ответственный\\s+([а-яёА-ЯЁ\\s]+?)\\s+срок', 'Ответственный [ДОЛЖНОСТЬ] срок', text)
    text = re.sub('руководителем ПО\\s*_+', 'руководителем ПО [ФИО]', text)
    text = re.sub('руководителем ПО\\s+[а-яёА-ЯЁ\\s]+(?=\\s+срок)', 'руководителем ПО [ФИО]', text)
    text = re.sub('[А-ЯЁ][а-яё]+\\s+[А-ЯЁ]\\.[А-ЯЁ]\\.', '[ФИО]', text)
    text = re.sub('[А-ЯЁ]\\.[А-ЯЁ]\\.\\s*[А-ЯЁ][а-яё]+', '[ФИО]', text)
    text = re.sub('И\\.О\\.\\s*Фамилия', '[ФИО]', text)
    text = re.sub('\\(Фамилия И\\.О\\.\\)', '', text)
    text = re.sub('(ООО|ЗАО|АО|ПАО)\\s*[«"][\\w\\s]+[»"]', '[ОРГАНИЗАЦИЯ]', text)
    text = re.sub('(ООО|ЗАО|АО|ПАО)\\s*[«"]_+[»"]', '[ОРГАНИЗАЦИЯ]', text)
    text = re.sub('\\(Указать наименование должности\\)', '[ДОЛЖНОСТЬ]', text)
    text = re.sub('№\\s*\\d+[а-яА-Я]*', '№ [НОМЕР]', text)
    text = re.sub('№\\s*_+', '№ [НОМЕР]', text)
    text = re.sub('_+', '[ПЛЕЙСХОЛДЕР]', text)
    text = re.sub('\\[ДАТА\\]\\s*\\[ДАТА\\]', '[ДАТА]', text)
    text = re.sub('\\[ФИО\\]\\s*\\[ФИО\\]', '[ФИО]', text)
    text = re.sub('\\[ДОЛЖНОСТЬ\\]\\s*\\[ДОЛЖНОСТЬ\\]', '[ДОЛЖНОСТЬ]', text)
    return text.strip()

@register_preprocessor('polozhenie_po', 'заголовок')
def preprocess_polozhenie_zagolovok(text: str) -> str:
    """Нормализация заголовка для сравнения с шаблоном."""
    return preprocessors_iter8_normalize_normalize_text_for_comparison(text)

@register_preprocessor('polozhenie_po', 'основной_текст')
def preprocess_polozhenie_osnovnoy_tekst(text: str) -> str:
    """Нормализация основного текста для структурного сравнения."""
    return preprocessors_iter8_normalize_normalize_text_for_comparison(text)

@register_preprocessor('polozhenie_po', 'структура_разделов')
def preprocess_polozhenie_struktura(text: str) -> str:
    """Нормализация структуры разделов для сравнения с шаблоном."""
    return preprocessors_iter8_normalize_normalize_text_for_comparison(text)

@register_preprocessor('polozhenie_po', 'пункт_1_5')
def preprocess_polozhenie_punkt_1_5(text: str) -> str:
    """Нормализация перечня документов для сравнения с шаблоном."""
    return preprocessors_iter8_normalize_normalize_text_for_comparison(text)

preprocessors_iter8_normalize_module = SimpleNamespace(normalize_text_for_comparison=preprocessors_iter8_normalize_normalize_text_for_comparison, preprocess_polozhenie_zagolovok=preprocess_polozhenie_zagolovok, preprocess_polozhenie_osnovnoy_tekst=preprocess_polozhenie_osnovnoy_tekst, preprocess_polozhenie_struktura=preprocess_polozhenie_struktura, preprocess_polozhenie_punkt_1_5=preprocess_polozhenie_punkt_1_5)

# END_SOURCE_PREPROCESSORS_ITER8_NORMALIZE

# START_SOURCE_PREPROCESSORS_POLOZHENIE_NORMALIZE
# PURPOSE: Inlined source from audit_engine/preprocessors/polozhenie_normalize.py.
@register_preprocessor('polozhenie_comp_ppu', 'шапка')
@register_preprocessor('polozhenie_ppu', 'шапка')
def preprocess_shapka(text: str) -> str:
    """
    Нормализует шапку:
    1) Удаляет дублирующийся хвост (Vision Parser иногда дублирует текст)
    2) Склеивает строку 'к приказу от <дата>' со строкой '№ <номер>'
       Vision иногда разбивает на 2 строки: 'от 13.11.2025г.
№ 10БП'
       Шаблон содержит их в одной строке: 'от <...> № <...>'
    """
    lines = text.split('\n')
    non_empty = [(i, l.strip()) for i, l in enumerate(lines) if l.strip()]
    if len(non_empty) >= 3:
        for tail_len in range(len(non_empty) // 2, 0, -1):
            tail_texts = [t for _, t in non_empty[-tail_len:]]
            for start in range(len(non_empty) - tail_len):
                match = all((non_empty[start + j][1] == tail_texts[j] for j in range(tail_len)))
                if match:
                    cut_line = non_empty[-tail_len][0]
                    while cut_line > 0 and (not lines[cut_line - 1].strip()):
                        cut_line -= 1
                    text = '\n'.join(lines[:cut_line])
                    break
            else:
                continue
            break
    text = re.sub('\\n{2,}', '\n', text)
    text = re.sub('(к приказу от[^\\n]*?)\\s*\\n\\s*(№)', '\\1 \\2', text)
    text = re.sub('к приказу от[^\\n]*\\n?', '', text)
    text = re.sub('^\\s*(?:Ne|No|№)\\s*["\\u201c\\u201e\\u00ab].*$', '', text, flags=re.MULTILINE)
    text = re.sub('(приложение)\\s*№\\s*\\S*', '\\1', text, flags=re.IGNORECASE)
    text = re.sub('\\n\\d+\\.\\s+.*', '', text, flags=re.DOTALL)
    text = text.lower()
    text = re.sub('\\s+', ' ', text).strip()
    return text

@register_preprocessor('polozhenie_comp_ppu', 'структура_разделов')
@register_preprocessor('polozhenie_ppu', 'структура_разделов')
@register_preprocessor('polozhenie_po', 'структура_разделов')
def normalize_structure(text: str) -> str:
    """
    Нормализует структуру разделов:
    1) Добавляет точку после голого номера раздела: '5 Порядок' → '5. Порядок'
       Vision иногда теряет точку при извлечении заголовков.
       НЕ затрагивает подразделы (2.2, 3.4.1) — у них уже есть точки.
    2) Убирает пустые строки
    """
    text = re.sub('^(\\d+)[ \\t]+', '\\1. ', text, flags=re.MULTILINE)
    known_sections = {'общие положения': '1', 'основные задачи': '2', 'организационная структура': '3', 'права': '4', 'ответственность': '5', 'порядок': '6'}
    restored_lines = []
    for line in text.split('\n'):
        stripped = line.strip()
        stripped_lower = stripped.lower()
        if stripped_lower in known_sections and (not re.match('^\\d', stripped)):
            num = known_sections[stripped_lower]
            restored_lines.append(f'{num}. {stripped}')
        else:
            restored_lines.append(line)
    text = '\n'.join(restored_lines)
    text = re.sub('\\n{2,}', '\n', text)
    lines = text.strip().split('\n')
    merged = []
    for i, line in enumerate(lines):
        stripped = line.strip()
        if not stripped:
            continue
        if re.match('^\\d', stripped):
            merged.append(stripped)
        elif merged:
            if 'не найдены' in stripped.lower():
                continue
            remaining = [l.strip() for l in lines[i + 1:] if l.strip()]
            has_next_section = any((re.match('^\\d', l) for l in remaining))
            if has_next_section:
                merged[-1] = merged[-1] + ' ' + stripped
    top_level = [l for l in merged if re.match('^\\d+\\.\\s', l)]
    if top_level:
        text = '\n'.join(top_level)
    return text

@register_preprocessor('polozhenie_comp_ppu', 'основной_текст')
@register_preprocessor('polozhenie_ppu', 'основной_текст')
@register_preprocessor('polozhenie_po', 'основной_текст')
def normalize_text(text: str) -> str:
    """
    Нормализует основной текст для сравнения с шаблоном.

    1) Заменяет блок «...на ФИО-Должность - председателя» на тег
       (и шаблон «на должность - председателя» тоже)
    2) Заменяет ФИО вида «Фамилия И.О.» на [ФИО]
    3) Нормализует множественные 
 → один 

    """
    text = re.sub('возлагается на .+?[ \\t]*-[ \\t]*председателя', 'возлагается на [ДОЛЖНОСТЬ_ФИО] - председателя', text)
    text = re.sub('[А-ЯЁ][а-яё]+\\s+[А-ЯЁ]\\.[А-ЯЁ]\\.', '[ФИО]', text)
    text = re.sub('[А-ЯЁ]\\.[А-ЯЁ]\\.\\s+[А-ЯЁ][а-яё]+', '[ФИО]', text)
    text = re.sub('\\n{2,}', '\n', text)
    text = re.sub('^(\\d+)[ \\t]+', '\\1. ', text, flags=re.MULTILINE)
    text = re.sub('^(\\d+\\.\\d+(?:\\.\\d+)?)\\.\\s', '\\1 ', text, flags=re.MULTILINE)
    lines_filtered = text.split('\n')
    text = '\n'.join((l for l in lines_filtered if not re.match('^\\d{1,2}$', l.strip())))
    text = '\n'.join((l for l in text.split('\n') if '\t' not in l))
    text = re.sub('\\nФорма №.*', '', text, flags=re.DOTALL)
    lines = text.split('\n')
    merged = []
    for line in lines:
        stripped = line.strip()
        if merged and stripped and (not re.match('^\\d', stripped)) and stripped[0].islower() and merged[-1].strip() and (merged[-1].strip()[-1] not in '.;:'):
            merged[-1] = merged[-1].rstrip() + ' ' + stripped
        else:
            merged.append(line)
    text = '\n'.join(merged)
    text = re.sub(';[ \\t]+([а-яё])', '.\\n\\1', text)
    text = re.sub('[;,](\\s*)$', '.\\1', text, flags=re.MULTILINE)
    text = re.sub('^[\\-–—]\\s*', '', text, flags=re.MULTILINE)
    text = re.sub('\\bне\\s+([а-яё])', 'не\\1', text)
    return text

preprocessors_polozhenie_normalize_module = SimpleNamespace(preprocess_shapka=preprocess_shapka, normalize_structure=normalize_structure, normalize_text=normalize_text)

# END_SOURCE_PREPROCESSORS_POLOZHENIE_NORMALIZE

# START_SOURCE_PREPROCESSORS_POLOZHENIE_PO
# PURPOSE: Inlined source from audit_engine/preprocessors/polozhenie_po.py.
@register_preprocessor('polozhenie_po', 'пункт_1_5')
def extract_longest_bullet_list(text: str) -> str:
    """
    Извлекает самый длинный маркированный список (•) из чанка.

    Алгоритм:
    1. Находит все строки с маркером •
    2. Группирует близкие буллеты (пустые строки между ними — норма для Paddle OCR)
    3. Выбирает группу с наибольшим количеством пунктов
    4. Возвращает заголовок секции + этот список

    Args:
        text: полный текст чанка пункт_1_5
    Returns:
        str: заголовок + самый длинный bullet-список
    """
    lines = text.split('\n')
    bullet_indices = [i for i, line in enumerate(lines) if line.strip().startswith('•')]
    if not bullet_indices:
        return text
    MAX_GAP = 3
    groups = []
    current_group = [bullet_indices[0]]
    for idx in bullet_indices[1:]:
        if idx - current_group[-1] <= MAX_GAP:
            current_group.append(idx)
        else:
            groups.append(current_group)
            current_group = [idx]
    groups.append(current_group)
    best_group = max(groups, key=len)
    best_start = best_group[0]
    best_end = best_group[-1]
    header_line = ''
    for i in range(best_start - 1, max(best_start - 5, -1), -1):
        if i < 0:
            break
        stripped = lines[i].strip()
        if stripped and (not stripped.startswith('•')) and (not stripped.startswith('[')):
            header_line = stripped
            break
    result_lines = []
    if header_line:
        result_lines.append(header_line)
    result_lines.extend(lines[best_start:best_end + 1])
    return '\n'.join(result_lines)

@register_preprocessor('polozhenie_po', 'лист_ознакомления')
def normalize_ознакомление(text: str) -> str:
    """
    Нормализует лист ознакомления для корректной проверки заполненности.

    Проблема: Paddle OCR не читает рукописный текст — ФИО и даты
    превращаются в мусор типа "('r>4' 1A,nQ(((%4". LLM видит мусор
    и считает строки пустыми.

    Решение: если в ячейке ФИО есть любые символы (длина > 2) — помечаем
    строку как "заполнено (рукописный текст, OCR не распознал)".

    Args:
        text: HTML-таблица листа ознакомления
    Returns:
        str: текст с пометками о заполненности
    """
    rows = re.findall('<tr>(.*?)</tr>', text, re.DOTALL)
    if len(rows) <= 1:
        return text
    filled_count = 0
    for row in rows[1:]:
        cells = re.findall('<td>(.*?)</td>', row, re.DOTALL)
        if len(cells) >= 2:
            fio_cell = cells[1].strip()
            cleaned = re.sub('[\\s\\-_|/\\\\.,;:!?\\\'"()\\[\\]{}]+', '', fio_cell)
            if len(cleaned) >= 1:
                filled_count += 1
    if filled_count > 0:
        summary = f'\n\n[СВОДКА: заполнено строк: {filled_count} (рукописный текст, OCR распознал частично)]'
        return text + summary
    return text

preprocessors_polozhenie_po_module = SimpleNamespace(extract_longest_bullet_list=extract_longest_bullet_list, normalize_ознакомление=normalize_ознакомление)

# END_SOURCE_PREPROCESSORS_POLOZHENIE_PO

# START_SOURCE_PREPROCESSORS_PRESENTATION_EU
# PURPOSE: Inlined source from audit_engine/preprocessors/presentation_eu.py.
def _compress_chunk(text: str, max_lines: int=3) -> str:
    """
    Сжимает текст чанка до первых значимых строк.

    Логика:
    1. Убираем пустые строки в начале
    2. Берём первые max_lines непустых строк (заголовки/начало контента)
    3. Добавляем сводку о размере оригинала
    """
    if not text or not text.strip():
        return '[Слайд отсутствует или пуст]'
    lines = text.strip().split('\n')
    meaningful = [l.strip() for l in lines if l.strip()]
    if not meaningful:
        return '[Слайд отсутствует или пуст]'
    header_lines = meaningful[:max_lines]
    total_lines = len(meaningful)
    result = '[Слайд присутствует]\n'
    result += '\n'.join(header_lines)
    if total_lines > max_lines:
        result += f'\n(... ещё {total_lines - max_lines} строк содержимого)'
    return result

@register_preprocessor('presentation_eu', 'титульный')
def compress_title(text: str) -> str:
    """Сжатие титульного слайда для структурной проверки."""
    return _compress_chunk(text)

@register_preprocessor('presentation_eu', 'план_мероприятий')
def compress_plan(text: str) -> str:
    """Сжатие слайда плана мероприятий."""
    return _compress_chunk(text)

@register_preprocessor('presentation_eu', 'инструменты_5с')
def compress_instruments(text: str) -> str:
    """Сжатие слайдов инструментов 5С."""
    return _compress_chunk(text, max_lines=5)

@register_preprocessor('presentation_eu', 'результаты_проблемы')
def compress_results(text: str) -> str:
    """Сжатие слайдов результатов."""
    return _compress_chunk(text, max_lines=4)

@register_preprocessor('presentation_eu', 'стандарты')
def compress_standards(text: str) -> str:
    """Сжатие слайдов стандартов."""
    return _compress_chunk(text)

@register_preprocessor('presentation_eu', 'последний')
def compress_last(text: str) -> str:
    """Сжатие последнего слайда."""
    return _compress_chunk(text)

preprocessors_presentation_eu_module = SimpleNamespace(_compress_chunk=_compress_chunk, compress_title=compress_title, compress_plan=compress_plan, compress_instruments=compress_instruments, compress_results=compress_results, compress_standards=compress_standards, compress_last=compress_last)

# END_SOURCE_PREPROCESSORS_PRESENTATION_EU

# START_SOURCE_PREPROCESSORS_PRIKAZ_COMP_PPU
# PURPOSE: Inlined source from audit_engine/preprocessors/prikaz_comp_ppu.py.
@register_preprocessor('prikaz_comp_ppu', 'текст_приказа')
def preprocessors_prikaz_comp_ppu_normalize_text_for_rule3(text: str) -> str:
    """
    Извлекает чистый текст приказа для сравнения с шаблоном (правило #3).

    Вход: полный текст страницы (шапка + заголовок + преамбула + пункты + подписант).
    Выход: от преамбулы «С целью...»/«В целях...» до последнего нумерованного пункта.

    Работает одинаково для таргета и шаблона — обе стороны получают
    идентичную структуру без шапки и подписанта.
    """
    lines = text.strip().split('\n')
    start_idx = None
    for i, line in enumerate(lines):
        s = line.strip()
        if s.startswith('С целью') or s.startswith('В целях'):
            start_idx = i
            break
        if s == 'ПРИКАЗЫВАЮ:':
            start_idx = i
            break
    if start_idx is None:
        start_idx = 0
    end_idx = len(lines) - 1
    for i in range(len(lines) - 1, -1, -1):
        s = lines[i].strip()
        if re.match('^\\d+\\.', s):
            end_idx = i
            break
    body = '\n'.join(lines[start_idx:end_idx + 1])
    body = re.sub('(возложить на\\s*).+\\.', '\\1[ДОЛЖНОСТЬ_ФИО].', body)
    body = re.sub('(возложить на\\s*)<[^>]+>\\s*и\\s*<[^>]+>', '\\1[ДОЛЖНОСТЬ_ФИО].', body)
    return body

preprocessors_prikaz_comp_ppu_module = SimpleNamespace(normalize_text_for_rule3=preprocessors_prikaz_comp_ppu_normalize_text_for_rule3)

# END_SOURCE_PREPROCESSORS_PRIKAZ_COMP_PPU

# START_SOURCE_PREPROCESSORS_PRIKAZ_IC
# PURPOSE: Inlined source from audit_engine/preprocessors/prikaz_ic.py.
def preprocessors_prikaz_ic_normalize_text_for_rule3(text: str) -> str:
    """
    Нормализует текст приказа для сравнения с шаблоном.

    Удаляет вариативные части (даты, ФИО, подписант),
    оставляя только текстовую структуру.
    """
    text = re.sub('(приступить к заполнению разделов по своим показателям)[ \\t]+с[ \\t]*[^\\n]*', '\\1', text)
    text = re.sub('(приступить к заполнению разделов по своим показателям)[ \\t]+с[ \\t]*\\[ДАТА\\]', '\\1', text)
    text = re.sub('4\\.\\s*.+?\\s+организовать', '4. [ДОЛЖНОСТЬ] организовать', text)
    text = re.sub('(в срок до)[ \\t]*[^\\n]+', '\\1', text)
    text = re.sub('(в срок до)[ \\t]*\\[ДАТА\\]', '\\1', text)
    lines = text.split('\n')
    new_lines = []
    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            new_lines.append(line)
            continue
        if line_stripped in ('[ПОДПИСАНТ]', 'Генеральный директор', 'И.О. Фамилия'):
            continue
        if re.search('\\s{5,}', line_stripped) and ('директор' in line_stripped.lower() or 'фамилия' in line_stripped.lower() or re.search('[А-Я]\\.[А-Я]\\.', line_stripped)):
            continue
        new_lines.append(line)
    text = '\n'.join(new_lines)
    lines = text.split('\n')
    text = '\n'.join((line.lstrip() for line in lines))
    return text

def preprocessors_prikaz_ic_normalize_header(text: str) -> str:
    """
    Нормализует шапку приказа — заменяет юридический адрес на краткий формат города.

    OCR иногда извлекает полный юрадрес: "109316, Москва г, Внутригородская..."
    LLM цепляется за формат адреса вместо проверки наличия города.
    Нормализуем: извлекаем город, заменяем всю строку адреса на "г. Город".
    """
    known_cities = ['Москва', 'Санкт-Петербург', 'Новосибирск', 'Екатеринбург', 'Казань', 'Нижний Новгород', 'Челябинск', 'Самара', 'Омск', 'Ростов-на-Дону', 'Уфа', 'Красноярск', 'Пермь', 'Воронеж', 'Волгоград', 'Краснодар', 'Тюмень', 'Тольятти', 'Барнаул']
    lines = text.split('\n')
    new_lines = []
    for line in lines:
        stripped = line.strip()
        if re.match('^\\d{5,6}\\s*,', stripped):
            for city in known_cities:
                if city.lower() in stripped.lower():
                    new_lines.append(f'г. {city}')
                    break
            else:
                new_lines.append(line)
        elif re.match('^(г\\.?\\s+)?(' + '|'.join(known_cities) + ')\\s+г\\.?$', stripped):
            city_match = re.search('(' + '|'.join(known_cities) + ')', stripped)
            if city_match:
                new_lines.append(f'г. {city_match.group(1)}')
            else:
                new_lines.append(line)
        else:
            new_lines.append(line)
    return '\n'.join(new_lines)

def trim_appendix2_to_relevant_sections(text: str) -> str:
    """
    Сокращает чанк приложение_2_к_приказу — убирает основное тело Регламента,
    оставляя заголовок (номер приказа, дата) и секцию «Приложение №1 к Регламенту».

    Без этого LLM путает нумерацию основного тела (1.1, 2.1, 4.1) с нумерацией
    Приложения к Регламенту (1.1.1, 2.1, 4.1), где находятся проверяемые поля.
    """
    lines = text.split('\n')
    header_lines = []
    appendix_lines = []
    in_appendix = False
    header_collected = False
    for line in lines:
        stripped = line.strip()
        if not header_collected:
            if re.match('^(РЕГЛАМЕНТ|1\\.\\s)', stripped):
                header_collected = True
            else:
                header_lines.append(line)
                continue
        if not in_appendix:
            if re.match('^Приложение\\s*№?\\s*1\\s*(к\\s+Регламенту|к\\s+регламенту)', stripped, re.IGNORECASE):
                in_appendix = True
                appendix_lines.append(line)
        else:
            appendix_lines.append(line)
    if appendix_lines:
        return '\n'.join(header_lines + [''] + appendix_lines)
    return text

preprocessors_prikaz_ic_module = SimpleNamespace(normalize_text_for_rule3=preprocessors_prikaz_ic_normalize_text_for_rule3, normalize_header=preprocessors_prikaz_ic_normalize_header, trim_appendix2_to_relevant_sections=trim_appendix2_to_relevant_sections)

# END_SOURCE_PREPROCESSORS_PRIKAZ_IC

# START_SOURCE_PREPROCESSORS_PRIKAZ_IC_POTOKA
# PURPOSE: Inlined source from audit_engine/preprocessors/prikaz_ic_potoka.py.
def preprocessors_prikaz_ic_potoka_normalize_text_for_rule3(text: str) -> str:
    """
    Нормализует текст приказа для сравнения с шаблоном.

    Заменяет переменные части на токены, оставляя структуру для сравнения.
    """
    text = re.sub('[ \\t]+', ' ', text)
    text = re.sub('\\n\\s*', '\n', text)
    text = re.sub('\\d{1,2}\\.\\d{1,2}\\.\\d{4}', '[ДАТА]', text)
    text = re.sub('_{2,}\\.\\s*_{2,}\\.\\s*202_?', '[ДАТА]', text)
    text = re.sub('\\(Указать наименование должности\\)', '[ДОЛЖНОСТЬ_ФИО]', text)
    text = re.sub('[А-ЯЁа-яё\\s]+[А-ЯЁ][а-яё]+\\s+[А-ЯЁ]\\.[А-ЯЁ]\\.\\s+организовать', '[ДОЛЖНОСТЬ_ФИО] организовать', text)
    text = re.sub('\\[ДАТА\\]\\s*\\[ДОЛЖНОСТЬ_ФИО\\]', '[ДАТА]\\n[ДОЛЖНОСТЬ_ФИО]', text)
    lines = text.split('\n')
    for i in range(len(lines) - 1, -1, -1):
        line = lines[i].strip()
        if line and ('директор' in line.lower() or 'должность' in line.lower() or 'фамилия' in line.lower() or ('фио' in line.lower()) or re.search('[А-ЯЁ]\\.[А-ЯЁ]\\.', line)):
            lines[i] = '[ПОДПИСАНТ]'
            break
    text = '\n'.join(lines)
    return text

@register_preprocessor('prikaz_ic_potoka', 'текст_приказа')
def preprocess_potoka_tekst(text: str) -> str:
    """Нормализация текста приказа ИЦ потока (бумажный)."""
    return preprocessors_prikaz_ic_potoka_normalize_text_for_rule3(text)

@register_preprocessor('prikaz_ic_potoka_el', 'текст_приказа')
def preprocess_potoka_el_tekst(text: str) -> str:
    """Нормализация текста приказа ИЦ потока (электронный)."""
    return preprocessors_prikaz_ic_potoka_normalize_text_for_rule3(text)

preprocessors_prikaz_ic_potoka_module = SimpleNamespace(normalize_text_for_rule3=preprocessors_prikaz_ic_potoka_normalize_text_for_rule3, preprocess_potoka_tekst=preprocess_potoka_tekst, preprocess_potoka_el_tekst=preprocess_potoka_el_tekst)

# END_SOURCE_PREPROCESSORS_PRIKAZ_IC_POTOKA

# START_SOURCE_PREPROCESSORS_PRIKAZ_PPU
# PURPOSE: Inlined source from audit_engine/preprocessors/prikaz_ppu.py.
@register_preprocessor('prikaz_ppu', 'заголовок_город')
def normalize_title(text: str) -> str:
    """
    Нормализует чанк заголовок_город для сравнения с шаблоном (правило #1).

    Vision извлекает элементы в разном порядке. Убираем вариативные части,
    оставляя ТОЛЬКО текст заголовка приказа.
    """
    lines = text.strip().split('\n')
    result = []
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if re.match('^П\\s*Р\\s*И\\s*К\\s*А\\s*З', stripped) or stripped.startswith('ПРИКАЗ'):
            continue
        if re.match('^г\\.\\s*', stripped):
            continue
        if re.match('^№\\s*', stripped):
            continue
        if re.match('^\\d{2}\\.\\d{2}\\.\\d{4}', stripped):
            continue
        if re.match('^[\\d"_<]', stripped) and (not re.match('^\\d+\\.', stripped)):
            continue
        if re.match('^(ООО|ОАО|ЗАО|ПАО|АО|ИП)\\s*[""«]', stripped):
            continue
        if stripped.startswith('<') and stripped.endswith('>'):
            continue
        if re.match('^[А-ЯЁ]\\.[А-ЯЁ]\\.\\s+[А-ЯЁ][а-яё]+$', stripped):
            continue
        result.append(stripped)
    return '\n'.join(result)

@register_preprocessor('prikaz_ppu', 'текст_приказа')
def preprocessors_prikaz_ppu_normalize_text_for_rule3(text: str) -> str:
    """
    Нормализует текст приказа о ППУ для сравнения с шаблоном (правило #3).

    Удаляет вариативные части (даты, ФИО, подписант),
    оставляя только текстовую структуру.
    """
    text = re.sub('(возложить на\\s*).+\\.', '\\1[ДОЛЖНОСТЬ_ФИО].', text)
    text = re.sub('(возложить на\\s*)<[^>]+>\\s*и\\s*<[^>]+>', '\\1[ДОЛЖНОСТЬ_ФИО].', text)
    text = re.sub('(в срок до)[ \\t]*[^\\n]+', '\\1', text)
    text = re.sub('\\d{2}\\.\\d{2}\\.\\d{4}(?:[ \\t]*г\\.?)?', '[ДАТА]', text)
    lines = text.split('\n')
    new_lines = []
    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            new_lines.append(line)
            continue
        if line_stripped in ('[ПОДПИСАНТ]', 'Генеральный директор', 'И.О. Фамилия'):
            continue
        if re.match('^[А-ЯЁ]\\.[А-ЯЁ]\\.\\s+[А-ЯЁ][а-яё]+$', line_stripped):
            continue
        if re.search('\\s{5,}', line_stripped) and ('директор' in line_stripped.lower() or 'фамилия' in line_stripped.lower() or re.search('[А-Я]\\.[А-Я]\\.', line_stripped)):
            continue
        if line_stripped.startswith('<') and line_stripped.endswith('>'):
            continue
        if re.match('^(ООО|ОАО|ЗАО|ПАО|АО|ИП)\\s*[""«]', line_stripped):
            continue
        new_lines.append(line)
    text = '\n'.join(new_lines)
    return text

preprocessors_prikaz_ppu_module = SimpleNamespace(normalize_title=normalize_title, normalize_text_for_rule3=preprocessors_prikaz_ppu_normalize_text_for_rule3)

# END_SOURCE_PREPROCESSORS_PRIKAZ_PPU

# START_SOURCE_PREPROCESSORS_PRIKAZ_VYHOD
# PURPOSE: Inlined source from audit_engine/preprocessors/prikaz_vyhod.py.
@register_preprocessor('prikaz_vyhod', 'текст_приказа')
def preprocessors_prikaz_vyhod_normalize_text_for_rule3(text: str) -> str:
    """
    Нормализует текст приказа для сравнения с шаблоном (правило #3).

    Заменяет:
    - п.1: "Назначить организатором...площадке <должность ФИО>." → [ДОЛЖНОСТЬ_ФИО]
    - п.2: "Назначить секретарем...площадке <должность ФИО>." → [ДОЛЖНОСТЬ_ФИО]

    ВАЖНО: используем DOTALL + non-greedy (.+?) с остановкой на границе абзаца,
    чтобы не съесть следующие пункты приказа (баг со старым жадным .+\\. + DOTALL).
    """
    text = re.sub('(1\\.\\s*Назначить организатором проведения обхода на\\s+(?:производственной\\s+)?площадке\\s+)(.+?)(?=\\s*\\n\\s*\\n|\\s*\\n\\s*\\d+\\.|\\s*$)', '\\1[ДОЛЖНОСТЬ_ФИО].', text, flags=re.DOTALL)
    text = re.sub('(2\\.\\s*Назначить секретарем проведения обхода на\\s+(?:производственной\\s+)?площадке\\s+)(.+?)(?=\\s*\\n\\s*\\n|\\s*\\n\\s*\\d+\\.|\\s*$)', '\\1[ДОЛЖНОСТЬ_ФИО].', text, flags=re.DOTALL)
    return text

@register_preprocessor('prikaz_vyhod', 'заголовок_город')
def preprocessors_prikaz_vyhod_normalize_header(text: str) -> str:
    """
    Нормализует заголовок для сравнения с шаблоном (правило #1).

    «П Р И К А З» (разреженное написание) → «Приказ» (как в шаблоне).
    gpt-4.1-mini игнорирует инструкцию в промпте, поэтому нормализуем до LLM.
    """
    text = re.sub('П\\s+Р\\s+И\\s+К\\s+А\\s+З', 'Приказ', text)
    return text

preprocessors_prikaz_vyhod_module = SimpleNamespace(normalize_text_for_rule3=preprocessors_prikaz_vyhod_normalize_text_for_rule3, normalize_header=preprocessors_prikaz_vyhod_normalize_header)

# END_SOURCE_PREPROCESSORS_PRIKAZ_VYHOD
# END_PREPROCESSORS

# START_PARSERS
# PURPOSE: Extract structured text from docx, pptx, pdf, xlsx, OCR, and special-engine sources.
# INPUTS: document file paths, chunk configs, model endpoints.
# OUTPUTS: Parsed document dictionaries and parser artifacts.
# KEYWORDS: parsers, docx, pptx, paddle, ocr, vision, xlsx.
# LINKS: audit_engine/*parser*.py, audit_engine/*/parser*.py.
# RATIONALE: Parsing is the widest dependency fan-out, so it sits before validation and runtime orchestration.

# START_SOURCE_VISION_PARSER_CONVERTER
# PURPOSE: Inlined source from audit_engine/vision_parser/converter.py.
class ConversionError(Exception):
    """Ошибка конвертации документа."""
    pass

class DocumentConverter:
    """
    Конвертер документов в PDF с использованием LibreOffice.

    Поддерживаемые форматы:
    - .docx, .doc, .odt → конвертация через LibreOffice
    - .pdf → возврат исходного пути без конвертации

    Атрибуты:
        temp_dir: Директория для временных файлов
        libreoffice_path: Путь к исполняемому файлу LibreOffice
    """
    CONVERTIBLE_EXTENSIONS = {'.docx', '.doc', '.odt', '.rtf', '.pptx', '.ppt', '.odp'}

    def __init__(self, temp_dir: Optional[str]=None):
        """
        Инициализация конвертера.

        Args:
            temp_dir: Директория для временных файлов.
                     Если не указана, используется системная temp директория.
        """
        self.temp_dir = Path(temp_dir) if temp_dir else Path(tempfile.gettempdir())
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self._created_files: list = []
        self.libreoffice_path = self._find_libreoffice()

    def _find_libreoffice(self) -> str:
        """
        Находит путь к LibreOffice.

        Returns:
            Путь к исполняемому файлу LibreOffice

        Raises:
            ConversionError: Если LibreOffice не найден
        """
        possible_paths = ['/Applications/LibreOffice.app/Contents/MacOS/soffice', '/usr/bin/soffice', '/usr/bin/libreoffice', '/usr/local/bin/soffice', 'soffice']
        for cmd in ['soffice', 'libreoffice']:
            result = shutil.which(cmd)
            if result:
                return result
        for path in possible_paths:
            if os.path.exists(path):
                return path
        raise ConversionError('LibreOffice не найден. Установите LibreOffice:\n  macOS: brew install --cask libreoffice\n  Linux: apt install libreoffice\n  Windows: https://www.libreoffice.org/download/')

    def convert(self, input_path: str) -> str:
        """
        Конвертирует документ в PDF.

        Args:
            input_path: Путь к исходному документу

        Returns:
            Путь к PDF файлу:
            - Исходный путь, если файл уже PDF
            - Путь к временному PDF, если была конвертация

        Raises:
            FileNotFoundError: Если исходный файл не найден
            ConversionError: Если конвертация не удалась
        """
        input_file = Path(input_path)
        if not input_file.exists():
            raise FileNotFoundError(f'Файл не найден: {input_path}')
        ext = input_file.suffix.lower()
        if ext == '.pdf':
            return str(input_file.absolute())
        if ext not in self.CONVERTIBLE_EXTENSIONS:
            raise ConversionError(f'Неподдерживаемый формат: {ext}. Поддерживаются: {', '.join(self.CONVERTIBLE_EXTENSIONS)}, .pdf')
        return self._convert_with_libreoffice(input_file)

    def _convert_with_libreoffice(self, input_file: Path) -> str:
        """
        Конвертирует документ в PDF через LibreOffice headless.

        Args:
            input_file: Путь к исходному файлу

        Returns:
            Путь к созданному PDF

        Raises:
            ConversionError: Если конвертация не удалась
        """
        output_dir = self.temp_dir / 'vision_converter'
        old_pdf = output_dir / f'{input_file.stem}.pdf'
        if old_pdf.exists():
            old_pdf.unlink()
        output_dir.mkdir(exist_ok=True)
        cmd = [self.libreoffice_path, '--headless', '--convert-to', 'pdf', '--outdir', str(output_dir), str(input_file.absolute())]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            if result.returncode != 0:
                raise ConversionError(f'LibreOffice вернул ошибку: {result.stderr}')
        except subprocess.TimeoutExpired:
            raise ConversionError('Таймаут конвертации (>120 секунд). Возможно, документ слишком большой.')
        except FileNotFoundError:
            raise ConversionError(f'Не удалось запустить LibreOffice: {self.libreoffice_path}')
        output_pdf = output_dir / f'{input_file.stem}.pdf'
        if not output_pdf.exists():
            raise ConversionError(f'PDF не создан. Вывод LibreOffice:\nstdout: {result.stdout}\nstderr: {result.stderr}')
        self._created_files.append(str(output_pdf))
        return str(output_pdf)

    def cleanup(self, pdf_path: str=None) -> None:
        """
        Удаляет временные файлы.

        Args:
            pdf_path: Конкретный PDF для удаления.
                     Если не указан, удаляет все созданные файлы.
        """
        if pdf_path:
            path = Path(pdf_path)
            if path.exists() and str(path) in self._created_files:
                path.unlink()
                self._created_files.remove(str(path))
        else:
            for file_path in self._created_files:
                path = Path(file_path)
                if path.exists():
                    path.unlink()
            self._created_files.clear()

    def __enter__(self):
        """Поддержка context manager."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Автоматическая очистка при выходе из context manager."""
        self.cleanup()
        return False

def convert_to_pdf(input_path: str, output_path: str=None) -> str:
    """
    Утилитарная функция для быстрой конвертации.

    Args:
        input_path: Путь к исходному документу
        output_path: Путь для результата (опционально)

    Returns:
        Путь к PDF файлу
    """
    converter = DocumentConverter()
    pdf_path = converter.convert(input_path)
    if output_path:
        shutil.copy(pdf_path, output_path)
        converter.cleanup(pdf_path)
        return output_path
    return pdf_path

vision_parser_converter_module = SimpleNamespace(ConversionError=ConversionError, DocumentConverter=DocumentConverter, convert_to_pdf=convert_to_pdf)

# END_SOURCE_VISION_PARSER_CONVERTER

# START_SOURCE_VISION_PARSER_AGGREGATOR
# PURPOSE: Inlined source from audit_engine/vision_parser/aggregator.py.
class ChunkAggregator:
    """
    Агрегатор для объединения и нормализации чанков.

    Преобразует сырые результаты Vision API в формат,
    совместимый с legacy-парсером для обратной совместимости.
    """

    def aggregate(self, raw_chunks: Dict[str, str], file_path: str) -> Dict[str, Any]:
        """
        Агрегирует и нормализует извлечённые чанки.

        Args:
            raw_chunks: Словарь {chunk_name: raw_text} от экстрактора-источника
            file_path: Путь к исходному файлу

        Returns:
            Словарь в формате legacy-парсера:
            {
                "filename": str,
                "path": str,
                "шапка": str,
                "преамбула": str,
                ...
            }
        """
        path = Path(file_path)
        result = {'filename': path.name, 'path': str(path.absolute())}
        for chunk_name, raw_text in raw_chunks.items():
            normalized = self._normalize(raw_text)
            result[chunk_name] = normalized
        return result

    def _normalize(self, text: str) -> str:
        """
        Нормализует текст чанка.

        Операции:
        1. Убирает markdown-разметку (если LLM её добавил)
        2. Нормализует пробелы и переносы строк
        3. Убирает артефакты OCR

        Args:
            text: Сырой текст от Vision LLM

        Returns:
            Нормализованный текст
        """
        if not text:
            return ''
        text = re.sub('```[\\w]*\\n?', '', text)
        text = re.sub('```', '', text)
        text = re.sub('^#+\\s*', '', text, flags=re.MULTILINE)
        text = re.sub('\\*\\*([^*]+)\\*\\*', '\\1', text)
        text = re.sub('\\*([^*]+)\\*', '\\1', text)
        text = re.sub('__([^_]+)__', '\\1', text)
        text = re.sub('_([^_]+)_', '\\1', text)
        text = re.sub('([^\\n]) {2,}', '\\1 ', text)
        text = re.sub('\\n{3,}', '\n\n', text)
        text = re.sub(' +\\n', '\n', text)
        text = re.sub('[\\x00-\\x08\\x0b\\x0c\\x0e-\\x1f]', '', text)
        text = text.replace('–', '-')
        text = text.replace('—', '-')
        text = text.replace('«', '"')
        text = text.replace('»', '"')
        text = text.replace('„', '"')
        text = text.replace('"', '"')
        text = text.replace('"', '"')
        text = text.strip()
        text = self._dedup_loop_lines(text)
        return text

    def _dedup_loop_lines(self, text: str, max_repeats: int=2) -> str:
        """
        Удаляет loop-артефакты Vision LLM — строки, повторяющиеся 3+ раз подряд.

        Vision API (gpt-4.1-mini) может зациклиться при обработке таблиц/списков,
        генерируя сотни одинаковых строк. Оставляем max_repeats экземпляров.

        Args:
            text: Текст после нормализации
            max_repeats: Максимум повторов одной строки подряд (по умолчанию 2)

        Returns:
            Текст без loop-артефактов
        """
        lines = text.split('\n')
        result = []
        prev_line = None
        repeat_count = 0
        for line in lines:
            stripped = line.strip()
            if stripped == prev_line and stripped:
                repeat_count += 1
                if repeat_count <= max_repeats:
                    result.append(line)
            else:
                prev_line = stripped
                repeat_count = 1
                result.append(line)
        return '\n'.join(result)

    def merge_chunks(self, chunks: Dict[str, str], order: list=None) -> str:
        """
        Объединяет несколько чанков в один текст.

        Args:
            chunks: Словарь чанков
            order: Порядок объединения (если не указан, используется порядок словаря)

        Returns:
            Объединённый текст
        """
        if order is None:
            order = list(chunks.keys())
        parts = []
        for name in order:
            if name in chunks and chunks[name]:
                parts.append(chunks[name])
        return '\n\n'.join(parts)

def aggregate_results(raw_chunks: Dict[str, str], file_path: str) -> Dict[str, Any]:
    """
    Утилитарная функция для быстрой агрегации.

    Args:
        raw_chunks: Сырые чанки от экстрактора-источника
        file_path: Путь к исходному файлу

    Returns:
        Агрегированный результат
    """
    aggregator = ChunkAggregator()
    return aggregator.aggregate(raw_chunks, file_path)

vision_parser_aggregator_module = SimpleNamespace(ChunkAggregator=ChunkAggregator, aggregate_results=aggregate_results)

# END_SOURCE_VISION_PARSER_AGGREGATOR

# START_SOURCE_OCR_PARSER_NORMALIZER
# PURPOSE: Inlined source from audit_engine/ocr_parser/normalizer.py.
def html_table_to_markdown(html: str) -> str:
    """
    Конвертация HTML-таблицы в Markdown-формат.

    Args:
        html: строка с <table>...</table>

    Returns:
        str: таблица в формате | col1 | col2 |
    """
    rows = re.findall('<tr>(.*?)</tr>', html, re.DOTALL)
    md_rows = []
    for row in rows:
        cells = re.findall('<(?:td|th)[^>]*>(.*?)</(?:td|th)>', row, re.DOTALL)
        clean = [re.sub('<[^>]+>', ' ', c).strip() for c in cells]
        if clean:
            md_rows.append('| ' + ' | '.join(clean) + ' |')
    if not md_rows:
        return html
    result = [md_rows[0]]
    ncols = md_rows[0].count('|') - 1
    result.append('|' + '|'.join(['---'] * ncols) + '|')
    result.extend(md_rows[1:])
    return '\n'.join(result)

def strip_arrow_annotations(text: str) -> str:
    """
    Удаляет аннотации Word-форм вида 'описание поля -> значение'.

    Такие аннотации появляются в документах с формами, где поля имеют
    подписи типа 'Форма и наименование предприятия -> ООО "Ромашка"'.
    HunyuanOCR захватывает их как текст, что вызывает ложные срабатывания
    при сверке с шаблоном.

    Строка с ' -> ' заменяется на значение после стрелки.
    Строки без ' -> ' не затрагиваются.

    Args:
        text: OCR-текст страницы

    Returns:
        str: текст без аннотаций (значения сохраняются)
    """
    if not text:
        return ''
    lines = text.split('\n')
    result = []
    for line in lines:
        if ' -> ' in line:
            value = line.split(' -> ', 1)[1].strip()
            if value:
                result.append(value)
        else:
            result.append(line)
    return '\n'.join(result)

def unwrap_latex_artifacts(text: str) -> str:
    """
    Убирает LaTeX-обёртки из OCR-текста HunyuanOCR.

    HunyuanOCR оборачивает подчёркнутый/выделенный текст в LaTeX-формат:
    $ \\underline{\\mathrm{TEXT}} $ → TEXT

    Особые случаи:
    - \\mathrm{No} → № (знак номера)

    Безопасно для реальных документов: LaTeX-паттерны не встречаются
    в обычном тексте, только в артефактах OCR.

    Args:
        text: OCR-текст с возможными LaTeX-артефактами

    Returns:
        str: текст с развёрнутыми LaTeX-обёртками
    """
    if not text or '$' not in text:
        return text

    def replace_latex(match):
        inner = match.group(1)
        if inner == 'No':
            return '№'
        return inner
    text = re.sub('\\$\\s*\\\\underline\\{\\\\mathrm\\{([^}]*)\\}\\}\\s*\\$', replace_latex, text)
    return text

def normalize_ocr_text(text: str) -> str:
    """
    Нормализация OCR-текста: LaTeX-артефакты, HTML-таблицы -> Markdown, cleanup.

    Операции:
    1. Разворачивает LaTeX-обёртки HunyuanOCR (underline/mathrm)
    2. Заменяет <table>...</table> блоки на Markdown-таблицы
    3. Убирает лишние пустые строки
    4. Убирает пробелы в конце строк

    Args:
        text: OCR-текст (может содержать LaTeX и <table> от HunyuanOCR)

    Returns:
        str: нормализованный текст
    """
    if not text:
        return ''
    text = unwrap_latex_artifacts(text)

    def replace_table(match):
        return html_table_to_markdown(match.group(0))
    text = re.sub('<table>.*?</table>', replace_table, text, flags=re.DOTALL)
    text = re.sub('\\n{3,}', '\n\n', text)
    text = re.sub(' +\\n', '\n', text)
    return text.strip()

ocr_parser_normalizer_module = SimpleNamespace(html_table_to_markdown=html_table_to_markdown, strip_arrow_annotations=strip_arrow_annotations, unwrap_latex_artifacts=unwrap_latex_artifacts, normalize_ocr_text=normalize_ocr_text)

# END_SOURCE_OCR_PARSER_NORMALIZER

# START_SOURCE_OCR_PARSER_ASSEMBLER
# PURPOSE: Inlined source from audit_engine/ocr_parser/assembler.py.
ocr_parser_assembler_log = logging.getLogger(__name__)

class ChunkAssembler:
    """
    Собирает чанки документа из постраничного OCR-текста.

    Для каждого чанка: берёт pages из конфига -> собирает текст из page_texts ->
    нормализует (HTML-таблицы -> Markdown) -> возвращает {chunk_name: text}.
    """

    def assemble(self, page_texts: Dict[int, str], chunks: List[ChunkConfig], total_pages: int, strip_annotations_chunks: Optional[List[str]]=None, normalizer_fn=None) -> Dict[str, str]:
        """
        Собирает чанки из постраничных OCR-текстов.

        Args:
            page_texts: {1-based page_num: OCR-текст} от экстрактора-источника
            chunks: конфигурации чанков из chunks_vision.json
            total_pages: общее количество страниц в PDF
            strip_annotations_chunks: имена чанков, где убирать
                аннотации Word-форм (паттерн 'текст -> значение')
            normalizer_fn: функция нормализации текста (по умолчанию normalize_ocr_text).
                Для Paddle pipeline передаётся normalize_paddle_text,
                которая сохраняет HTML-таблицы с colspan/rowspan.

        Returns:
            Dict[str, str]: {chunk_name: собранный_и_нормализованный_текст}
        """
        _normalize = normalizer_fn or normalize_ocr_text
        _strip_chunks = set(strip_annotations_chunks or [])
        result = {}
        for chunk in chunks:
            page_indices = chunk.get_page_indices(total_pages)
            page_nums = [idx + 1 for idx in page_indices]
            need_strip = chunk.name in _strip_chunks
            parts = []
            for pn in page_nums:
                if pn in page_texts:
                    text = page_texts[pn]
                    if need_strip:
                        text = strip_arrow_annotations(text)
                    if text and text.strip():
                        if len(page_nums) > 1:
                            parts.append(f'[СТРАНИЦА {pn}]\n{text.strip()}')
                        else:
                            parts.append(text.strip())
                else:
                    ocr_parser_assembler_log.warning(f"Чанк '{chunk.name}': страница {pn} отсутствует в OCR")
            if not parts:
                result[chunk.name] = ''
                ocr_parser_assembler_log.warning(f"Чанк '{chunk.name}': пустой результат (страницы {page_nums})")
                continue
            raw_text = '\n\n'.join(parts)
            normalized = _normalize(raw_text)
            result[chunk.name] = normalized
            ocr_parser_assembler_log.info(f"Чанк '{chunk.name}': страницы {page_nums}, {len(normalized)} символов")
        return result

def collect_unique_pages(chunks: List[ChunkConfig], total_pages: int) -> List[int]:
    """
    Собирает уникальные 0-based индексы страниц из ВСЕХ чанков.

    Используется для дедупликации: каждую страницу OCR'им один раз.

    Args:
        chunks: список конфигураций чанков
        total_pages: общее количество страниц

    Returns:
        List[int]: отсортированные уникальные 0-based индексы
    """
    all_indices = set()
    for chunk in chunks:
        indices = chunk.get_page_indices(total_pages)
        all_indices.update(indices)
    return sorted(all_indices)

ocr_parser_assembler_module = SimpleNamespace(log=ocr_parser_assembler_log, ChunkAssembler=ChunkAssembler, collect_unique_pages=collect_unique_pages)

# END_SOURCE_OCR_PARSER_ASSEMBLER

# START_SOURCE_PADDLE_PARSER_LAYOUT_CLIENT
# PURPOSE: Inlined source from audit_engine/paddle_parser/layout_client.py.
paddle_parser_layout_client_logger = logging.getLogger(__name__)

class RemoteLayoutDetector:
    """
    Удалённый layout detector через HTTP API.

    Совместим по интерфейсу с LayoutDetector — метод detect() принимает
    PIL Image и возвращает список регионов.

    Args:
        base_url: URL layout-сервиса (например http://172.16.10.35:8012)
        timeout: таймаут запроса в секундах
    """

    def __init__(self, base_url: str, timeout: int=60):
        self.base_url = base_url.rstrip('/')
        self.timeout = timeout
        paddle_parser_layout_client_logger.info(f'RemoteLayoutDetector: {self.base_url}')

    def detect(self, image: Union[str, Path, Image.Image]) -> List[Dict[str, Any]]:
        """
        Отправляет изображение на удалённый layout-сервис.

        Args:
            image: путь к файлу или PIL Image

        Returns:
            [{"class_name": str, "bbox": [x1,y1,x2,y2], "score": float}, ...]
        """
        if isinstance(image, (str, Path)):
            image = Image.open(image).convert('RGB')
        elif image.mode != 'RGB':
            image = image.convert('RGB')
        buf = io.BytesIO()
        image.save(buf, format='PNG')
        buf.seek(0)
        try:
            resp = requests.post(f'{self.base_url}/detect', files={'file': ('page.png', buf, 'image/png')}, timeout=self.timeout)
            resp.raise_for_status()
            data = resp.json()
            regions = data.get('regions', [])
            time_ms = data.get('time_ms', 0)
            paddle_parser_layout_client_logger.info(f'Remote layout: {len(regions)} регионов за {time_ms}ms')
            return regions
        except requests.exceptions.RequestException as e:
            paddle_parser_layout_client_logger.error(f'Layout API ошибка: {e}')
            raise ConnectionError(f'Layout Detection API недоступен: {self.base_url} — {e}') from e

    def visualize(self, image, regions, output_path):
        """Визуализация — делегируем в PIL (без torch)."""
        from PIL import ImageDraw
        if isinstance(image, (str, Path)):
            img = Image.open(image).convert('RGB')
        else:
            img = image.copy().convert('RGB')
        draw = ImageDraw.Draw(img)
        colors = {'Text': 'blue', 'Title': 'red', 'Section-header': 'darkred', 'Table': 'green', 'Picture': 'purple', 'Caption': 'pink'}
        for i, region in enumerate(regions):
            bbox = region['bbox']
            cls = region['class_name']
            color = colors.get(cls, 'yellow')
            draw.rectangle(bbox, outline=color, width=3)
            draw.text((bbox[0] + 2, bbox[1] + 2), f'[{i}] {cls} {region['score']:.2f}', fill=color)
        img.save(output_path)

paddle_parser_layout_client_module = SimpleNamespace(logger=paddle_parser_layout_client_logger, RemoteLayoutDetector=RemoteLayoutDetector)

# END_SOURCE_PADDLE_PARSER_LAYOUT_CLIENT

# START_SOURCE_PADDLE_PARSER_LAYOUT_DETECTOR
# PURPOSE: Inlined source from audit_engine/paddle_parser/layout_detector.py.
try:
    import torch
except ImportError:
    torch = None

paddle_parser_layout_detector_logger = logging.getLogger(__name__)

class LayoutDetector:
    """
    Детектор layout документов на базе Docling Heron-101 (RT-DETRv2).

    Args:
        model_name: HF repo_id модели. По умолчанию — из config.
        device: устройство для инференса ("cuda:0" / "cpu").
    """

    def __init__(self, model_name: Optional[str]=None, device: str='cuda:0'):
        if torch is None:
            raise ImportError('torch не установлен — используйте RemoteLayoutDetector через layout_base_url')
        from transformers import RTDetrV2ForObjectDetection, RTDetrImageProcessor
        repo = model_name or paddle_parser_config_module.LAYOUT_MODEL_REPO
        paddle_parser_layout_detector_logger.info(f'Загрузка Docling Heron-101: {repo} (device={device})')
        self.processor = RTDetrImageProcessor.from_pretrained(repo)
        torch.cuda.empty_cache()
        self.device = device
        self.model = RTDetrV2ForObjectDetection.from_pretrained(repo, torch_dtype=torch.float16).to(device)
        self.model.eval()
        self.id2label = self.model.config.id2label
        paddle_parser_layout_detector_logger.info(f'Docling Heron-101 загружена, классов: {len(self.id2label)}')
        paddle_parser_layout_detector_logger.info(f'Классы: {self.id2label}')

    def detect(self, image: Union[str, Path, Image.Image]) -> List[Dict[str, Any]]:
        """
        Детектирует layout-элементы на изображении.

        Args:
            image: путь к файлу или PIL Image.

        Returns:
            Список словарей:
            [{"class_name": str, "bbox": [x1,y1,x2,y2], "score": float}, ...]
            Отсортировано по score desc.
        """
        if isinstance(image, (str, Path)):
            image = Image.open(image).convert('RGB')
        elif image.mode != 'RGB':
            image = image.convert('RGB')
        inputs = self.processor(images=[image], return_tensors='pt')
        model_dtype = next(self.model.parameters()).dtype
        inputs = {k: v.to(device=self.device, dtype=model_dtype if v.is_floating_point() else None) for k, v in inputs.items()}
        with torch.no_grad():
            outputs = self.model(**inputs)
        target_sizes = torch.tensor([image.size[::-1]], device=self.device)
        results = self.processor.post_process_object_detection(outputs, target_sizes=target_sizes, threshold=paddle_parser_config_module.LAYOUT_CONF_THRESHOLD)[0]
        regions = []
        for score, label_id, box in zip(results['scores'], results['labels'], results['boxes']):
            class_name = self.id2label[label_id.item()]
            regions.append({'class_name': class_name, 'bbox': box.tolist(), 'score': score.item()})
        paddle_parser_layout_detector_logger.info(f'Детекция: {len(regions)} регионов (до фильтрации)')
        regions = self._post_filter(regions)
        paddle_parser_layout_detector_logger.info(f'После фильтрации: {len(regions)} регионов')
        return regions

    def _post_filter(self, regions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Пост-фильтрация: убрать слабые, дедупликация по IoU и containment NMS.

        Containment NMS удаляет вложенные регионы: если маленький регион на >70%
        внутри большого (или наоборот), убираем тот, у которого ниже score.
        Это решает проблему, когда модель детектирует и крупный Text-блок,
        и отдельные List-item/Section-header внутри него.

        Args:
            regions: сырые детекции.

        Returns:
            Отфильтрованный список.
        """
        regions = [r for r in regions if r['score'] >= paddle_parser_config_module.LAYOUT_POST_CONF_MIN]
        regions.sort(key=lambda r: r['score'], reverse=True)
        kept = []
        for region in regions:
            is_dup = False
            for existing in kept:
                if self._iou(region['bbox'], existing['bbox']) > paddle_parser_config_module.LAYOUT_DEDUP_IOU:
                    is_dup = True
                    break
            if not is_dup:
                kept.append(region)
        kept.sort(key=lambda r: r['score'], reverse=True)
        final = []
        for region in kept:
            is_contained = False
            for existing in final:
                if self._containment(region['bbox'], existing['bbox']) > paddle_parser_config_module.LAYOUT_CONTAINMENT_THR:
                    is_contained = True
                    break
            if not is_contained:
                final.append(region)
        return final

    @staticmethod
    def _iou(box1: list, box2: list) -> float:
        """
        Вычисляет IoU двух bbox [x1,y1,x2,y2].

        Args:
            box1: первый bbox.
            box2: второй bbox.

        Returns:
            IoU значение (0.0 - 1.0).
        """
        x1 = max(box1[0], box2[0])
        y1 = max(box1[1], box2[1])
        x2 = min(box1[2], box2[2])
        y2 = min(box1[3], box2[3])
        inter = max(0, x2 - x1) * max(0, y2 - y1)
        area1 = (box1[2] - box1[0]) * (box1[3] - box1[1])
        area2 = (box2[2] - box2[0]) * (box2[3] - box2[1])
        union = area1 + area2 - inter
        return inter / union if union > 0 else 0.0

    @staticmethod
    def _containment(box1: list, box2: list) -> float:
        """
        Вычисляет степень вложенности: intersection / min(area1, area2).

        Высокое значение означает, что меньший bbox почти целиком внутри большего.
        Используется для удаления дублирующих вложенных регионов.

        Args:
            box1: первый bbox [x1,y1,x2,y2].
            box2: второй bbox [x1,y1,x2,y2].

        Returns:
            Containment ratio (0.0 - 1.0).
        """
        x1 = max(box1[0], box2[0])
        y1 = max(box1[1], box2[1])
        x2 = min(box1[2], box2[2])
        y2 = min(box1[3], box2[3])
        inter = max(0, x2 - x1) * max(0, y2 - y1)
        area1 = (box1[2] - box1[0]) * (box1[3] - box1[1])
        area2 = (box2[2] - box2[0]) * (box2[3] - box2[1])
        min_area = min(area1, area2)
        return inter / min_area if min_area > 0 else 0.0

    def visualize(self, image: Union[str, Path, Image.Image], regions: List[Dict[str, Any]], output_path: str) -> None:
        """
        Рисует bbox на изображении и сохраняет результат.

        Args:
            image: исходное изображение.
            regions: список регионов с bbox.
            output_path: путь для сохранения аннотированного изображения.
        """
        from PIL import ImageDraw
        if isinstance(image, (str, Path)):
            img = Image.open(image).convert('RGB')
        else:
            img = image.copy().convert('RGB')
        draw = ImageDraw.Draw(img)
        colors = {'Text': 'blue', 'Title': 'red', 'Section-header': 'darkred', 'Table': 'green', 'Picture': 'purple', 'Formula': 'orange', 'Caption': 'pink', 'Footnote': 'brown', 'Page-header': 'gray', 'Page-footer': 'gray', 'List-item': 'cyan', 'Document Index': 'teal', 'Code': 'darkblue', 'Checkbox-Selected': 'lime', 'Checkbox-Unselected': 'olive', 'Form': 'magenta', 'Key-Value Region': 'gold'}
        for i, region in enumerate(regions):
            bbox = region['bbox']
            cls = region['class_name']
            score = region['score']
            color = colors.get(cls, 'yellow')
            draw.rectangle(bbox, outline=color, width=3)
            label = f'[{i}] {cls} {score:.2f}'
            draw.text((bbox[0] + 2, bbox[1] + 2), label, fill=color)
        img.save(output_path)
        paddle_parser_layout_detector_logger.info(f'Визуализация сохранена: {output_path}')

paddle_parser_layout_detector_module = SimpleNamespace(logger=paddle_parser_layout_detector_logger, LayoutDetector=LayoutDetector)

# END_SOURCE_PADDLE_PARSER_LAYOUT_DETECTOR

# START_SOURCE_PADDLE_PARSER_READING_ORDER
# PURPOSE: Inlined source from audit_engine/paddle_parser/reading_order.py.
paddle_parser_reading_order_logger = logging.getLogger(__name__)

def sort_by_reading_order(regions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Сортирует регионы по reading order (XY-cut).

    Args:
        regions: список [{"class_name": str, "bbox": [x1,y1,x2,y2], ...}, ...]

    Returns:
        Тот же список, отсортированный в порядке чтения.
    """
    if len(regions) <= 1:
        return regions
    indices = list(range(len(regions)))
    ordered_indices = _xycut(indices, regions)
    result = [regions[i] for i in ordered_indices]
    paddle_parser_reading_order_logger.info(f'Reading order: {len(result)} элементов отсортировано')
    return result

def _xycut(indices: List[int], regions: List[Dict[str, Any]]) -> List[int]:
    """
    Рекурсивный XY-cut.

    1. Проецирует bbox на Y-ось → ищет горизонтальные разрывы.
    2. Разбивает на горизонтальные полосы.
    3. Внутри каждой полосы проецирует на X-ось → ищет колонки.
    4. Рекурсивно повторяет.

    Args:
        indices: индексы регионов для сортировки.
        regions: все регионы (для доступа к bbox).

    Returns:
        Отсортированный список индексов.
    """
    if len(indices) <= 1:
        return indices
    boxes = [(i, regions[i]['bbox']) for i in indices]
    y_groups = _split_axis(boxes, axis='y')
    if len(y_groups) > 1:
        result = []
        for group in y_groups:
            group_indices = [idx for idx, _ in group]
            result.extend(_xycut(group_indices, regions))
        return result
    x_groups = _split_axis(boxes, axis='x')
    if len(x_groups) > 1:
        result = []
        for group in x_groups:
            group_indices = [idx for idx, _ in group]
            result.extend(_xycut(group_indices, regions))
        return result
    boxes.sort(key=lambda b: (b[1][1], b[1][0]))
    return [idx for idx, _ in boxes]

def _split_axis(boxes: List[tuple], axis: str, gap_ratio: float=0.02) -> List[List[tuple]]:
    """
    Разбивает bbox по одной оси при наличии разрывов.

    Проецирует bbox на ось, ищет промежутки между проекциями.
    Порог разрыва = gap_ratio * общий диапазон оси.

    Args:
        boxes: список (index, [x1,y1,x2,y2]).
        axis: "x" или "y".
        gap_ratio: минимальный разрыв как доля от диапазона.

    Returns:
        Список групп. Каждая группа — список (index, bbox).
    """
    if axis == 'y':
        projections = [(b[1][1], b[1][3], b) for b in boxes]
    else:
        projections = [(b[1][0], b[1][2], b) for b in boxes]
    projections.sort(key=lambda p: p[0])
    total_min = min((p[0] for p in projections))
    total_max = max((p[1] for p in projections))
    total_range = total_max - total_min
    if total_range <= 0:
        return [boxes]
    min_gap = total_range * gap_ratio
    groups = []
    current_group = [projections[0][2]]
    current_end = projections[0][1]
    for i in range(1, len(projections)):
        start_i = projections[i][0]
        end_i = projections[i][1]
        if start_i - current_end > min_gap:
            groups.append(current_group)
            current_group = [projections[i][2]]
            current_end = end_i
        else:
            current_group.append(projections[i][2])
            current_end = max(current_end, end_i)
    groups.append(current_group)
    return groups

paddle_parser_reading_order_module = SimpleNamespace(logger=paddle_parser_reading_order_logger, sort_by_reading_order=sort_by_reading_order, _xycut=_xycut, _split_axis=_split_axis)

# END_SOURCE_PADDLE_PARSER_READING_ORDER

# START_SOURCE_PADDLE_PARSER_TABLE_PARSER
# PURPOSE: Inlined source from audit_engine/paddle_parser/table_parser.py.
paddle_parser_table_parser_logger = logging.getLogger(__name__)

@dataclass
class Cell:
    """Одна ячейка таблицы."""
    text: str = ''
    cell_type: str = 'empty'
    colspan: int = 1
    rowspan: int = 1
    covered: bool = False

def _tokenize_row(row_str: str) -> List[Cell]:
    """
    Парсит строку PaddleOCR-VL в список Cell с сохранением типа тега.

    Args:
        row_str: строка вида "<fcel>Текст<ecel><lcel><fcel>Текст2".

    Returns:
        Список Cell.
    """
    cells = []
    tokens = re.split('(<fcel>|<ecel>|<lcel>|<ucel>)', row_str)
    current_type = None
    current_text = ''
    for token in tokens:
        if token in ('<fcel>', '<ecel>', '<lcel>', '<ucel>'):
            if current_type is not None:
                cells.append(Cell(text=current_text.strip(), cell_type=current_type))
            type_map = {'<fcel>': 'filled', '<ecel>': 'empty', '<lcel>': 'lcel', '<ucel>': 'ucel'}
            current_type = type_map[token]
            current_text = ''
        else:
            current_text += token
    if current_type is not None:
        cells.append(Cell(text=current_text.strip(), cell_type=current_type))
    return cells

def _build_grid(raw: str) -> Optional[List[List[Cell]]]:
    """
    Парсит сырой PaddleOCR-VL текст в сетку Cell с colspan/rowspan.

    Обрабатывает <lcel> (увеличивает colspan предыдущей fcel) и
    <ucel> (увеличивает rowspan ячейки сверху).

    Args:
        raw: сырой текст с тегами.

    Returns:
        Сетка Cell или None если формат не распознан.
    """
    if '<fcel>' not in raw and '<ecel>' not in raw:
        return None
    rows_raw = re.split('<nl>', raw)
    rows_raw = [r.strip() for r in rows_raw if r.strip()]
    if not rows_raw:
        return None
    grid: List[List[Cell]] = []
    for row_str in rows_raw:
        cells = _tokenize_row(row_str)
        if cells:
            grid.append(cells)
    if not grid:
        return None
    for row in grid:
        i = 0
        while i < len(row):
            if row[i].cell_type == 'lcel':
                for j in range(i - 1, -1, -1):
                    if row[j].cell_type != 'lcel':
                        row[j].colspan += 1
                        row[i].covered = True
                        break
            i += 1
    max_logical_cols = 0
    for row in grid:
        width = sum((c.colspan for c in row if not c.covered))
        max_logical_cols = max(max_logical_cols, width)
    position_map: List[List[Optional[Cell]]] = []
    for row_idx, row in enumerate(grid):
        pos_row: List[Optional[Cell]] = [None] * max_logical_cols
        col = 0
        for cell in row:
            if cell.covered:
                continue
            while col < max_logical_cols and pos_row[col] is not None:
                col += 1
            if col >= max_logical_cols:
                break
            if cell.cell_type == 'ucel':
                cell.covered = True
                for prev_row_idx in range(row_idx - 1, -1, -1):
                    if prev_row_idx < len(position_map):
                        source = position_map[prev_row_idx][col]
                        if source is not None and (not source.covered):
                            source.rowspan += 1
                            break
                        elif source is not None and source.covered:
                            break
                col += 1
            else:
                pos_row[col] = cell
                for cs in range(1, cell.colspan):
                    if col + cs < max_logical_cols:
                        pos_row[col + cs] = cell
                col += cell.colspan
        if position_map:
            prev = position_map[-1]
            for c in range(max_logical_cols):
                if pos_row[c] is None and prev[c] is not None:
                    src = prev[c]
                    if not src.covered and src.rowspan > 1:
                        remaining = src.rowspan - (row_idx - _find_origin_row(position_map, src))
                        if remaining > 0:
                            pos_row[c] = src
        position_map.append(pos_row)
    n_rows = len(grid)
    n_cols = max_logical_cols
    paddle_parser_table_parser_logger.info(f'Таблица: {n_rows} строк × {n_cols} колонок')
    return grid

def _find_origin_row(position_map: List[List[Optional[Cell]]], cell: Cell) -> int:
    """Находит номер строки, где ячейка впервые появляется."""
    for row_idx, row in enumerate(position_map):
        if cell in row:
            return row_idx
    return 0

def parse_paddle_table_html(raw: str) -> str:
    """
    Конвертирует табличный формат PaddleOCR-VL в HTML-таблицу.

    Сохраняет colspan и rowspan.

    Args:
        raw: сырой текст с тегами <fcel>, <ecel>, <lcel>, <ucel>, <nl>.

    Returns:
        HTML-таблица. Если формат не распознан, возвращает сырой текст.
    """
    grid = _build_grid(raw)
    if grid is None:
        return raw
    lines = ['<table>']
    for row in grid:
        lines.append('<tr>')
        for cell in row:
            if cell.covered:
                continue
            attrs = ''
            if cell.colspan > 1:
                attrs += f' colspan="{cell.colspan}"'
            if cell.rowspan > 1:
                attrs += f' rowspan="{cell.rowspan}"'
            text = cell.text if cell.text else ''
            lines.append(f'  <td{attrs}>{text}</td>')
        lines.append('</tr>')
    lines.append('</table>')
    return '\n'.join(lines)

def parse_paddle_table(raw: str) -> str:
    """
    Конвертирует табличный формат PaddleOCR-VL в markdown-таблицу.

    Markdown не поддерживает colspan/rowspan — merged cells становятся пустыми.

    Args:
        raw: сырой текст с тегами <fcel>, <ecel>, <nl>.

    Returns:
        Markdown-таблица. Если формат не распознан, возвращает сырой текст.
    """
    grid = _build_grid(raw)
    if grid is None:
        return raw
    max_cols = 0
    text_rows: List[List[str]] = []
    for row in grid:
        texts = []
        for cell in row:
            if cell.covered:
                texts.append('')
            else:
                texts.append(cell.text)
                for _ in range(cell.colspan - 1):
                    texts.append('')
        max_cols = max(max_cols, len(texts))
        text_rows.append(texts)
    for row in text_rows:
        while len(row) < max_cols:
            row.append('')
    lines = []
    header = '| ' + ' | '.join(text_rows[0]) + ' |'
    lines.append(header)
    sep = '|' + '|'.join(['---'] * max_cols) + '|'
    lines.append(sep)
    for row in text_rows[1:]:
        line = '| ' + ' | '.join(row) + ' |'
        lines.append(line)
    return '\n'.join(lines)

def is_table_format(text: str) -> bool:
    """
    Определяет, содержит ли текст табличный формат PaddleOCR-VL.

    Args:
        text: текст от VLM.

    Returns:
        True если текст в табличном формате.
    """
    return '<fcel>' in text or ('<ecel>' in text and '<nl>' in text) or '<lcel>' in text

paddle_parser_table_parser_module = SimpleNamespace(logger=paddle_parser_table_parser_logger, Cell=Cell, _tokenize_row=_tokenize_row, _build_grid=_build_grid, _find_origin_row=_find_origin_row, parse_paddle_table_html=parse_paddle_table_html, parse_paddle_table=parse_paddle_table, is_table_format=is_table_format)

# END_SOURCE_PADDLE_PARSER_TABLE_PARSER

# START_SOURCE_PADDLE_PARSER_NORMALIZER
# PURPOSE: Inlined source from audit_engine/paddle_parser/normalizer.py.
paddle_parser_normalizer_logger = logging.getLogger(__name__)

def build_page_markdown(recognized_regions: List[Dict[str, Any]], page_num: int=0) -> str:
    """
    Собирает markdown из распознанных регионов одной страницы.

    Фильтрует регионы без промптов и пустые. Таблицы конвертируются
    из формата PaddleOCR-VL (<fcel>/<nl>) в HTML. Каждый блок
    оборачивается в markdown по CLASS_MD_WRAPPER.

    Args:
        recognized_regions: список словарей:
            [{"class_name": str, "text": str, "bbox": list, "score": float}, ...]
            Уже отсортированный по reading order.
        page_num: номер страницы (для логирования).

    Returns:
        str: markdown-текст страницы.
    """
    parts = []
    for region in recognized_regions:
        cls = region['class_name']
        text = region.get('text', '').strip()
        if paddle_parser_config_module.CLASS_PROMPTS.get(cls) is None:
            paddle_parser_normalizer_logger.debug(f'Пропуск: {cls} (нет промпта)')
            continue
        if not text:
            paddle_parser_normalizer_logger.debug(f'Пропуск: {cls} (пустой текст)')
            continue
        if cls == 'Table' and is_table_format(text):
            if paddle_parser_config_module.TABLE_OUTPUT_FORMAT == 'html':
                text = parse_paddle_table_html(text)
            else:
                text = parse_paddle_table(text)
        wrapper = paddle_parser_config_module.CLASS_MD_WRAPPER.get(cls, '{text}')
        md_block = wrapper.format(text=text)
        parts.append(md_block)
    result = '\n\n'.join(parts)
    paddle_parser_normalizer_logger.info(f'Страница {page_num}: {len(parts)} блоков, {len(result)} символов markdown')
    return result

def merge_text_and_tables(full_page_text: str, table_entries: list, page_size: tuple) -> str:
    """
    Вставляет HTML-таблицы в текст страницы по y-позиции на странице.

    Таблицы были замаскированы белым при full-page OCR, поэтому их нет
    в full_page_text. Здесь мы вставляем их обратно в правильное место.

    Args:
        full_page_text: текст страницы от VLM (без таблиц).
        table_entries: [{"bbox": [x1,y1,x2,y2], "html": str}],
                       отсортированные по y (сверху вниз).
        page_size: (width, height) страницы в пикселях.

    Returns:
        str: текст с вставленными HTML-таблицами, нормализованный.
    """
    if not table_entries:
        return normalize_paddle_text(full_page_text)
    _, page_h = page_size
    lines = full_page_text.split('\n')
    total_lines = len(lines)
    if total_lines == 0:
        parts = [t['html'] for t in table_entries]
        return '\n\n'.join(parts)
    insertions = []
    for entry in table_entries:
        y1 = entry['bbox'][1]
        y2 = entry['bbox'][3]
        y_center = (y1 + y2) / 2
        relative_pos = y_center / page_h
        insert_at = int(relative_pos * total_lines)
        insert_at = max(0, min(insert_at, total_lines))
        insertions.append((insert_at, entry['html']))
    insertions.sort(key=lambda x: x[0], reverse=True)
    for line_idx, html in insertions:
        lines.insert(line_idx, f'\n{html}\n')
    result = '\n'.join(lines)
    return normalize_paddle_text(result)

def normalize_paddle_text(text: str) -> str:
    """
    Нормализация OCR-текста от Paddle pipeline для ChunkAssembler.

    Отличия от normalize_ocr_text (HunyuanOCR):
    - НЕ вызывает unwrap_latex_artifacts() — PaddleOCR-VL не генерирует LaTeX-артефакты
    - НЕ вызывает html_table_to_markdown() — таблицы уже в HTML из build_page_markdown,
      и мы СОХРАНЯЕМ HTML с colspan/rowspan для LLM-аудита

    Операции:
    1. Множественные пустые строки -> одна
    2. Пробелы в конце строк

    Args:
        text: текст страницы (может содержать HTML-таблицы)

    Returns:
        str: нормализованный текст
    """
    if not text:
        return ''
    text = re.sub('\\n{3,}', '\n\n', text)
    text = re.sub(' +\\n', '\n', text)
    return text.strip()

paddle_parser_normalizer_module = SimpleNamespace(logger=paddle_parser_normalizer_logger, build_page_markdown=build_page_markdown, merge_text_and_tables=merge_text_and_tables, normalize_paddle_text=normalize_paddle_text)

# END_SOURCE_PADDLE_PARSER_NORMALIZER

# START_SOURCE_PADDLE_PARSER_VLM_CLIENT
# PURPOSE: Inlined source from audit_engine/paddle_parser/vlm_client.py.
paddle_parser_vlm_client_logger = logging.getLogger(__name__)

class VLMClient:
    """
    Клиент к VLM-серверу (PaddleOCR-VL-1.5 через vLLM).

    Args:
        base_url: URL vLLM-сервера.
        model_name: имя модели (None = автоопределение через /v1/models).
        api_key: API-ключ (vLLM обычно не требует).
        max_tokens: максимум токенов в ответе.
        temperature: температура генерации.
    """

    def __init__(self, base_url: str='http://localhost:8010/v1', model_name: Optional[str]=None, api_key: Optional[str]=None, max_tokens: Optional[int]=None, temperature: Optional[float]=None):
        self.base_url = base_url
        self.max_tokens = max_tokens or paddle_parser_config_module.VLM_MAX_TOKENS
        self.temperature = temperature if temperature is not None else paddle_parser_config_module.VLM_TEMPERATURE
        self.client = AsyncOpenAI(base_url=self.base_url, api_key=api_key or paddle_parser_config_module.VLM_API_KEY)
        self.model_name = model_name or self._detect_model()

    def _detect_model(self) -> str:
        """
        Автоопределение модели на vLLM через /v1/models.

        Returns:
            str: имя модели

        Raises:
            ConnectionError: если vLLM не отвечает
        """
        from openai import OpenAI
        sync_client = OpenAI(base_url=self.base_url, api_key=paddle_parser_config_module.VLM_API_KEY)
        try:
            models = sync_client.models.list()
            if models.data:
                model_id = models.data[0].id
                paddle_parser_vlm_client_logger.info(f'VLM модель (авто): {model_id}')
                return model_id
            raise ConnectionError('vLLM не вернул ни одной модели')
        except Exception as e:
            raise ConnectionError(f'Не удалось подключиться к VLM ({self.base_url}): {e}')

    async def recognize(self, image: Image.Image, prompt: str) -> Dict[str, Any]:
        """
        Отправляет изображение + промпт в VLM и возвращает текст.

        Args:
            image: PIL Image (crop региона).
            prompt: текстовый промпт (напр. "OCR:", "Table Recognition:").

        Returns:
            {"text": str, "tokens": int, "time_sec": float}
        """
        b64 = self._image_to_base64(image)
        content = [{'type': 'image_url', 'image_url': {'url': f'data:image/png;base64,{b64}'}}, {'type': 'text', 'text': prompt}]
        t0 = time.time()
        response = await self.client.chat.completions.create(model=self.model_name, messages=[{'role': 'user', 'content': content}], max_tokens=self.max_tokens, temperature=self.temperature)
        elapsed = time.time() - t0
        text = response.choices[0].message.content or ''
        tokens = response.usage.completion_tokens if response.usage else 0
        if elapsed > 0:
            paddle_parser_vlm_client_logger.info(f'VLM ответ: {tokens} tok за {elapsed:.1f}s ({tokens / elapsed:.0f} tok/s)')
        return {'text': text.strip(), 'tokens': tokens, 'time_sec': round(elapsed, 2)}

    async def recognize_batch(self, items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Параллельное распознавание нескольких регионов.

        Args:
            items: список [{"image": PIL.Image, "prompt": str, "index": int}, ...]

        Returns:
            Список результатов [{"text": str, "tokens": int, "time_sec": float, "index": int}, ...]
        """
        tasks = []
        for item in items:
            tasks.append(self._recognize_with_index(item))
        results = await asyncio.gather(*tasks, return_exceptions=True)
        output = []
        for i, result in enumerate(results):
            if isinstance(result, Exception):
                paddle_parser_vlm_client_logger.error(f'Ошибка VLM для региона {items[i].get('index', i)}: {result}')
                output.append({'text': f'[ОШИБКА: {result}]', 'tokens': 0, 'time_sec': 0, 'index': items[i].get('index', i)})
            else:
                output.append(result)
        return output

    async def _recognize_with_index(self, item: Dict[str, Any]) -> Dict[str, Any]:
        """Обёртка для сохранения index в результате."""
        result = await self.recognize(item['image'], item['prompt'])
        result['index'] = item.get('index', 0)
        return result

    @staticmethod
    def _image_to_base64(image: Image.Image) -> str:
        """
        Конвертирует PIL Image в base64 строку.

        Args:
            image: PIL Image.

        Returns:
            Base64 строка.
        """
        buf = io.BytesIO()
        image.save(buf, format='PNG')
        buf.seek(0)
        return base64.b64encode(buf.read()).decode('utf-8')

paddle_parser_vlm_client_module = SimpleNamespace(logger=paddle_parser_vlm_client_logger, VLMClient=VLMClient)

# END_SOURCE_PADDLE_PARSER_VLM_CLIENT

# START_SOURCE_PADDLE_PARSER_EXTRACTOR
# PURPOSE: Inlined source from audit_engine/paddle_parser/extractor.py.
paddle_parser_extractor_logger = logging.getLogger(__name__)

class PaddleExtractor:
    """
    Layout-aware OCR экстрактор: Heron-101 + PaddleOCR-VL-1.5.

    Модели загружаются лениво при первом вызове extract_pages().

    Args:
        vlm_base_url: URL vLLM-сервера PaddleOCR-VL-1.5.
        vlm_model: имя VLM-модели (None = автоопределение).
        layout_model_repo: HF repo для Heron-101 (None = дефолт).
        layout_device: GPU для layout detection.
        dpi: разрешение PDF → PNG.
        log_dir: директория для логов (визуализации, crop'ы, raw.json).
    """

    def __init__(self, vlm_base_url: str='http://localhost:8010/v1', vlm_model: Optional[str]=None, layout_model_repo: Optional[str]=None, layout_device: str='cuda:0', layout_base_url: Optional[str]=None, dpi: int=200, log_dir: Optional[str]=None):
        self.vlm_base_url = vlm_base_url
        self.vlm_model = vlm_model
        self.layout_model_repo = layout_model_repo
        self.layout_device = layout_device
        self.layout_base_url = layout_base_url
        self.dpi = dpi
        self.max_page_pixels = 4000
        self.log_dir = Path(log_dir) if log_dir else None
        self._detector = None
        self._vlm: Optional[VLMClient] = None

    def _ensure_models(self) -> None:
        """Загружает модели если ещё не загружены."""
        if self._detector is None:
            if self.layout_base_url:
                paddle_parser_extractor_logger.info(f'Использую RemoteLayoutDetector: {self.layout_base_url}')
                self._detector = RemoteLayoutDetector(base_url=self.layout_base_url)
            else:
                paddle_parser_extractor_logger.info('Загрузка Heron-101 layout detector...')
                self._detector = LayoutDetector(model_name=self.layout_model_repo, device=self.layout_device)
        if self._vlm is None:
            paddle_parser_extractor_logger.info('Инициализация VLM-клиента...')
            self._vlm = VLMClient(base_url=self.vlm_base_url, model_name=self.vlm_model)

    @staticmethod
    def _is_blank_page(image: Image.Image, threshold: float=0.995) -> bool:
        """
        Проверяет пустая ли страница (почти полностью белая).

        Args:
            image: PIL Image страницы
            threshold: доля белых пикселей для признания пустой (0.995 = 99.5%)

        Returns:
            bool: True если страница пустая
        """
        arr = np.array(image)
        white_ratio = (arr >= 250).sum() / arr.size
        return white_ratio >= threshold

    def _render_page(self, doc: fitz.Document, page_idx: int, target_dpi: int=72) -> Image.Image:
        """Рендер одной страницы PDF через pymupdf с ограничением max_page_pixels."""
        page = doc[page_idx]
        scale = target_dpi / 72.0
        long_side = max(page.rect.width, page.rect.height) * scale
        if long_side > self.max_page_pixels:
            scale = self.max_page_pixels / max(page.rect.width, page.rect.height)
        mat = fitz.Matrix(scale, scale)
        pix = page.get_pixmap(matrix=mat)
        return Image.frombytes('RGB', [pix.width, pix.height], pix.samples)

    def detect_blank_pages(self, pdf_path: str) -> Tuple[int, Set[int]]:
        """
        Определяет пустые страницы в PDF (быстрый рендер, pymupdf).

        Args:
            pdf_path: путь к PDF

        Returns:
            (total_pages, blank_indices): общее число страниц и множество
            0-based индексов пустых страниц
        """
        doc = fitz.open(pdf_path)
        total = len(doc)
        blanks = set()
        for i in range(total):
            img = self._render_page(doc, i, target_dpi=72)
            if self._is_blank_page(img):
                blanks.add(i)
                paddle_parser_extractor_logger.info(f'  стр.{i + 1}: пустая (пропускаем)')
        if blanks:
            paddle_parser_extractor_logger.info(f'Пустых страниц: {len(blanks)} из {total}')
        return (total, blanks)

    def extract_pages(self, pdf_path: str, page_indices: List[int]) -> Dict[int, str]:
        """
        Layout-aware OCR указанных страниц PDF.

        Каждая страница: layout detection → crop → VLM OCR → markdown.

        Args:
            pdf_path: путь к PDF файлу
            page_indices: список 0-based индексов страниц для OCR

        Returns:
            Dict[int, str]: {1-based номер страницы: markdown-текст}
        """
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f'PDF не найден: {pdf_path}')
        self._ensure_models()
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        if total_pages == 0:
            raise ValueError('PDF не содержит страниц')
        unique_indices = sorted(set((i for i in page_indices if 0 <= i < total_pages)))
        if not unique_indices:
            paddle_parser_extractor_logger.warning(f'Нет валидных страниц из {page_indices} (всего {total_pages})')
            return {}
        paddle_parser_extractor_logger.info(f'Paddle OCR: {len(unique_indices)} страниц из {total_pages}')
        page_texts = asyncio.run(self._process_pages_from_pdf(doc, unique_indices))
        return page_texts

    async def _process_pages_from_pdf(self, doc: fitz.Document, page_indices: List[int]) -> Dict[int, str]:
        """Постраничный рендер (pymupdf) + обработка."""
        page_texts = {}
        for idx in page_indices:
            page_num = idx + 1
            try:
                img = self._render_page(doc, idx, target_dpi=self.dpi)
                md = await self._process_single_page(img, page_num)
                page_texts[page_num] = md
            except Exception as e:
                paddle_parser_extractor_logger.error(f'Ошибка Paddle OCR стр.{page_num}: {e}')
                page_texts[page_num] = f'[ОШИБКА Paddle OCR: {e}]'
        return page_texts

    async def _process_pages_async(self, images: List[Image.Image], page_indices: List[int]) -> Dict[int, str]:
        """
        Последовательная обработка страниц (layout → VLM OCR внутри страницы параллельный).

        Args:
            images: все изображения PDF
            page_indices: 0-based индексы страниц

        Returns:
            Dict[int, str]: {1-based page_num: markdown}
        """
        page_texts = {}
        for idx in page_indices:
            page_num = idx + 1
            try:
                md = await self._process_single_page(images[idx], page_num)
                page_texts[page_num] = md
            except Exception as e:
                paddle_parser_extractor_logger.error(f'Ошибка Paddle OCR стр.{page_num}: {e}')
                page_texts[page_num] = f'[ОШИБКА Paddle OCR: {e}]'
        return page_texts

    async def _process_single_page(self, page_image: Image.Image, page_num: int) -> str:
        """
        Обрабатывает одну страницу: full-page OCR + отдельные кропы таблиц.

        Вместо кропа каждого layout-элемента отправляем полную страницу
        в VLM с промптом "OCR:". Таблицы маскируются белым и кропаются
        отдельно с промптом "Table Recognition:". Результаты склеиваются
        по y-координатам.

        Args:
            page_image: изображение страницы.
            page_num: номер страницы (1-based).

        Returns:
            str: текст страницы с вставленными HTML-таблицами.
        """
        paddle_parser_extractor_logger.info(f'=== Страница {page_num} ===')
        t0 = time.time()
        regions = self._detector.detect(page_image)
        layout_time = time.time() - t0
        class_counts = Counter((r['class_name'] for r in regions))
        paddle_parser_extractor_logger.info(f'Layout: {len(regions)} регионов за {layout_time:.2f}s | {dict(class_counts)}')
        page_w, page_h = page_image.size
        page_area = page_w * page_h
        min_area = page_area * paddle_parser_config_module.SMALL_ELEMENT_MIN_AREA_RATIO
        filtered = []
        for r in regions:
            bbox = r['bbox']
            box_area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
            if r['class_name'] in paddle_parser_config_module.SMALL_ELEMENT_CLASSES and box_area < min_area:
                paddle_parser_extractor_logger.debug(f'  Пропуск мелкого {r['class_name']}: area={box_area:.0f} < {min_area:.0f}')
                continue
            filtered.append(r)
        if len(filtered) != len(regions):
            paddle_parser_extractor_logger.info(f'Отфильтровано мелких: {len(regions) - len(filtered)}')
        regions = filtered
        if self.log_dir:
            self.log_dir.mkdir(parents=True, exist_ok=True)
            vis_path = self.log_dir / f'page_{page_num}_layout.png'
            self._detector.visualize(page_image, regions, str(vis_path))
        table_regions = [r for r in regions if r['class_name'] == 'Table']
        table_regions.sort(key=lambda r: r['bbox'][1])
        n_tables = len(table_regions)
        paddle_parser_extractor_logger.info(f'Таблиц: {n_tables}, full-page OCR + {n_tables} table crop(s)')
        vlm_items = []
        if table_regions:
            ocr_page = self._mask_table_regions(page_image, table_regions)
        else:
            ocr_page = page_image
        vlm_items.append({'image': ocr_page, 'prompt': 'OCR:', 'index': -1})
        for i, table_reg in enumerate(table_regions):
            crop = self._crop_region(page_image, table_reg['bbox'])
            vlm_items.append({'image': crop, 'prompt': 'Table Recognition:', 'index': i})
        if self.log_dir:
            if table_regions:
                masked_path = self.log_dir / f'page_{page_num}_masked.png'
                ocr_page.save(str(masked_path))
            for i in range(n_tables):
                crop_path = self.log_dir / f'page_{page_num}_table_{i}.png'
                vlm_items[1 + i]['image'].save(str(crop_path))
        t0 = time.time()
        vlm_results = await self._vlm.recognize_batch(vlm_items)
        vlm_time = time.time() - t0
        paddle_parser_extractor_logger.info(f'VLM: {len(vlm_results)} запросов за {vlm_time:.2f}s')
        vlm_map = {r['index']: r['text'] for r in vlm_results}
        full_page_text = vlm_map.get(-1, '')
        table_entries = []
        for i, table_reg in enumerate(table_regions):
            raw_table = vlm_map.get(i, '')
            if is_table_format(raw_table):
                html = parse_paddle_table_html(raw_table)
            else:
                html = raw_table
            table_entries.append({'bbox': table_reg['bbox'], 'html': html})
        if self.log_dir:
            raw_data = {'full_page_text': full_page_text, 'tables': [{'bbox': [int(c) for c in table_regions[i]['bbox']], 'score': round(table_regions[i]['score'], 3), 'raw_text': vlm_map.get(i, ''), 'parsed_html': entry['html']} for i, entry in enumerate(table_entries)]}
            raw_path = self.log_dir / f'page_{page_num}_raw.json'
            with open(raw_path, 'w', encoding='utf-8') as f:
                json.dump(raw_data, f, ensure_ascii=False, indent=2)
        md = merge_text_and_tables(full_page_text, table_entries, page_image.size)
        return md

    @staticmethod
    def _mask_table_regions(page_image: Image.Image, table_regions: list, padding: int=None) -> Image.Image:
        """
        Закрашивает области таблиц белым на копии страницы.

        Позволяет full-page OCR не читать содержимое таблиц
        (они распознаются отдельными кропами с "Table Recognition:").

        Args:
            page_image: оригинал страницы.
            table_regions: регионы с class_name="Table".
            padding: отступ маски в пикселях.

        Returns:
            Копия страницы с замаскированными таблицами.
        """
        from PIL import ImageDraw
        padding = padding if padding is not None else paddle_parser_config_module.TABLE_MASK_PADDING_PX
        masked = page_image.copy()
        draw = ImageDraw.Draw(masked)
        for region in table_regions:
            x1, y1, x2, y2 = [int(c) for c in region['bbox']]
            x1 = max(0, x1 - padding)
            y1 = max(0, y1 - padding)
            x2 = min(masked.width, x2 + padding)
            y2 = min(masked.height, y2 + padding)
            draw.rectangle([x1, y1, x2, y2], fill='white')
        return masked

    @staticmethod
    def _crop_region(page_image: Image.Image, bbox: list, padding: int=None) -> Image.Image:
        """
        Вырезает регион из изображения страницы с padding.

        Args:
            page_image: изображение страницы.
            bbox: [x1, y1, x2, y2].
            padding: отступ в пикселях.

        Returns:
            Вырезанное изображение.
        """
        padding = padding if padding is not None else paddle_parser_config_module.CROP_PADDING_PX
        x1, y1, x2, y2 = [int(c) for c in bbox]
        w, h = page_image.size
        x1c = max(0, x1 - padding)
        y1c = max(0, y1 - padding)
        x2c = min(w, x2 + padding)
        y2c = min(h, y2 + padding)
        return page_image.crop((x1c, y1c, x2c, y2c))

paddle_parser_extractor_module = SimpleNamespace(logger=paddle_parser_extractor_logger, PaddleExtractor=PaddleExtractor)

# END_SOURCE_PADDLE_PARSER_EXTRACTOR

# START_SOURCE_PADDLE_PARSER
# PURPOSE: Inlined source from audit_engine/paddle_parser/__init__.py.
paddle_parser_log = logging.getLogger(__name__)

paddle_parser___all__ = ['PaddleParser']

class PaddleParser:
    """
    Главный класс Paddle OCR Pipeline.

    Объединяет:
    1. DocumentConverter — DOCX → PDF
    2. PaddleExtractor — PDF → layout + VLM OCR → постраничный markdown
    3. ChunkAssembler — сборка чанков из страниц по конфигу
    4. ChunkAggregator — добавление метаданных (имя файла, путь)

    Возвращаемое значение: Dict[str, Any] с метаданными файла и чанками.
    """

    def __init__(self, config: AuditConfig, log_dir: Optional[str]=None):
        """
        Инициализация Paddle-парсера.

        Args:
            config: конфигурация аудита (содержит paddle_* и ocr_* параметры)
            log_dir: директория для логов (layout-визуализации, crop'ы, raw.json)
        """
        self.config = config
        self.chunks = load_chunks_config(str(config.chunks_vision_path))
        self.converter = DocumentConverter()
        self.extractor = PaddleExtractor(vlm_base_url=config.ocr_base_url, vlm_model=config.paddle_vlm_model, layout_model_repo=config.paddle_layout_model, layout_device=config.paddle_layout_device, layout_base_url=config.paddle_layout_base_url, dpi=config.ocr_dpi, log_dir=log_dir)
        self.assembler = ChunkAssembler()
        self.aggregator = ChunkAggregator()
        self._temp_pdf: Optional[str] = None

    def parse(self, file_path: str, chunk_filter: Optional[str]=None) -> Dict[str, Any]:
        """
        Парсинг документа через Paddle OCR Pipeline.

        Args:
            file_path: путь к документу (DOCX, PDF)
            chunk_filter: имя чанка (если указан — собираем только его)

        Returns:
            Dict[str, Any] в формате legacy-парсера:
            {
                "filename": str,
                "path": str,
                "шапка": str,
                "текст_приказа": str,
                ...
            }

        Raises:
            FileNotFoundError: файл не найден
            ConnectionError: PaddleOCR-VL vLLM недоступен
        """
        input_path = Path(file_path)
        if not input_path.exists():
            raise FileNotFoundError(f'Файл не найден: {file_path}')
        try:
            pdf_path = self.converter.convert(str(input_path))
            if pdf_path != str(input_path.absolute()):
                self._temp_pdf = pdf_path
            chunks = self.chunks
            if chunk_filter:
                chunks = [c for c in self.chunks if c.name == chunk_filter]
                if not chunks:
                    raise ValueError(f"Чанк '{chunk_filter}' не найден в конфигурации")
            total_pages, blank_pages = self.extractor.detect_blank_pages(pdf_path)
            non_blank = [i for i in range(total_pages) if i not in blank_pages]
            effective_total = max(non_blank) + 1 if non_blank else total_pages
            if effective_total != total_pages:
                paddle_parser_log.info(f'Страниц в PDF: {total_pages}, эффективных: {effective_total} (пустые: {sorted((i + 1 for i in blank_pages))})')
            unique_page_indices = collect_unique_pages(chunks, effective_total)
            unique_page_indices = [i for i in unique_page_indices if i not in blank_pages]
            paddle_parser_log.info(f'Уникальных страниц для Paddle OCR: {len(unique_page_indices)} из {total_pages}')
            page_texts = self.extractor.extract_pages(pdf_path, unique_page_indices)
            chunk_texts = self.assembler.assemble(page_texts, chunks, effective_total, strip_annotations_chunks=self.config.strip_annotations_chunks, normalizer_fn=normalize_paddle_text)
            result = self.aggregator.aggregate(chunk_texts, str(input_path))
            return result
        finally:
            self._cleanup()

    def _cleanup(self):
        """Очистка временных файлов."""
        if self._temp_pdf:
            self.converter.cleanup(self._temp_pdf)
            self._temp_pdf = None

paddle_parser_module = SimpleNamespace(log=paddle_parser_log, __all__=paddle_parser___all__, PaddleParser=PaddleParser)

# END_SOURCE_PADDLE_PARSER

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

# START_SOURCE_PARSERS_GRAFIK_OBHOD
# PURPOSE: Inlined source from audit_engine/parsers/grafik_obhod.py.
def parsers_grafik_obhod_parse(xlsx_path: str) -> dict:
    """
    Парсит документ "График обхода ОМ".

    Args:
        xlsx_path: путь к .xlsx файлу

    Returns:
        dict с чанками: шапка, фио_должности (БЕЗ префиксов)
    """
    wb = load_workbook(xlsx_path, data_only=True)
    sheet_name = None
    for name in wb.sheetnames:
        if name.lower() in ['график', 'шаблон']:
            sheet_name = name
            break
    if not sheet_name:
        sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]
    company_name = ''
    prikaz_ref = ''
    employees = []
    for row_num in range(1, min(51, ws.max_row + 1)):
        row_values = []
        for col_num in range(1, min(35, ws.max_column + 1)):
            cell = ws.cell(row=row_num, column=col_num)
            if cell.value:
                row_values.append(str(cell.value).strip())
        for val in row_values:
            if re.search('\\b(ООО|ЗАО|АО|ПАО|ИП)\\b', val) and (not company_name):
                company_name = val
        for val in row_values:
            if 'приложение' in val.lower() and 'приказ' in val.lower():
                prikaz_ref = val.replace('\n', ' ').strip()
        if len(row_values) >= 2:
            first_val = row_values[0]
            second_val = row_values[1]
            if first_val.lower() not in ['фио', 'график', ''] and 'график обхода' not in first_val.lower() and ('приложение' not in first_val.lower()):
                if re.search('[А-ЯЁ][а-яё]+\\s+[А-ЯЁ]\\.?[А-ЯЁа-яё]?\\.?', first_val) or re.search('[А-ЯЁ][а-яё]+\\s+[А-ЯЁ][а-яё]+', first_val):
                    if second_val and second_val.lower() != 'должность':
                        employees.append(f'{first_val} — {second_val}')
    chunks = {'шапка': f'{company_name}\n{prikaz_ref}'.strip(), 'фио_должности': '\n'.join(dict.fromkeys(employees))}
    return chunks

parsers_grafik_obhod_module = SimpleNamespace(parse=parsers_grafik_obhod_parse)

# END_SOURCE_PARSERS_GRAFIK_OBHOD

# START_SOURCE_DRIVERS_PARSER
# PURPOSE: Inlined source from audit_engine/drivers/parser.py.
MAIN_NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

RELS_NS = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}

CELL_REF_RE = re.compile('^([A-Z]+)(\\d+)$')

@dataclass(frozen=True)
class Row:
    index: int
    cells: Dict[str, str]

    def get(self, col: str, default: str='') -> str:
        return self.cells.get(col, default)

    @property
    def text(self) -> str:
        parts = [value.strip() for col, value in sorted(self.cells.items()) if value and value.strip()]
        return ' '.join(parts)

@dataclass(frozen=True)
class MergeRange:
    start_col: str
    start_row: int
    end_col: str
    end_row: int

    @property
    def start_col_index(self) -> int:
        return column_to_index(self.start_col)

    @property
    def end_col_index(self) -> int:
        return column_to_index(self.end_col)

@dataclass(frozen=True)
class ColumnLayout:
    respondents: Dict[str, str]
    total_column: Optional[str]
    total_label: Optional[str]
    average_column: Optional[str]
    average_label: Optional[str]

@dataclass
class Section:
    title: Optional[str]
    header_rows: List[Row]
    questions: List[Dict[str, object]]
    summary: Optional[str]
    column_layout: Optional[ColumnLayout] = None

def column_to_index(col: str) -> int:
    col = col.upper()
    result = 0
    for ch in col:
        if not 'A' <= ch <= 'Z':
            raise ValueError(f'Invalid column letter: {col}')
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result

def split_cell_reference(ref: str) -> tuple[str, int]:
    match = CELL_REF_RE.match(ref)
    if not match:
        raise ValueError(f'Unexpected cell reference: {ref}')
    col, row = match.groups()
    return (col, int(row))

def load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    try:
        data = zf.read('xl/sharedStrings.xml')
    except KeyError:
        return []
    root = ET.fromstring(data)
    strings: List[str] = []
    for si in root.findall('s:si', MAIN_NS):
        texts: List[str] = []
        for node in si.findall('.//s:t', MAIN_NS):
            texts.append(node.text or '')
        strings.append(''.join(texts))
    return strings

def resolve_sheet_path(zf: zipfile.ZipFile, preferred_name: Optional[str]=None) -> str:
    workbook = ET.fromstring(zf.read('xl/workbook.xml'))
    sheets = workbook.find('s:sheets', MAIN_NS)
    if sheets is None:
        raise RuntimeError('Workbook does not contain <sheets>')
    candidates: List[tuple[str, str, Optional[str]]] = []
    for sheet in sheets.findall('s:sheet', MAIN_NS):
        name = sheet.get('name')
        rel_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        state = sheet.get('state')
        if not rel_id:
            continue
        candidates.append((name or '', rel_id, state))
    if not candidates:
        raise RuntimeError('No sheets found in workbook')
    target_rel_id: Optional[str] = None
    if preferred_name:
        for name, rel_id, state in candidates:
            if name == preferred_name:
                target_rel_id = rel_id
                break
    if target_rel_id is None:
        for name, rel_id, state in candidates:
            if state != 'hidden':
                target_rel_id = rel_id
                break
    if target_rel_id is None:
        target_rel_id = candidates[0][1]
    rels_root = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
    for rel in rels_root.findall('rel:Relationship', RELS_NS):
        if rel.get('Id') == target_rel_id:
            target = rel.get('Target')
            if not target:
                break
            if not target.startswith('/'):
                return f'xl/{target}'
            return target.lstrip('/')
    raise RuntimeError(f'Failed to resolve target for relationship {target_rel_id}')

def parse_cell_value(cell: ET.Element, shared_strings: List[str]) -> Optional[str]:
    cell_type = cell.get('t')
    if cell_type == 's':
        value_node = cell.find('s:v', MAIN_NS)
        if value_node is None or value_node.text is None:
            return ''
        idx = int(value_node.text)
        return shared_strings[idx] if 0 <= idx < len(shared_strings) else ''
    if cell_type == 'inlineStr':
        texts = [node.text or '' for node in cell.findall('.//s:t', MAIN_NS)]
        return ''.join(texts)
    value_node = cell.find('s:v', MAIN_NS)
    if value_node is None or value_node.text is None:
        return ''
    return value_node.text

def read_rows(zf: zipfile.ZipFile, sheet_path: str, shared_strings: List[str]) -> tuple[List[Row], List[MergeRange]]:
    sheet_xml = zf.read(sheet_path)
    sheet_root = ET.fromstring(sheet_xml)
    merge_ranges: List[MergeRange] = []
    merge_cells = sheet_root.find('s:mergeCells', MAIN_NS)
    if merge_cells is not None:
        for merge in merge_cells.findall('s:mergeCell', MAIN_NS):
            ref = merge.get('ref')
            if not ref:
                continue
            start, end = ref.split(':')
            start_col, start_row = split_cell_reference(start)
            end_col, end_row = split_cell_reference(end)
            merge_ranges.append(MergeRange(start_col, start_row, end_col, end_row))
    rows: List[Row] = []
    sheet_data = sheet_root.find('s:sheetData', MAIN_NS)
    if sheet_data is None:
        return (rows, merge_ranges)
    for row_node in sheet_data.findall('s:row', MAIN_NS):
        idx = int(row_node.get('r'))
        cells: Dict[str, str] = {}
        for cell in row_node.findall('s:c', MAIN_NS):
            ref = cell.get('r')
            if not ref:
                continue
            col, _ = split_cell_reference(ref)
            value = parse_cell_value(cell, shared_strings)
            if value is None:
                continue
            value = value.replace('\r', '').strip()
            if value:
                cells[col] = value
        if cells:
            rows.append(Row(idx, cells))
    rows.sort(key=lambda r: r.index)
    return (rows, merge_ranges)

def detect_summary_rows(merges: Iterable[MergeRange]) -> set[int]:
    summary_starts = set()
    for mr in merges:
        if mr.start_col.upper() == 'B' and mr.end_col_index >= column_to_index('J'):
            summary_starts.add(mr.start_row)
    return summary_starts

def looks_like_question(cell_value: str) -> bool:
    trimmed = cell_value.strip()
    if not trimmed:
        return False
    return trimmed[0].isdigit()

def is_header_marker(row: Row) -> bool:
    number_cell = (row.get('B') or '').strip()
    if number_cell and looks_like_question(number_cell):
        return False
    for col in row.cells:
        if column_to_index(col) >= column_to_index('E'):
            return True
    return False

def normalize_header_value(value: str) -> str:
    cleaned = value.replace('\xa0', ' ').replace('\r', ' ').replace('\n', ' ')
    cleaned = cleaned.strip().rstrip(':')
    return ' '.join(cleaned.split())

def is_numeric_text(text: str) -> bool:
    if not text:
        return False
    normalized = text.replace(',', '.')
    try:
        float(normalized)
        return True
    except ValueError:
        return False

def detect_column_layout(header_rows: List[Row]) -> ColumnLayout:
    candidate_labels: Dict[str, str] = {}
    for row in header_rows:
        for col, raw_value in row.cells.items():
            if column_to_index(col) < column_to_index('E'):
                continue
            normalized = normalize_header_value(raw_value or '')
            if not normalized:
                continue
            lowered = normalized.lower().replace('ё', 'е')
            if 'балл' in lowered:
                continue
            if is_numeric_text(normalized):
                continue
            candidate_labels.setdefault(col, normalized)
    respondents: Dict[str, str] = {}
    total_column: Optional[str] = None
    total_label: Optional[str] = None
    average_column: Optional[str] = None
    average_label: Optional[str] = None
    for col, label in sorted(candidate_labels.items(), key=lambda item: column_to_index(item[0])):
        lowered = label.lower().replace('ё', 'е')
        if any((token in lowered for token in ('всего', 'итог', 'итого', 'overall', 'total', 'общее', 'суммар'))):
            total_column = col
            total_label = label
            continue
        if any((token in lowered for token in ('средн', 'average', 'avg', 'mean'))):
            average_column = col
            average_label = label
            continue
        respondents[col] = label
    return ColumnLayout(respondents=respondents, total_column=total_column, total_label=total_label, average_column=average_column, average_label=average_label)

def assemble_structure(rows: List[Row], summary_rows: set[int]) -> Dict[str, object]:
    meta_info: List[Dict[str, object]] = []
    sections: List[Section] = []
    current_section: Optional[Section] = None
    last_question: Optional[Dict[str, object]] = None
    for row in rows:
        idx = row.index
        if idx in summary_rows:
            if current_section is not None:
                summary_text = row.get('B') or row.get('C') or row.get('D')
                if summary_text:
                    current_section.summary = summary_text.strip()
            continue
        if row.get('E') == 'Баллы':
            if current_section is not None:
                sections.append(current_section)
            title = row.get('B') or None
            current_section = Section(title=title.strip() if title else None, header_rows=[row], questions=[], summary=None)
            last_question = None
            continue
        if current_section is None:
            text = row.text
            if text:
                meta_info.append({'row': idx, 'text': text})
            continue
        if current_section.title is None and row.get('B'):
            current_section.title = row.get('B').strip()
            current_section.header_rows.append(row)
            continue
        if is_header_marker(row):
            current_section.header_rows.append(row)
            continue
        number_cell = row.get('B')
        if number_cell and looks_like_question(number_cell):
            if current_section.column_layout is None:
                current_section.column_layout = detect_column_layout(current_section.header_rows)
            layout = current_section.column_layout
            scores: Dict[str, str] = {}
            for col, label in layout.respondents.items():
                scores[label] = row.get(col, '').strip()
            if layout.total_column:
                scores['Всего'] = row.get(layout.total_column, '').strip()
            if layout.average_column:
                scores['Средняя'] = row.get(layout.average_column, '').strip()
            question_entry = {'row': idx, 'number': number_cell.strip(), 'question': row.get('C').strip() if row.get('C') else '', 'problem_comment': row.get('D').strip() if row.get('D') else '', 'scores': scores, 'notes': []}
            current_section.questions.append(question_entry)
            last_question = question_entry
            continue
        note_text = row.get('C') or row.get('D')
        if note_text and last_question is not None:
            last_question['notes'].append({'row': idx, 'text': note_text.strip()})
            continue
    if current_section is not None:
        sections.append(current_section)
    serialisable_sections: List[Dict[str, object]] = []
    for section in sections:
        header_texts = [row.text for row in section.header_rows if row.text]
        serialisable_sections.append({'title': section.title, 'headers': header_texts, 'questions': section.questions, 'summary': section.summary})
    return {'meta': meta_info, 'sections': serialisable_sections}

def parse_excel_to_json(excel_path: Path, sheet_name: Optional[str]=None) -> Dict[str, object]:
    if excel_path.name.startswith('~$'):
        raise ValueError(f'Temporary Excel file detected: {excel_path.name}')
    try:
        with zipfile.ZipFile(excel_path) as zf:
            shared_strings = load_shared_strings(zf)
            sheet_path = resolve_sheet_path(zf, sheet_name)
            rows, merges = read_rows(zf, sheet_path, shared_strings)
    except zipfile.BadZipFile as exc:
        raise ValueError(f'{excel_path} is not a valid XLSX archive') from exc
    summary_rows = detect_summary_rows(merges)
    return assemble_structure(rows, summary_rows)

drivers_parser_module = SimpleNamespace(MAIN_NS=MAIN_NS, RELS_NS=RELS_NS, CELL_REF_RE=CELL_REF_RE, Row=Row, MergeRange=MergeRange, ColumnLayout=ColumnLayout, Section=Section, column_to_index=column_to_index, split_cell_reference=split_cell_reference, load_shared_strings=load_shared_strings, resolve_sheet_path=resolve_sheet_path, parse_cell_value=parse_cell_value, read_rows=read_rows, detect_summary_rows=detect_summary_rows, looks_like_question=looks_like_question, is_header_marker=is_header_marker, normalize_header_value=normalize_header_value, is_numeric_text=is_numeric_text, detect_column_layout=detect_column_layout, assemble_structure=assemble_structure, parse_excel_to_json=parse_excel_to_json)

# END_SOURCE_DRIVERS_PARSER

# START_SOURCE_PLAN_GRAFIK_PARSER
# PURPOSE: Inlined source from audit_engine/plan_grafik/parser.py.
def parse_plan_grafik(file_path: str) -> Dict[str, Any]:
    """
    Парсит XLSX план-графика.

    Args:
        file_path: путь к XLSX

    Returns:
        Dict со всеми извлечёнными данными
    """
    wb = load_workbook(file_path, data_only=False)
    sheets = wb.sheetnames
    result = {'filename': Path(file_path).name, 'sheets': sheets, 'approval': {}, 'dates': {}, 'title': '', 'headers_row16': {}, 'headers_row17': {}, 'data_rows': [], 'status_columns': {}, 'signature': '', 'formulas': {}}
    if 'План мероприятий' not in sheets:
        result['error'] = "Лист 'План мероприятий' не найден"
        return result
    ws = wb['План мероприятий']
    result['title'] = plan_grafik_parser__cell_str(ws, 'B', 8)
    dm_col = column_index_from_string('DM')
    result['approval'] = {'marker': _cell_str_by_idx(ws, dm_col, 8), 'position': _cell_str_by_idx(ws, dm_col, 9), 'company': _cell_str_by_idx(ws, dm_col, 10), 'fio': _cell_str_by_idx(ws, dm_col, 11), 'date_line': _cell_str_by_idx(ws, dm_col, 12)}
    i12 = ws.cell(row=12, column=column_index_from_string('I')).value
    i13 = ws.cell(row=13, column=column_index_from_string('I')).value
    result['dates'] = {'start': str(i12) if i12 else '', 'end': str(i13) if i13 else '', 'start_raw': i12, 'end_raw': i13}
    for col in range(1, 160):
        val = ws.cell(row=16, column=col).value
        if val and str(val).strip():
            letter = get_column_letter(col)
            result['headers_row16'][letter] = str(val).strip()
    for col in range(1, 160):
        val = ws.cell(row=17, column=col).value
        if val and str(val).strip():
            letter = get_column_letter(col)
            result['headers_row17'][letter] = str(val).strip()
    max_row = ws.max_row
    for row in range(18, min(max_row + 1, 500)):
        a_val = ws.cell(row=row, column=1).value
        j_val = ws.cell(row=row, column=10).value
        i_val = ws.cell(row=row, column=9).value
        k_val = ws.cell(row=row, column=11).value
        l_val = ws.cell(row=row, column=12).value
        c_val = ws.cell(row=row, column=3).value
        d_val = ws.cell(row=row, column=4).value
        if all((v is None for v in [a_val, j_val, i_val, k_val, l_val, c_val, d_val])):
            continue
        row_data = {'row': row, 'problem_num': str(a_val).strip() if a_val else '', 'problem': str(c_val).strip() if c_val else '', 'measure': str(d_val).strip() if d_val else '', 'responsible': str(i_val).strip() if i_val else '', 'plan_fact': str(j_val).strip() if j_val else '', 'start_date': k_val, 'end_date': l_val}
        result['data_rows'].append(row_data)
    dj_col = column_index_from_string('DJ')
    dk_col = column_index_from_string('DK')
    dl_col = column_index_from_string('DL')
    dm_col_idx = column_index_from_string('DM')
    result['status_columns'] = {'status_header': _cell_str_by_idx(ws, dj_col, 16), 'comments_header': _cell_str_by_idx(ws, dm_col_idx, 16), 'status_present': bool(ws.cell(row=16, column=dj_col).value), 'comments_present': bool(ws.cell(row=16, column=dm_col_idx).value)}
    formulas = {}
    dk17 = ws.cell(row=17, column=dk_col)
    formulas['DK17'] = {'value': str(dk17.value) if dk17.value else '', 'is_formula': str(dk17.value).startswith('=') if dk17.value else False}
    dk18 = ws.cell(row=18, column=dk_col)
    formulas['DK18'] = {'value': str(dk18.value) if dk18.value else '', 'is_formula': str(dk18.value).startswith('=') if dk18.value else False}
    dl18 = ws.cell(row=18, column=dl_col)
    formulas['DL18'] = {'value': str(dl18.value) if dl18.value else '', 'is_formula': str(dl18.value).startswith('=') if dl18.value else False}
    dk12 = ws.cell(row=12, column=dk_col)
    formulas['DK12'] = {'value': str(dk12.value) if dk12.value else '', 'is_formula': str(dk12.value).startswith('=') if dk12.value else False, 'has_error': '#REF' in str(dk12.value) if dk12.value else False}
    n12 = ws.cell(row=12, column=column_index_from_string('N'))
    formulas['N12'] = {'value': str(n12.value) if n12.value else '', 'is_formula': str(n12.value).startswith('=') if n12.value else False}
    result['formulas'] = formulas
    for row in range(max(1, max_row - 5), max_row + 1):
        for col in range(1, 160):
            val = ws.cell(row=row, column=col).value
            if val and str(val).strip() and ('подпись' in str(val).lower() or '___' in str(val) or '202' in str(val)):
                result['signature'] += f'{get_column_letter(col)}{row}: {str(val).strip()}\n'
    return result

def plan_grafik_parser__cell_str(ws, col_letter: str, row: int) -> str:
    """Извлекает строковое значение ячейки по букве столбца и номеру строки."""
    col_idx = column_index_from_string(col_letter)
    val = ws.cell(row=row, column=col_idx).value
    return str(val).strip() if val else ''

def _cell_str_by_idx(ws, col_idx: int, row: int) -> str:
    """Извлекает строковое значение ячейки по индексу столбца и номеру строки."""
    val = ws.cell(row=row, column=col_idx).value
    return str(val).strip() if val else ''

plan_grafik_parser_module = SimpleNamespace(parse_plan_grafik=parse_plan_grafik, _cell_str=plan_grafik_parser__cell_str, _cell_str_by_idx=_cell_str_by_idx)

# END_SOURCE_PLAN_GRAFIK_PARSER

# START_SOURCE_KPSC_SHEET_FINDER
# PURPOSE: Inlined source from audit_engine/kpsc/sheet_finder.py.
kpsc_sheet_finder__LATIN_TO_CYRILLIC = str.maketrans({'A': 'А', 'B': 'В', 'C': 'С', 'E': 'Е', 'H': 'Н', 'K': 'К', 'M': 'М', 'O': 'О', 'P': 'Р', 'T': 'Т', 'X': 'Х', 'a': 'а', 'c': 'с', 'e': 'е', 'o': 'о', 'p': 'р', 'x': 'х'})

def kpsc_sheet_finder__normalize(name: str) -> str:
    """Нормализация имени листа: strip, lower, Latin→Cyrillic, убираем спец-символы."""
    name = name.strip().lower()
    name = name.translate(kpsc_sheet_finder__LATIN_TO_CYRILLIC)
    return name

def find_sheet(wb: Workbook, keywords: List[str], *, exclude_keywords: Optional[List[str]]=None, prefer_keywords: Optional[List[str]]=None) -> Optional[Worksheet]:
    """
    Нечёткий поиск листа в workbook по ключевым словам.

    Алгоритм:
    1. Нормализуем имена листов (strip, lower, Latin→Cyrillic)
    2. Ищем листы, содержащие ВСЕ keywords
    3. Исключаем листы с exclude_keywords
    4. Из оставшихся предпочитаем с prefer_keywords
    5. Из финальных кандидатов выбираем лист с максимумом данных

    Args:
        wb: openpyxl Workbook
        keywords: обязательные подстроки (нормализованные, lowercase)
        exclude_keywords: исключающие подстроки (если есть — лист отбрасывается)
        prefer_keywords: предпочтительные подстроки (приоритет при нескольких кандидатах)

    Returns:
        Worksheet или None, если не найден
    """
    if exclude_keywords is None:
        exclude_keywords = []
    if prefer_keywords is None:
        prefer_keywords = []
    kw_norm = [kpsc_sheet_finder__normalize(kw) for kw in keywords]
    excl_norm = [kpsc_sheet_finder__normalize(ek) for ek in exclude_keywords]
    pref_norm = [kpsc_sheet_finder__normalize(pk) for pk in prefer_keywords]
    candidates = []
    for sheet_name in wb.sheetnames:
        name_norm = kpsc_sheet_finder__normalize(sheet_name)
        if not all((kw in name_norm for kw in kw_norm)):
            continue
        if any((ek in name_norm for ek in excl_norm)):
            continue
        candidates.append(sheet_name)
    if not candidates:
        return None
    if len(candidates) == 1:
        return wb[candidates[0]]
    if pref_norm:
        preferred = []
        for sn in candidates:
            name_norm = kpsc_sheet_finder__normalize(sn)
            if any((pk in name_norm for pk in pref_norm)):
                preferred.append(sn)
        if preferred:
            candidates = preferred
    if len(candidates) == 1:
        return wb[candidates[0]]
    best_sheet = None
    best_count = -1
    for sn in candidates:
        ws = wb[sn]
        count = 0
        for row in ws.iter_rows(min_row=1, max_row=min(20, ws.max_row or 1)):
            for cell in row:
                if cell.value not in (None, ''):
                    count += 1
        if count > best_count:
            best_count = count
            best_sheet = sn
    return wb[best_sheet] if best_sheet else None

def kpsc_sheet_finder_find_sheet_or_raise(wb: Workbook, keywords: List[str], parser_name: str, **kwargs) -> Worksheet:
    """
    find_sheet() с выбросом исключения если лист не найден.

    Args:
        wb: openpyxl Workbook
        keywords: ключевые слова для поиска
        parser_name: имя парсера (для сообщения об ошибке)
        **kwargs: дополнительные аргументы для find_sheet()

    Returns:
        Worksheet

    Raises:
        ValueError: если лист не найден
    """
    ws = find_sheet(wb, keywords, **kwargs)
    if ws is None:
        raise ValueError(f'[{parser_name}] Лист не найден по keywords={keywords} среди {wb.sheetnames}')
    return ws

kpsc_sheet_finder_module = SimpleNamespace(_LATIN_TO_CYRILLIC=kpsc_sheet_finder__LATIN_TO_CYRILLIC, _normalize=kpsc_sheet_finder__normalize, find_sheet=find_sheet, find_sheet_or_raise=kpsc_sheet_finder_find_sheet_or_raise)

# END_SOURCE_KPSC_SHEET_FINDER

# START_SOURCE_KPSC_PARSE_KPSC_HEADER
# PURPOSE: Inlined source from audit_engine/kpsc/parser_scripts/parse_kpsc_header.py.
kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW = 15

def kpsc_parse_kpsc_header_collect_cells(ws, max_row: int=kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW, max_col: Optional[int]=None):
    """Собираем все непустые ячейки в заголовочном регионе."""
    if max_col is None:
        max_col = ws.max_column or 50
    cells = []
    comments = []
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            if val not in (None, ''):
                cells.append({'coord': f'{get_column_letter(c)}{r}', 'row': r, 'col': c, 'value': val})
            if cell.comment:
                comments.append({'coord': f'{get_column_letter(c)}{r}', 'row': r, 'col': c, 'author': cell.comment.author, 'text': cell.comment.text})
    return (cells, comments)

def kpsc_parse_kpsc_header_collect_merged(ws, max_row: int=kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW, max_col: Optional[int]=None):
    """Собираем merged-диапазоны в заголовочном регионе."""
    if max_col is None:
        max_col = ws.max_column or 50
    merged = []
    for m in ws.merged_cells.ranges:
        if m.min_row <= max_row and m.min_col <= max_col:
            merged.append({'coord': m.coord, 'min_row': m.min_row, 'max_row': m.max_row, 'min_col': m.min_col, 'max_col': m.max_col})
    return merged

def kpsc_parse_kpsc_header__find_label_value(ws, label_keywords: List[str], max_row: int=kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW) -> Optional[Any]:
    """
    Динамический поиск значения по лейблу.

    Ищем ячейку, содержащую одно из label_keywords, затем берём значение
    из ближайшей непустой ячейки справа в той же строке.

    Если лейбл содержит значение inline (например "Наименование потока: Производство..."),
    извлекаем часть после двоеточия.
    """
    for r in range(1, max_row + 1):
        for c in range(1, min(5, (ws.max_column or 5) + 1)):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            v_lower = v.strip().lower()
            for kw in label_keywords:
                if kw not in v_lower:
                    continue
                if ':' in v:
                    parts = v.split(':', 1)
                    inline_val = parts[1].strip()
                    if inline_val:
                        return inline_val
                for vc in range(c + 1, min(c + 4, (ws.max_column or c) + 1)):
                    val = ws.cell(row=r, column=vc).value
                    if val not in (None, ''):
                        return val
                return None
    return None

def kpsc_parse_kpsc_header__find_title(ws, max_row: int=kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW) -> Optional[str]:
    """
    Извлекает заголовок карты потока.

    Ищем в первых строках длинную строку, содержащую ключевые слова
    типа "карта потока", "КПСЦ", "текущее состояние".
    """
    title_keywords = ['карта потока', 'кпсц', 'текущее состояние', 'наименование потока']
    for r in range(1, min(4, max_row + 1)):
        for c in range(1, 4):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and len(v.strip()) > 15:
                v_lower = v.strip().lower()
                if any((tk in v_lower for tk in title_keywords)):
                    return v.strip()
    for r in range(1, 3):
        for c in range(1, 4):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and len(v.strip()) > 15:
                return v.strip()
    return None

def kpsc_parse_kpsc_header__find_organization(ws, max_row: int=3) -> Optional[str]:
    """
    Ищем название организации (ООО/АО/ПАО + наименование) во всех ячейках первых строк.

    Некоторые компании (biznes_otel) размещают ООО в merged-ячейках далеко справа (col 45+),
    поэтому сканируем всю ширину строки.
    """
    import re
    org_pattern = re.compile('(ООО|АО|ПАО|ОАО|ЗАО)\\s*[«"\\\'"].+?[»"\\\'\\"]', re.IGNORECASE)
    max_col = ws.max_column or 50
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str):
                m = org_pattern.search(v)
                if m:
                    return m.group(0)
    return None

def kpsc_parse_kpsc_header_extract_fields(ws) -> Dict[str, Any]:
    """
    Динамическое извлечение полей заголовка КПСЦ.

    Вместо захардкоженных координат ищем лейблы по ключевым словам
    и берём значения из соседних ячеек.
    """
    return {'title': kpsc_parse_kpsc_header__find_title(ws), 'organization': kpsc_parse_kpsc_header__find_organization(ws), 'flow_name': kpsc_parse_kpsc_header__find_label_value(ws, ['поток:', 'наименование потока']), 'responsible': kpsc_parse_kpsc_header__find_label_value(ws, ['ответственн']), 'date_developed': kpsc_parse_kpsc_header__find_label_value(ws, ['дата разработ']), 'date_implementation': kpsc_parse_kpsc_header__find_label_value(ws, ['дата реализ', 'дата достиж']), 'compiled_by': kpsc_parse_kpsc_header__find_label_value(ws, ['составил', 'разработал']), 'takt_time': kpsc_parse_kpsc_header__find_label_value(ws, ['такт', 'время такта', 'takt'])}

def kpsc_parse_kpsc_header__find_takt_time_in_pokazateli(wb) -> Optional[Any]:
    """
    Fallback: ищет 'Время такта' на листе Показатели если не найдено в шапке КПСЦ.
    Сканирует весь лист, ищет строку-лейбл и берёт значение из соседней ячейки справа.
    """
    ws = find_sheet(wb, keywords=['показател'])
    if ws is None:
        return None
    for r in range(1, (ws.max_row or 50) + 1):
        for c in range(1, min(5, (ws.max_column or 5) + 1)):
            v = ws.cell(row=r, column=c).value
            if not isinstance(v, str):
                continue
            if 'время такта' in v.strip().lower() or 'такт' in v.strip().lower():
                for vc in range(c + 1, min(c + 4, (ws.max_column or c) + 1)):
                    val = ws.cell(row=r, column=vc).value
                    if val not in (None, ''):
                        return val
    return None

def kpsc_parse_kpsc_header_build_payload(xlsx: Path, sheet_name: Optional[str]=None):
    """Строит payload из данных header-блока КПСЦ."""
    wb = load_workbook(xlsx, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = find_sheet(wb, keywords=['кпсц'], exclude_keywords=['спагетти', 'укрупн', 'оцифровк'], prefer_keywords=['тс', 'текущ'])
        if ws is None:
            raise ValueError(f'Лист КПСЦ не найден среди {wb.sheetnames}')
    actual_sheet = ws.title
    max_col = ws.max_column or 50
    cells, _comments = kpsc_parse_kpsc_header_collect_cells(ws, max_col=max_col)
    fields = kpsc_parse_kpsc_header_extract_fields(ws)
    if not fields.get('takt_time'):
        fields['takt_time'] = kpsc_parse_kpsc_header__find_takt_time_in_pokazateli(wb)
    return {'meta': {'workbook': str(xlsx), 'sheet': actual_sheet, 'region': f'A1:{get_column_letter(max_col)}{kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW}'}, 'fields': fields}

def kpsc_parse_kpsc_header_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит верхний блок КПСЦ и сохраняет результат в output_dir/kpsc_header_v2.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_kpsc_header_build_payload(xlsx_path)
    output_path = output_dir / 'kpsc_header_v2.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_kpsc_header_main():
    ap = argparse.ArgumentParser(description='Парсер верхнего блока КПСЦ')
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_kpsc_header_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

kpsc_parse_kpsc_header_module = SimpleNamespace(HEADER_SCAN_MAX_ROW=kpsc_parse_kpsc_header_HEADER_SCAN_MAX_ROW, collect_cells=kpsc_parse_kpsc_header_collect_cells, collect_merged=kpsc_parse_kpsc_header_collect_merged, _find_label_value=kpsc_parse_kpsc_header__find_label_value, _find_title=kpsc_parse_kpsc_header__find_title, _find_organization=kpsc_parse_kpsc_header__find_organization, extract_fields=kpsc_parse_kpsc_header_extract_fields, _find_takt_time_in_pokazateli=kpsc_parse_kpsc_header__find_takt_time_in_pokazateli, build_payload=kpsc_parse_kpsc_header_build_payload, parse=kpsc_parse_kpsc_header_parse, main=kpsc_parse_kpsc_header_main)

# END_SOURCE_KPSC_PARSE_KPSC_HEADER

# START_SOURCE_KPSC_PARSE_KPSC_TABLE1
# PURPOSE: Inlined source from audit_engine/kpsc/parser_scripts/parse_kpsc_table1.py.
def kpsc_parse_kpsc_table1_find_section_anchor(ws, phrase: str) -> Optional[Tuple[int, int]]:
    """
    Ищем начало секции таблицы КПСЦ с показателями потока.

    Стратегия (по приоритету):
    1. "1. Определение показателей потока" — стандартный формат (biznes_otel)
    2. "Название этапа процесса" — плоская таблица (mapper, sodex)
    3. "Показатель" рядом с "Ед. измерения" — сводная таблица (rotosnab, ruslet)
    """
    phrase_low = phrase.lower()
    for row in ws.iter_rows():
        for cell in row:
            val = cell.value
            if isinstance(val, str) and phrase_low in val.lower():
                if val.strip().startswith('1'):
                    return (cell.row, cell.column)
    for row in ws.iter_rows(max_row=min(15, ws.max_row)):
        for cell in row:
            val = cell.value
            if isinstance(val, str) and 'название этапа' in val.lower():
                return (max(1, cell.row - 1), cell.column)
    for r in range(1, ws.max_row + 1):
        row_texts = []
        first_col = None
        for c in range(1, min(10, (ws.max_column or 10) + 1)):
            v = ws.cell(r, c).value
            if isinstance(v, str):
                row_texts.append(v.lower())
                if first_col is None:
                    first_col = c
        joined = ' '.join(row_texts)
        if 'показатель' in joined and 'ед.' in joined:
            return (max(1, r - 1), first_col or 1)
    return None

def kpsc_parse_kpsc_table1_find_header_row(ws, anchor_row: int) -> int:
    for r in range(anchor_row + 1, anchor_row + 5):
        if any((c.value not in (None, '') for c in ws[r])):
            return r
    return anchor_row + 1

def kpsc_parse_kpsc_table1_find_bottom_row(ws, header_row: int, section_col: int) -> int:
    """Ищем конец таблицы: первую строку, где в колонке section_col начинается следующая секция ("2.")."""
    r = header_row + 1
    while r <= ws.max_row:
        cell_val = ws.cell(row=r, column=section_col).value
        if isinstance(cell_val, str) and cell_val.strip().startswith('2.'):
            return r - 1
        row_vals = [c.value for c in ws[r]]
        if all((v in (None, '') for v in row_vals)) and r > header_row + 2:
            return r - 1
        r += 1
    return ws.max_row

def kpsc_parse_kpsc_table1_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    """
    Границы по строке заголовков:
      left  — первый непустой столбец,
      right — последний столбец в заголовке, где текст содержит 'ед. измерения'
              (регистр не важен); если не найдено, берем правый непустой.
    """
    cells = list(ws[header_row])
    left = None
    right = None
    right_by_text = None
    for c in cells:
        val = c.value
        if val not in (None, ''):
            if left is None:
                left = c.column
            if isinstance(val, str) and 'ед. измерения' in val.lower():
                right_by_text = c.column
            right = c.column
    if right_by_text:
        right = right_by_text
    if left is None:
        left = 1
    if right is None:
        right = left
    return (left, right)

def kpsc_parse_kpsc_table1_build_merged_lookup(ws):
    lookup = {}
    for merge in ws.merged_cells.ranges:
        coord = merge.coord
        min_row, min_col, max_row, max_col = (merge.min_row, merge.min_col, merge.max_row, merge.max_col)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                lookup[r, c] = coord
    return lookup

def kpsc_parse_kpsc_table1_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_kpsc_table1_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            cell = ws.cell(row=r, column=c)
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = cell.value
            row_cells.append({'coord': coord, 'col': c, 'row': r, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_kpsc_table1_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        ws = find_sheet(wb, keywords=['кпсц'], exclude_keywords=['спагетти', 'укрупн', 'оцифровк'], prefer_keywords=['тс', 'текущ'])
        if ws is None:
            raise ValueError(f'Лист КПСЦ не найден среди {wb.sheetnames}')
    actual_sheet = ws.title
    anchor = kpsc_parse_kpsc_table1_find_section_anchor(ws, 'определение показателей потока')
    if not anchor:
        return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'section_title_cell': None}, 'bounds': None, 'rows': []}
    anchor_row, anchor_col = anchor
    header_row = kpsc_parse_kpsc_table1_find_header_row(ws, anchor_row)
    bottom_row = kpsc_parse_kpsc_table1_find_bottom_row(ws, header_row, anchor_col)
    left_col, right_col = kpsc_parse_kpsc_table1_compute_col_bounds(ws, header_row)
    table_rows = kpsc_parse_kpsc_table1_extract_table(ws, header_row, bottom_row, left_col, right_col)
    payload = {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'section_title_cell': f'{get_column_letter(anchor_col)}{anchor_row}'}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col, 'left_letter': get_column_letter(left_col), 'right_letter': get_column_letter(right_col), 'height': bottom_row - header_row + 1, 'width': right_col - left_col + 1}, 'rows': table_rows}
    return payload

def kpsc_parse_kpsc_table1_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит таблицу '1. Определение показателей потока' и сохраняет в kpsc_table1_v2.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_kpsc_table1_build_payload(xlsx_path)
    output_path = output_dir / 'kpsc_table1_v2.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_kpsc_table1_main():
    parser = argparse.ArgumentParser(description="Парсер таблицы '1. Определение показателей потока'")
    parser.add_argument('-i', '--input', required=True, help='XLSX файл')
    parser.add_argument('-s', '--sheet', default=None, help='Лист (default: auto)')
    parser.add_argument('-o', '--output', default=None, help='JSON файл вывода (stdout если не указан)')
    args = parser.parse_args()
    payload = kpsc_parse_kpsc_table1_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

kpsc_parse_kpsc_table1_module = SimpleNamespace(find_section_anchor=kpsc_parse_kpsc_table1_find_section_anchor, find_header_row=kpsc_parse_kpsc_table1_find_header_row, find_bottom_row=kpsc_parse_kpsc_table1_find_bottom_row, compute_col_bounds=kpsc_parse_kpsc_table1_compute_col_bounds, build_merged_lookup=kpsc_parse_kpsc_table1_build_merged_lookup, extract_table=kpsc_parse_kpsc_table1_extract_table, build_payload=kpsc_parse_kpsc_table1_build_payload, parse=kpsc_parse_kpsc_table1_parse, main=kpsc_parse_kpsc_table1_main)

# END_SOURCE_KPSC_PARSE_KPSC_TABLE1

# START_SOURCE_KPSC_PARSE_LEGEND
# PURPOSE: Inlined source from audit_engine/kpsc/parser_scripts/parse_legend.py.
kpsc_parse_legend_NS = {'wb': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}

def kpsc_parse_legend_read_sheet_drawing(xlsx: Path, sheet_name: str):
    with zipfile.ZipFile(xlsx) as z:
        wb_xml = ET.fromstring(z.read('xl/workbook.xml'))
        wb_rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rel_map = {rel.attrib['Id']: rel.attrib['Target'] for rel in wb_rels}
        sheet_target = None
        for sheet in wb_xml.find('wb:sheets', kpsc_parse_legend_NS):
            if sheet.attrib['name'] == sheet_name:
                rid = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
                sheet_target = rel_map[rid]
                break
        rel_path = f'xl/worksheets/_rels/{posixpath.basename(sheet_target)}.rels'
        rels = ET.fromstring(z.read(rel_path))
        drawing_path = None
        for rel in rels:
            if rel.attrib.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing':
                drawing_path = posixpath.normpath(posixpath.join('xl/worksheets', rel.attrib['Target']))
        d_rels_path = posixpath.join(posixpath.dirname(drawing_path), '_rels', posixpath.basename(drawing_path) + '.rels')
        d_rels = ET.fromstring(z.read(d_rels_path)) if d_rels_path in z.namelist() else None
        rel_pic = {rel.attrib['Id']: rel.attrib['Target'] for rel in d_rels} if d_rels is not None else {}
        drawing_root = ET.fromstring(z.read(drawing_path))
    return (drawing_root, rel_pic)

def kpsc_parse_legend_parse_pictures(drawing_root: ET.Element, rel_pic: Dict[str, str]):
    pics = []
    for anc in drawing_root.findall('./', kpsc_parse_legend_NS):
        pic_el = anc.find('xdr:pic', kpsc_parse_legend_NS)
        if pic_el is None:
            continue
        pf = anc.find('xdr:from', kpsc_parse_legend_NS)
        pt = anc.find('xdr:to', kpsc_parse_legend_NS)
        fcol = int(pf.find('xdr:col', kpsc_parse_legend_NS).text)
        frow = int(pf.find('xdr:row', kpsc_parse_legend_NS).text)
        tcol = int(pt.find('xdr:col', kpsc_parse_legend_NS).text)
        trow = int(pt.find('xdr:row', kpsc_parse_legend_NS).text)
        blip = pic_el.find('.//a:blip', kpsc_parse_legend_NS)
        rid = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed') if blip is not None else None
        pics.append({'row': frow + 1, 'col': fcol + 1, 'bbox': {'from': {'row': frow + 1, 'col': fcol + 1}, 'to': {'row': trow + 1, 'col': tcol + 1}}, 'rel_id': rid, 'target': rel_pic.get(rid)})
    return pics

def kpsc_parse_legend_collect_text(ws):
    texts = []
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=2).value
        if val not in (None, ''):
            texts.append({'row': row, 'col': 2, 'text': val})
    return texts

def kpsc_parse_legend_match_pics(texts: List[Dict[str, Any]], pics: List[Dict[str, Any]]):
    matched = []
    texts_order = [t for t in sorted(texts, key=lambda x: x['row']) if t['text'] != 'Расшифровка или пояснение']
    pics_order = sorted(pics, key=lambda x: x['row'])
    n = min(len(texts_order), len(pics_order))
    for i, t in enumerate(texts_order):
        pic = pics_order[i] if i < len(pics_order) else None
        matched.append({'row': t['row'], 'text': t['text'], 'picture': pic})
    return matched

def kpsc_parse_legend_build_payload(xlsx: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['условн', 'обозн'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx), 'sheet': None}, 'pictures': [], 'entries': []}
        actual_sheet = ws.title
    drawing_root, rel_pic = kpsc_parse_legend_read_sheet_drawing(xlsx, actual_sheet)
    pics = kpsc_parse_legend_parse_pictures(drawing_root, rel_pic)
    texts = kpsc_parse_legend_collect_text(ws)
    entries = kpsc_parse_legend_match_pics(texts, pics)
    return {'meta': {'workbook': str(xlsx), 'sheet': actual_sheet}, 'pictures': pics, 'entries': entries}

def kpsc_parse_legend_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит лист 'Условные обозначения' и сохраняет в legend_v2.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_legend_build_payload(xlsx_path)
    output_path = output_dir / 'legend_v2.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    return payload

def kpsc_parse_legend_main():
    ap = argparse.ArgumentParser(description="Парсер листа 'Условные обозначения'")
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_legend_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

kpsc_parse_legend_module = SimpleNamespace(NS=kpsc_parse_legend_NS, read_sheet_drawing=kpsc_parse_legend_read_sheet_drawing, parse_pictures=kpsc_parse_legend_parse_pictures, collect_text=kpsc_parse_legend_collect_text, match_pics=kpsc_parse_legend_match_pics, build_payload=kpsc_parse_legend_build_payload, parse=kpsc_parse_legend_parse, main=kpsc_parse_legend_main)

# END_SOURCE_KPSC_PARSE_LEGEND

# START_SOURCE_KPSC_PARSE_LOSS_DIGITIZATION
# PURPOSE: Inlined source from audit_engine/kpsc/parser_scripts/parse_loss_digitization.py.
kpsc_parse_loss_digitization_HEADER_KEYS = ('описание проблемы', 'вид потери')

def kpsc_parse_loss_digitization_find_header_row(ws) -> Optional[int]:
    """Находим строку заголовков по ключевым фразам."""
    for r in range(1, ws.max_row + 1):
        lower_vals = [str(c.value).lower() for c in ws[r] if isinstance(c.value, str)]
        if lower_vals and all((any((key in v for v in lower_vals)) for key in kpsc_parse_loss_digitization_HEADER_KEYS)):
            return r
    return None

def kpsc_parse_loss_digitization_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    """Границы по непустым ячейкам строки заголовков (от первой до последней)."""
    left = None
    right = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v not in (None, ''):
            if left is None:
                left = c
            right = c
    if left is None:
        left = 1
        right = 1
    return (left, right)

def kpsc_parse_loss_digitization_find_bottom_row(ws, header_row: int, left_col: int, right_col: int) -> int:
    """
    Ищем конец таблицы: первая строка после заголовка, где данные отсутствуют
    во всех колонках кроме порядкового номера (left_col). Строки с одним
    номером считаем пустыми.
    """
    for r in range(header_row + 1, ws.max_row + 1):
        has_data = any((ws.cell(row=r, column=c).value not in (None, '') for c in range(left_col + 1, right_col + 1)))
        if not has_data:
            return r - 1
    return ws.max_row

def kpsc_parse_loss_digitization_build_merged_lookup(ws):
    lookup = {}
    for merge in ws.merged_cells.ranges:
        coord = merge.coord
        for r in range(merge.min_row, merge.max_row + 1):
            for c in range(merge.min_col, merge.max_col + 1):
                lookup[r, c] = coord
    return lookup

def kpsc_parse_loss_digitization_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_loss_digitization_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = ws.cell(row=r, column=c).value
            row_cells.append({'coord': coord, 'row': r, 'col': c, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_loss_digitization_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['оцифровк'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx_path), 'sheet': None, 'header_row': None}, 'bounds': None, 'rows': []}
        actual_sheet = ws.title
    header_row = kpsc_parse_loss_digitization_find_header_row(ws)
    if not header_row:
        return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': None}, 'bounds': None, 'rows': []}
    left_col, right_col = kpsc_parse_loss_digitization_compute_col_bounds(ws, header_row)
    bottom_row = kpsc_parse_loss_digitization_find_bottom_row(ws, header_row, left_col, right_col)
    rows = kpsc_parse_loss_digitization_extract_table(ws, header_row, bottom_row, left_col, right_col)
    return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': header_row}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col, 'left_letter': get_column_letter(left_col), 'right_letter': get_column_letter(right_col), 'height': bottom_row - header_row + 1, 'width': right_col - left_col + 1}, 'rows': rows}

def kpsc_parse_loss_digitization_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит лист 'Оцифровка потерь КПСЦ' и сохраняет в ocifrovka_poteri_v2.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_loss_digitization_build_payload(xlsx_path)
    output_path = output_dir / 'ocifrovka_poteri_v2.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_loss_digitization_main():
    ap = argparse.ArgumentParser(description="Парсер листа 'Оцифровка потерь КПСЦ'")
    ap.add_argument('-i', '--input', required=True, help='Путь к XLSX файлу')
    ap.add_argument('-s', '--sheet', default=None, help='Имя листа (default: auto)')
    ap.add_argument('-o', '--output', help='JSON файл вывода (stdout если не указан)')
    args = ap.parse_args()
    payload = kpsc_parse_loss_digitization_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

kpsc_parse_loss_digitization_module = SimpleNamespace(HEADER_KEYS=kpsc_parse_loss_digitization_HEADER_KEYS, find_header_row=kpsc_parse_loss_digitization_find_header_row, compute_col_bounds=kpsc_parse_loss_digitization_compute_col_bounds, find_bottom_row=kpsc_parse_loss_digitization_find_bottom_row, build_merged_lookup=kpsc_parse_loss_digitization_build_merged_lookup, extract_table=kpsc_parse_loss_digitization_extract_table, build_payload=kpsc_parse_loss_digitization_build_payload, parse=kpsc_parse_loss_digitization_parse, main=kpsc_parse_loss_digitization_main)

# END_SOURCE_KPSC_PARSE_LOSS_DIGITIZATION

# START_SOURCE_KPSC_PARSE_PA1_CHART
# PURPOSE: Inlined source from audit_engine/kpsc/parser_scripts/parse_pa1_chart.py.
def kpsc_parse_pa1_chart__sheet_name_from_range(rng: str) -> Tuple[str, str]:
    """Возвращает (sheet_name, range_part) из строки вида 'Лист'!$A$1:$B$2.

    Обрабатывает ссылки на внешние книги: '[2]ПА1 '!$A$1 → 'ПА1 '.
    """
    import re
    if '!' in rng:
        sheet, r = rng.split('!', 1)
        sheet = sheet.strip("'")
        sheet = re.sub('^\\[\\d+\\]', '', sheet)
        return (sheet, r)
    return ('', rng)

def kpsc_parse_pa1_chart__values_from_range(wb, sheet_name: str, rng: str):
    ws = wb[sheet_name]
    min_col, min_row, max_col, max_row = range_boundaries(rng)
    values = []
    for r in range(min_row, max_row + 1):
        row = []
        for c in range(min_col, max_col + 1):
            row.append(ws.cell(r, c).value)
        values.append(row)
    if min_row == max_row or min_col == max_col:
        return [v[0] if min_col == max_col else v for v in values] if min_row != max_row else values[0]
    return values

def kpsc_parse_pa1_chart__chart_title(chart: ChartBase) -> Optional[str]:
    t = chart.title
    if t is None:
        return None
    try:
        if t.tx and t.tx.rich and t.tx.rich.p:
            return ''.join((r.t for r in t.tx.rich.p[0].r))
    except Exception:
        pass
    return None

def kpsc_parse_pa1_chart_extract_chart_payload(chart: ChartBase, wb_data, wb_formulas) -> Dict[str, Any]:
    payload: Dict[str, Any] = {'type': type(chart).__name__, 'title': kpsc_parse_pa1_chart__chart_title(chart)}
    if getattr(chart, 'anchor', None) and getattr(chart.anchor, '_from', None):
        a_from = chart.anchor._from
        a_to = chart.anchor.to
        payload['anchor'] = {'from': {'col': a_from.col + 1, 'row': a_from.row + 1}, 'to': {'col': a_to.col + 1, 'row': a_to.row + 1}}
    series_list = []
    for s in chart.series:
        name = None
        if s.title:
            if getattr(s.title, 'v', None):
                name = s.title.v
            elif getattr(s.title, 'strRef', None) and s.title.strRef.strCache and s.title.strRef.strCache.pt:
                name = s.title.strRef.strCache.pt[0].v
        val_ref = getattr(getattr(s, 'val', None), 'numRef', None)
        val_range = val_ref.f if val_ref else None
        values = None
        if val_range:
            sheet_name, rng = kpsc_parse_pa1_chart__sheet_name_from_range(val_range)
            values = kpsc_parse_pa1_chart__values_from_range(wb_data, sheet_name, rng)
        cat_ref_obj = getattr(getattr(s, 'cat', None), 'strRef', None)
        cat_range = cat_ref_obj.f if cat_ref_obj else None
        categories = None
        if cat_range:
            sheet_name, rng = kpsc_parse_pa1_chart__sheet_name_from_range(cat_range)
            categories = kpsc_parse_pa1_chart__values_from_range(wb_data, sheet_name, rng)
        series_list.append({'name': name, 'values_range': val_range, 'values': values, 'categories_range': cat_range, 'categories': categories})
    payload['series'] = series_list
    return payload

def kpsc_parse_pa1_chart__find_pa1_sheet(wb) -> Optional[str]:
    """Находит лист ПА-1 через sheet_finder."""
    ws = find_sheet(wb, keywords=['па'], exclude_keywords=['спагетти', 'кпсц', 'ямадз'])
    return ws.title if ws else None

def kpsc_parse_pa1_chart_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb_formulas = load_workbook(xlsx_path, data_only=False)
    wb_data = load_workbook(xlsx_path, data_only=True)
    if not sheet_name:
        ws_found = find_sheet(wb_formulas, keywords=['па'], exclude_keywords=['спагетти', 'кпсц', 'ямадз'])
        if ws_found is None:
            return {'meta': {'workbook': str(xlsx_path), 'sheet': None}, 'charts': [], 'text_boxes': []}
        sheet_name = ws_found.title
    ws = wb_formulas[sheet_name]
    charts = getattr(ws, '_charts', [])
    charts_payload = [kpsc_parse_pa1_chart_extract_chart_payload(ch, wb_data, wb_formulas) for ch in charts]
    text_boxes: List[Dict[str, Any]] = []
    try:
        with zipfile.ZipFile(xlsx_path) as zf:
            import xml.etree.ElementTree as ET
            ns_wb = {'n': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
            wb_root = ET.fromstring(zf.read('xl/workbook.xml'))
            rid = None
            for sh in wb_root.findall('n:sheets/n:sheet', ns_wb):
                if sh.attrib.get('name') == sheet_name:
                    rid = sh.attrib[f'{{{ns_wb['r']}}}id']
                    break
            if rid:
                wb_rels = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
                sheet_target = None
                for rel in wb_rels:
                    if rel.attrib.get('Id') == rid:
                        sheet_target = rel.attrib['Target']
                        break
                if sheet_target:
                    sheet_rels_path = 'xl/' + sheet_target.replace('worksheets/', 'worksheets/_rels/') + '.rels'
                    if sheet_rels_path in zf.namelist():
                        sheet_rels = ET.fromstring(zf.read(sheet_rels_path))
                        drawing_target = None
                        for rel in sheet_rels:
                            if rel.attrib.get('Type') == 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing':
                                drawing_target = rel.attrib['Target']
                                break
                        if drawing_target:
                            drawing_target = drawing_target.lstrip('../')
                            drawing_path = 'xl/' + drawing_target
                            if drawing_path in zf.namelist():
                                ns = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', 'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
                                drw_root = ET.fromstring(zf.read(drawing_path))
                                for two in drw_root.findall('xdr:twoCellAnchor', ns):
                                    text_parts = [t.text or '' for t in two.findall('.//a:t', ns)]
                                    if text_parts:
                                        frm = two.find('xdr:from', ns)
                                        to = two.find('xdr:to', ns)
                                        text_boxes.append({'text': ''.join(text_parts).strip(), 'anchor': {'from': {'col': int(frm.find('xdr:col', ns).text) + 1, 'row': int(frm.find('xdr:row', ns).text) + 1}, 'to': {'col': int(to.find('xdr:col', ns).text) + 1, 'row': int(to.find('xdr:row', ns).text) + 1}}})
    except Exception:
        pass
    return {'meta': {'workbook': str(xlsx_path), 'sheet': sheet_name}, 'charts': charts_payload, 'text_boxes': text_boxes}

def kpsc_parse_pa1_chart_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит диаграммы на листе 'ПА-1' и сохраняет в pa1_chart_v3.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_pa1_chart_build_payload(xlsx_path)
    output_path = output_dir / 'pa1_chart_v3.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_pa1_chart_main():
    ap = argparse.ArgumentParser(description="Парсер диаграмм на листе 'ПА-1' (после таблицы)")
    ap.add_argument('-i', '--input', required=True, help='Путь к XLSX')
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output', help='JSON вывод (stdout если не указан)')
    args = ap.parse_args()
    payload = kpsc_parse_pa1_chart_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

kpsc_parse_pa1_chart_module = SimpleNamespace(_sheet_name_from_range=kpsc_parse_pa1_chart__sheet_name_from_range, _values_from_range=kpsc_parse_pa1_chart__values_from_range, _chart_title=kpsc_parse_pa1_chart__chart_title, extract_chart_payload=kpsc_parse_pa1_chart_extract_chart_payload, _find_pa1_sheet=kpsc_parse_pa1_chart__find_pa1_sheet, build_payload=kpsc_parse_pa1_chart_build_payload, parse=kpsc_parse_pa1_chart_parse, main=kpsc_parse_pa1_chart_main)

# END_SOURCE_KPSC_PARSE_PA1_CHART

# START_SOURCE_KPSC_PARSE_PA1_TABLE
# PURPOSE: Inlined source from audit_engine/kpsc/parser_scripts/parse_pa1_table.py.
def kpsc_parse_pa1_table_find_header_row(ws) -> Optional[int]:
    """Ищем строку, где в заголовках встречается слово 'итого'."""
    for r in range(1, ws.max_row + 1):
        if any((isinstance(c.value, str) and 'итого' in c.value.lower() for c in ws[r])):
            return r
    return None

def kpsc_parse_pa1_table_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    """
    Берём минимальный/максимальный столбцы с данными в строках заголовка
    и двух строках ниже (чтобы захватить пустой заголовок первого столбца,
    но заполненные значения 'Замер 1' и т.п.).
    """
    left = None
    right = 0
    last_row = min(ws.max_row, header_row + 2)
    for r in range(header_row, last_row + 1):
        for c in range(1, ws.max_column + 1):
            if ws.cell(row=r, column=c).value not in (None, ''):
                left = c if left is None else min(left, c)
                right = max(right, c)
    if left is None:
        left = 1
        right = 1
    return (left, right)

def kpsc_parse_pa1_table_find_bottom_row(ws, header_row: int, left_col: int, right_col: int) -> int:
    """Первая строка, полностью пустая в пределах таблицы, завершает данные."""
    r = header_row + 1
    while r <= ws.max_row:
        if all((ws.cell(row=r, column=c).value in (None, '') for c in range(left_col, right_col + 1))):
            return r - 1
        r += 1
    return ws.max_row

def kpsc_parse_pa1_table_build_merged_lookup(ws):
    lookup = {}
    for m in ws.merged_cells.ranges:
        coord = m.coord
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                lookup[r, c] = coord
    return lookup

def kpsc_parse_pa1_table_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_pa1_table_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = ws.cell(row=r, column=c).value
            row_cells.append({'coord': coord, 'row': r, 'col': c, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_pa1_table_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['па'], exclude_keywords=['спагетти', 'кпсц', 'ямадз'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx_path), 'sheet': None, 'header_row': None}, 'bounds': None, 'rows': []}
        actual_sheet = ws.title
    header_row = kpsc_parse_pa1_table_find_header_row(ws)
    if not header_row:
        return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': None}, 'bounds': None, 'rows': []}
    left_col, right_col = kpsc_parse_pa1_table_compute_col_bounds(ws, header_row)
    bottom_row = kpsc_parse_pa1_table_find_bottom_row(ws, header_row, left_col, right_col)
    rows = kpsc_parse_pa1_table_extract_table(ws, header_row, bottom_row, left_col, right_col)
    return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': header_row}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col, 'left_letter': get_column_letter(left_col), 'right_letter': get_column_letter(right_col), 'height': bottom_row - header_row + 1, 'width': right_col - left_col + 1}, 'rows': rows}

def kpsc_parse_pa1_table_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит стартовую таблицу на листе 'ПА-1' и сохраняет в pa1_table_v1.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_pa1_table_build_payload(xlsx_path)
    output_path = output_dir / 'pa1_table_v1.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_pa1_table_main():
    ap = argparse.ArgumentParser(description="Парсер стартовой таблицы на листе 'ПА-1'")
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_pa1_table_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

kpsc_parse_pa1_table_module = SimpleNamespace(find_header_row=kpsc_parse_pa1_table_find_header_row, compute_col_bounds=kpsc_parse_pa1_table_compute_col_bounds, find_bottom_row=kpsc_parse_pa1_table_find_bottom_row, build_merged_lookup=kpsc_parse_pa1_table_build_merged_lookup, extract_table=kpsc_parse_pa1_table_extract_table, build_payload=kpsc_parse_pa1_table_build_payload, parse=kpsc_parse_pa1_table_parse, main=kpsc_parse_pa1_table_main)

# END_SOURCE_KPSC_PARSE_PA1_TABLE

# START_SOURCE_KPSC_PARSE_POKAZATELI
# PURPOSE: Inlined source from audit_engine/kpsc/parser_scripts/parse_pokazateli.py.
kpsc_parse_pokazateli_TITLE_PHRASE = 'текущие показатели потока'

def kpsc_parse_pokazateli_find_title(ws) -> int:
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and kpsc_parse_pokazateli_TITLE_PHRASE in v.lower():
                return r
    raise ValueError("Title 'Текущие показатели потока' not found")

def kpsc_parse_pokazateli_find_header_row(ws, title_row: int) -> int:
    for r in range(title_row + 1, title_row + 5):
        if any((ws.cell(row=r, column=c).value not in (None, '') for c in range(1, ws.max_column + 1))):
            return r
    return title_row + 1

def kpsc_parse_pokazateli_find_bottom_row(ws, header_row: int) -> int:
    r = header_row + 1
    while r <= ws.max_row:
        if all((ws.cell(row=r, column=c).value in (None, '') for c in range(1, ws.max_column + 1))):
            return r - 1
        r += 1
    return ws.max_row

def kpsc_parse_pokazateli_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    """
    Граница по строке заголовков: берём первую непустую ячейку и продолжаем вправо,
    пока идут непустые. Если встречаем пустую колонку после начала таблицы — там обрываем.
    Это отсечёт служебные столбцы справа (I,J...).
    """
    left = None
    right = None
    started = False
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v not in (None, ''):
            if left is None:
                left = c
            right = c
            started = True
        elif started:
            break
    if left is None:
        left = 1
        right = 1
    return (left, right)

def kpsc_parse_pokazateli_build_merged_lookup(ws):
    lookup = {}
    for m in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = (m.min_col, m.min_row, m.max_col, m.max_row)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                lookup[r, c] = m.coord
    return lookup

def kpsc_parse_pokazateli_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_pokazateli_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = ws.cell(row=r, column=c).value
            row_cells.append({'coord': coord, 'row': r, 'col': c, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_pokazateli_build_payload(xlsx: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['показател'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx), 'sheet': None, 'title_row': None}, 'bounds': None, 'rows': []}
        actual_sheet = ws.title
    title_row = None
    try:
        title_row = kpsc_parse_pokazateli_find_title(ws)
    except ValueError:
        return {'meta': {'workbook': str(xlsx), 'sheet': actual_sheet, 'title_row': None}, 'bounds': None, 'rows': []}
    header_row = kpsc_parse_pokazateli_find_header_row(ws, title_row)
    bottom_row = kpsc_parse_pokazateli_find_bottom_row(ws, header_row)
    left_col, right_col = kpsc_parse_pokazateli_compute_col_bounds(ws, header_row)
    table_rows = kpsc_parse_pokazateli_extract_table(ws, header_row, bottom_row, left_col, right_col)
    return {'meta': {'workbook': str(xlsx), 'sheet': actual_sheet, 'title_row': title_row}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col}, 'rows': table_rows}

def kpsc_parse_pokazateli_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит 'Текущие показатели потока' и сохраняет в pokazateli_v3.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_pokazateli_build_payload(xlsx_path)
    output_path = output_dir / 'pokazateli_v3.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_pokazateli_main():
    ap = argparse.ArgumentParser(description="Парсер 'Текущие показатели потока'")
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_pokazateli_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

kpsc_parse_pokazateli_module = SimpleNamespace(TITLE_PHRASE=kpsc_parse_pokazateli_TITLE_PHRASE, find_title=kpsc_parse_pokazateli_find_title, find_header_row=kpsc_parse_pokazateli_find_header_row, find_bottom_row=kpsc_parse_pokazateli_find_bottom_row, compute_col_bounds=kpsc_parse_pokazateli_compute_col_bounds, build_merged_lookup=kpsc_parse_pokazateli_build_merged_lookup, extract_table=kpsc_parse_pokazateli_extract_table, build_payload=kpsc_parse_pokazateli_build_payload, parse=kpsc_parse_pokazateli_parse, main=kpsc_parse_pokazateli_main)

# END_SOURCE_KPSC_PARSE_POKAZATELI

# START_SOURCE_KPSC_PARSE_SPAGHETTI_PROBLEMS
# PURPOSE: Inlined source from audit_engine/kpsc/parser_scripts/parse_spaghetti_problems.py.
kpsc_parse_spaghetti_problems_HEADER_KEY = 'описание проблемы'

def kpsc_parse_spaghetti_problems_find_header_row(ws) -> Optional[int]:
    for r in range(1, ws.max_row + 1):
        if any((isinstance(c.value, str) and kpsc_parse_spaghetti_problems_HEADER_KEY in c.value.lower() for c in ws[r])):
            return r
    return None

def kpsc_parse_spaghetti_problems_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    left = None
    right = 0
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v not in (None, ''):
            if left is None:
                left = c
            right = c
    if left is None:
        left = 1
        right = 1
    return (left, right)

def kpsc_parse_spaghetti_problems_find_bottom_row(ws, header_row: int, left_col: int, right_col: int) -> int:
    r = header_row + 1
    while r <= ws.max_row:
        if all((ws.cell(row=r, column=c).value in (None, '') for c in range(left_col, right_col + 1))):
            return r - 1
        r += 1
    return ws.max_row

def kpsc_parse_spaghetti_problems_build_merged_lookup(ws):
    lookup = {}
    for m in ws.merged_cells.ranges:
        coord = m.coord
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                lookup[r, c] = coord
    return lookup

def kpsc_parse_spaghetti_problems_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_spaghetti_problems_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = ws.cell(row=r, column=c).value
            row_cells.append({'coord': coord, 'row': r, 'col': c, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_spaghetti_problems_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['спагетти'], exclude_keywords=['диаграмм'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx_path), 'sheet': None, 'header_row': None}, 'bounds': None, 'rows': []}
        actual_sheet = ws.title
    header_row = kpsc_parse_spaghetti_problems_find_header_row(ws)
    if not header_row:
        return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': None}, 'bounds': None, 'rows': []}
    left_col, right_col = kpsc_parse_spaghetti_problems_compute_col_bounds(ws, header_row)
    bottom_row = kpsc_parse_spaghetti_problems_find_bottom_row(ws, header_row, left_col, right_col)
    rows = kpsc_parse_spaghetti_problems_extract_table(ws, header_row, bottom_row, left_col, right_col)
    return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': header_row}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col, 'left_letter': get_column_letter(left_col), 'right_letter': get_column_letter(right_col), 'height': bottom_row - header_row + 1, 'width': right_col - left_col + 1}, 'rows': rows}

def kpsc_parse_spaghetti_problems_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит лист 'Перечень проблем по спагетти' и сохраняет в spaghetti_problems_v1.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_spaghetti_problems_build_payload(xlsx_path)
    output_path = output_dir / 'spaghetti_problems_v1.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_spaghetti_problems_main():
    ap = argparse.ArgumentParser(description="Парсер листа 'Перечень проблем по спагетти'")
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_spaghetti_problems_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

kpsc_parse_spaghetti_problems_module = SimpleNamespace(HEADER_KEY=kpsc_parse_spaghetti_problems_HEADER_KEY, find_header_row=kpsc_parse_spaghetti_problems_find_header_row, compute_col_bounds=kpsc_parse_spaghetti_problems_compute_col_bounds, find_bottom_row=kpsc_parse_spaghetti_problems_find_bottom_row, build_merged_lookup=kpsc_parse_spaghetti_problems_build_merged_lookup, extract_table=kpsc_parse_spaghetti_problems_extract_table, build_payload=kpsc_parse_spaghetti_problems_build_payload, parse=kpsc_parse_spaghetti_problems_parse, main=kpsc_parse_spaghetti_problems_main)

# END_SOURCE_KPSC_PARSE_SPAGHETTI_PROBLEMS

# START_SOURCE_KPSC_PARSE_SPAGHETTI_SHEET
# PURPOSE: Inlined source from audit_engine/kpsc/parser_scripts/parse_spaghetti_sheet.py.
kpsc_parse_spaghetti_sheet_HEADER_PHRASES = ['шаги процесса', 'путь', 'перемещени']

def kpsc_parse_spaghetti_sheet_find_header_row(ws) -> Optional[int]:
    """Ищем строку-заголовок таблицы перемещений по ключевым словам."""
    for r in range(1, min(20, ws.max_row + 1)):
        row_vals = [c.value for c in ws[r]]
        for v in row_vals:
            if isinstance(v, str):
                v_low = v.lower()
                if any((phrase in v_low for phrase in kpsc_parse_spaghetti_sheet_HEADER_PHRASES)):
                    return r
    return None

def kpsc_parse_spaghetti_sheet_compute_col_bounds(ws, header_row: int) -> Tuple[int, int]:
    left = None
    right = 0
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v not in (None, ''):
            left = c
            break
    if left is None:
        left = 1
    for r in range(header_row, ws.max_row + 1):
        for c in range(left, ws.max_column + 1):
            if ws.cell(row=r, column=c).value not in (None, ''):
                right = max(right, c)
    if right < left:
        right = left
    return (left, right)

def kpsc_parse_spaghetti_sheet_find_bottom_row(ws, header_row: int, left_col: int, right_col: int) -> int:
    r = header_row + 1
    while r <= ws.max_row:
        if all((ws.cell(row=r, column=c).value in (None, '') for c in range(left_col, right_col + 1))):
            return r - 1
        r += 1
    return ws.max_row

def kpsc_parse_spaghetti_sheet_build_merged_lookup(ws):
    lookup = {}
    for m in ws.merged_cells.ranges:
        coord = m.coord
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                lookup[r, c] = coord
    return lookup

def kpsc_parse_spaghetti_sheet_extract_table(ws, top_row: int, bottom_row: int, left_col: int, right_col: int):
    merged_lookup = kpsc_parse_spaghetti_sheet_build_merged_lookup(ws)
    rows = []
    for r in range(top_row, bottom_row + 1):
        row_cells = []
        for c in range(left_col, right_col + 1):
            coord = f'{get_column_letter(c)}{r}'
            merge_range = merged_lookup.get((r, c))
            if merge_range:
                min_col_m, min_row_m, max_col_m, max_row_m = range_boundaries(merge_range)
                anchor = r == min_row_m and c == min_col_m
                value = ws.cell(row=min_row_m, column=min_col_m).value
            else:
                anchor = True
                value = ws.cell(row=r, column=c).value
            row_cells.append({'coord': coord, 'row': r, 'col': c, 'value': value, 'merge_range': merge_range, 'merge_anchor': anchor if merge_range else False})
        rows.append({'row': r, 'cells': row_cells})
    return rows

def kpsc_parse_spaghetti_sheet_extract_pre_table(ws, header_row: int):
    cells = []
    for r in range(1, header_row):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v not in (None, ''):
                cells.append({'coord': f'{get_column_letter(c)}{r}', 'row': r, 'col': c, 'value': v})
    return cells

def kpsc_parse_spaghetti_sheet_build_payload(xlsx_path: Path, sheet_name: Optional[str]=None):
    wb = load_workbook(xlsx_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
        actual_sheet = sheet_name
    else:
        ws = find_sheet(wb, keywords=['спагетти'], exclude_keywords=['пробл', 'улучш', 'перечень'])
        if ws is None:
            return {'meta': {'workbook': str(xlsx_path), 'sheet': None, 'header_row': None}, 'bounds': None, 'pre_table_cells': [], 'rows': []}
        actual_sheet = ws.title
    header_row = kpsc_parse_spaghetti_sheet_find_header_row(ws)
    if not header_row:
        return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': None}, 'bounds': None, 'pre_table_cells': [], 'rows': []}
    left_col, right_col = kpsc_parse_spaghetti_sheet_compute_col_bounds(ws, header_row)
    bottom_row = kpsc_parse_spaghetti_sheet_find_bottom_row(ws, header_row, left_col, right_col)
    pre_table = kpsc_parse_spaghetti_sheet_extract_pre_table(ws, header_row)
    table_rows = kpsc_parse_spaghetti_sheet_extract_table(ws, header_row, bottom_row, left_col, right_col)
    return {'meta': {'workbook': str(xlsx_path), 'sheet': actual_sheet, 'header_row': header_row}, 'bounds': {'top_row': header_row, 'bottom_row': bottom_row, 'left_col': left_col, 'right_col': right_col, 'left_letter': get_column_letter(left_col), 'right_letter': get_column_letter(right_col), 'height': bottom_row - header_row + 1, 'width': right_col - left_col + 1}, 'pre_table_cells': pre_table, 'rows': table_rows}

def kpsc_parse_spaghetti_sheet_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """Парсит лист 'Диаграмма Спагетти' и сохраняет в spaghetti_sheet_v2.json."""
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = kpsc_parse_spaghetti_sheet_build_payload(xlsx_path)
    output_path = output_dir / 'spaghetti_sheet_v2.json'
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    return payload

def kpsc_parse_spaghetti_sheet_main():
    ap = argparse.ArgumentParser(description="Парсер листа 'Диаграмма Спагетти' (без диаграммы)")
    ap.add_argument('-i', '--input', required=True)
    ap.add_argument('-s', '--sheet', default=None)
    ap.add_argument('-o', '--output')
    args = ap.parse_args()
    payload = kpsc_parse_spaghetti_sheet_build_payload(Path(args.input), args.sheet)
    data = json.dumps(payload, ensure_ascii=False, indent=2, default=str)
    if args.output:
        Path(args.output).write_text(data, encoding='utf-8')
        print(f'Wrote {args.output}')
    else:
        print(data)

kpsc_parse_spaghetti_sheet_module = SimpleNamespace(HEADER_PHRASES=kpsc_parse_spaghetti_sheet_HEADER_PHRASES, find_header_row=kpsc_parse_spaghetti_sheet_find_header_row, compute_col_bounds=kpsc_parse_spaghetti_sheet_compute_col_bounds, find_bottom_row=kpsc_parse_spaghetti_sheet_find_bottom_row, build_merged_lookup=kpsc_parse_spaghetti_sheet_build_merged_lookup, extract_table=kpsc_parse_spaghetti_sheet_extract_table, extract_pre_table=kpsc_parse_spaghetti_sheet_extract_pre_table, build_payload=kpsc_parse_spaghetti_sheet_build_payload, parse=kpsc_parse_spaghetti_sheet_parse, main=kpsc_parse_spaghetti_sheet_main)

# END_SOURCE_KPSC_PARSE_SPAGHETTI_SHEET

# START_SOURCE_KARTOCHKA_PROEKTA_PARSE_DROPDOWN
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/parser_scripts/parse_dropdown.py.
def kartochka_proekta_parse_dropdown__find_sheet(wb, keyword: str):
    """Поиск листа по подстроке в имени."""
    for name in wb.sheetnames:
        if keyword.lower() in name.lower():
            return wb[name]
    return None

def kartochka_proekta_parse_dropdown_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """
    Парсит лист 'выпадающий список' — справочник допустимых единиц измерения.

    Строка 1 — заголовки (названия категорий): A, B, C, D.
    Строки 2+ — значения единиц измерения по столбцам.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    ws = kartochka_proekta_parse_dropdown__find_sheet(wb, 'выпадающий список')
    if ws is None:
        wb.close()
        result = {'categories': {}}
        output_file = output_dir / 'dropdown_units.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        return result
    categories = {}
    columns = ['A', 'B', 'C', 'D']
    headers = []
    for col in columns:
        val = ws[f'{col}1'].value
        header = str(val).strip() if val else None
        headers.append(header)
    for col, header in zip(columns, headers):
        if header is None:
            continue
        units = []
        for row in range(2, ws.max_row + 1):
            val = ws[f'{col}{row}'].value
            if val is not None:
                unit = str(val).strip()
                if unit:
                    units.append(unit)
        categories[header] = units
    wb.close()
    result = {'categories': categories}
    output_file = output_dir / 'dropdown_units.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return result

kartochka_proekta_parse_dropdown_module = SimpleNamespace(_find_sheet=kartochka_proekta_parse_dropdown__find_sheet, parse=kartochka_proekta_parse_dropdown_parse)

# END_SOURCE_KARTOCHKA_PROEKTA_PARSE_DROPDOWN

# START_SOURCE_KARTOCHKA_PROEKTA_PARSE_KARTOCHKA_MAIN
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/parser_scripts/parse_kartochka_main.py.
def kartochka_proekta_parse_kartochka_main__find_sheet(wb, keyword: str, exclude: str='ШАБЛОН'):
    """Поиск листа по подстроке в имени, исключая шаблоны."""
    for name in wb.sheetnames:
        if keyword.lower() in name.lower() and exclude.upper() not in name.upper():
            return wb[name]
    if len(wb.sheetnames) > 1:
        return wb[wb.sheetnames[1]]
    return wb[wb.sheetnames[0]]

def kartochka_proekta_parse_kartochka_main__cell_str(ws, coord: str) -> str:
    """Получить строковое значение ячейки, пустая → ''."""
    val = ws[coord].value
    if val is None:
        return ''
    return str(val).strip()

def kartochka_proekta_parse_kartochka_main__cell_value(ws, coord: str) -> Any:
    """Получить значение ячейки as-is (число, дата, строка, None)."""
    return ws[coord].value

def kartochka_proekta_parse_kartochka_main__format_date(val) -> Optional[str]:
    """Преобразование даты в ISO-строку."""
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.strftime('%Y-%m-%d')
    s = str(val).strip()
    if not s:
        return None
    return s

def kartochka_proekta_parse_kartochka_main__extract_fio(raw: str) -> str:
    """Извлечение ФИО из строки вида '____________ И.И. Иванов'."""
    cleaned = re.sub('_+', '', raw).strip()
    return cleaned

def kartochka_proekta_parse_kartochka_main_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """
    Парсит лист 'Карточка проекта' из XLSX-файла.

    Ищет лист по подстроке 'Карточка проекта', исключая 'ШАБЛОН'.
    Возвращает и сохраняет структурированный JSON.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    ws = kartochka_proekta_parse_kartochka_main__find_sheet(wb, 'Карточка проекта')
    result = {'meta': {'workbook': Path(xlsx_path).name, 'sheet': ws.title, 'max_row': ws.max_row}, 'header': {}, 'section1': {}, 'section2': {}, 'indicators': [], 'indicator_dates': {}, 'events': []}
    result['header'] = {'org_name': kartochka_proekta_parse_kartochka_main__cell_str(ws, 'B2'), 'project_name': kartochka_proekta_parse_kartochka_main__cell_str(ws, 'B4'), 'signee_position': kartochka_proekta_parse_kartochka_main__cell_str(ws, 'K4'), 'signee_name': kartochka_proekta_parse_kartochka_main__extract_fio(kartochka_proekta_parse_kartochka_main__cell_str(ws, 'K7')), 'signee_date': kartochka_proekta_parse_kartochka_main__cell_str(ws, 'K8'), 'has_utverzhday': 'утверждаю' in kartochka_proekta_parse_kartochka_main__cell_str(ws, 'K3').lower()}
    leader_raw = kartochka_proekta_parse_kartochka_main__cell_str(ws, 'C15')
    leader = re.sub('^Руководитель\\s+проекта\\s*:\\s*', '', leader_raw, flags=re.IGNORECASE).strip()
    team_raw = kartochka_proekta_parse_kartochka_main__cell_str(ws, 'C16')
    team = re.sub('^Команда\\s+проекта\\s*:\\s*', '', team_raw, flags=re.IGNORECASE).strip()
    result['section1'] = {'clients': kartochka_proekta_parse_kartochka_main__cell_str(ws, 'E11'), 'perimeter': kartochka_proekta_parse_kartochka_main__cell_str(ws, 'E12'), 'owner': kartochka_proekta_parse_kartochka_main__cell_str(ws, 'E13'), 'boundaries': kartochka_proekta_parse_kartochka_main__cell_str(ws, 'E14'), 'leader': leader, 'team': team}
    result['section2'] = {'key_risk': kartochka_proekta_parse_kartochka_main__cell_str(ws, 'M11'), 'justification': kartochka_proekta_parse_kartochka_main__cell_str(ws, 'M13')}
    result['indicator_dates'] = {'base_date': kartochka_proekta_parse_kartochka_main__format_date(kartochka_proekta_parse_kartochka_main__cell_value(ws, 'F21')), 'target_date': kartochka_proekta_parse_kartochka_main__format_date(kartochka_proekta_parse_kartochka_main__cell_value(ws, 'G21')), 'ideal_date': kartochka_proekta_parse_kartochka_main__format_date(kartochka_proekta_parse_kartochka_main__cell_value(ws, 'H21'))}
    for row in range(22, 40):
        c_val = kartochka_proekta_parse_kartochka_main__cell_value(ws, f'C{row}')
        if c_val is None:
            break
        try:
            num = int(c_val)
        except (ValueError, TypeError):
            break
        d_val = kartochka_proekta_parse_kartochka_main__cell_str(ws, f'D{row}')
        e_val = kartochka_proekta_parse_kartochka_main__cell_str(ws, f'E{row}')
        f_val = kartochka_proekta_parse_kartochka_main__cell_value(ws, f'F{row}')
        g_val = kartochka_proekta_parse_kartochka_main__cell_value(ws, f'G{row}')
        h_val = kartochka_proekta_parse_kartochka_main__cell_value(ws, f'H{row}')
        if not d_val and (not e_val) and (f_val is None):
            continue
        result['indicators'].append({'number': num, 'name': d_val if d_val else None, 'unit': e_val if e_val else None, 'base_value': f_val, 'target_value': g_val, 'ideal_value': h_val})
    for row in range(20, 40):
        k_val = kartochka_proekta_parse_kartochka_main__cell_str(ws, f'K{row}')
        if not k_val:
            q_check = kartochka_proekta_parse_kartochka_main__cell_value(ws, f'Q{row}')
            if q_check is None:
                continue
        if not k_val and kartochka_proekta_parse_kartochka_main__cell_value(ws, f'Q{row}') is None:
            continue
        q_val = kartochka_proekta_parse_kartochka_main__cell_value(ws, f'Q{row}')
        s_val = kartochka_proekta_parse_kartochka_main__cell_value(ws, f'S{row}')
        if k_val and 'ключевые события' in k_val.lower():
            continue
        result['events'].append({'name': k_val.strip() if k_val else '', 'start_date': kartochka_proekta_parse_kartochka_main__format_date(q_val), 'end_date': kartochka_proekta_parse_kartochka_main__format_date(s_val)})
    wb.close()
    output_file = output_dir / 'kartochka_main.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return result

kartochka_proekta_parse_kartochka_main_module = SimpleNamespace(_find_sheet=kartochka_proekta_parse_kartochka_main__find_sheet, _cell_str=kartochka_proekta_parse_kartochka_main__cell_str, _cell_value=kartochka_proekta_parse_kartochka_main__cell_value, _format_date=kartochka_proekta_parse_kartochka_main__format_date, _extract_fio=kartochka_proekta_parse_kartochka_main__extract_fio, parse=kartochka_proekta_parse_kartochka_main_parse)

# END_SOURCE_KARTOCHKA_PROEKTA_PARSE_KARTOCHKA_MAIN

# START_SOURCE_KARTOCHKA_PROEKTA_PARSE_METODIKA
# PURPOSE: Inlined source from audit_engine/kartochka_proekta/parser_scripts/parse_metodika.py.
def kartochka_proekta_parse_metodika__find_sheet(wb, keyword: str, exclude: str='ШАБЛОН'):
    """Поиск листа по подстроке в имени, исключая шаблоны."""
    for name in wb.sheetnames:
        if keyword.lower() in name.lower() and exclude.upper() not in name.upper():
            return wb[name]
    return None

def kartochka_proekta_parse_metodika__cell_str(ws, coord: str) -> str:
    """Получить строковое значение ячейки."""
    val = ws[coord].value
    if val is None:
        return ''
    return str(val).strip()

def kartochka_proekta_parse_metodika__is_field_label(text: str, label: str) -> bool:
    """Проверяет, начинается ли текст с метки поля (напр. 'Единицы измерения:')."""
    return text.lower().startswith(label.lower())

def kartochka_proekta_parse_metodika_parse(xlsx_path: Path, output_dir: Path) -> dict:
    """
    Парсит лист 'Методика расчета' из XLSX-файла.

    Алгоритм: итерируем строки, ищем названия показателей (не являющиеся метками полей),
    затем читаем 3 строки ниже: единицы, способ расчёта, источник данных.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    ws = kartochka_proekta_parse_metodika__find_sheet(wb, 'Методика расчет')
    if ws is None:
        wb.close()
        result = {'meta': {'sheet': None}, 'header': {}, 'indicators': []}
        output_file = output_dir / 'metodika.json'
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        return result
    result = {'meta': {'sheet': ws.title}, 'header': {'org_name': kartochka_proekta_parse_metodika__cell_str(ws, 'B2'), 'project_name': kartochka_proekta_parse_metodika__cell_str(ws, 'B4')}, 'indicators': []}
    field_labels = ['единицы измерения', 'способ расчет', 'источник данных', 'методика расчет']
    skip_keywords = ['карточка проекта', 'методика расчет']
    row = 1
    max_row = ws.max_row
    while row <= max_row:
        b_val = kartochka_proekta_parse_metodika__cell_str(ws, f'B{row}')
        if not b_val:
            row += 1
            continue
        b_lower = b_val.lower()
        is_label = any((b_lower.startswith(lbl) for lbl in field_labels))
        is_skip = any((kw in b_lower for kw in skip_keywords))
        if is_label or is_skip:
            row += 1
            continue
        next_b = kartochka_proekta_parse_metodika__cell_str(ws, f'B{row + 1}') if row + 1 <= max_row else ''
        if kartochka_proekta_parse_metodika__is_field_label(next_b, 'единицы измерения'):
            indicator_name = re.sub(':$', '', b_val).strip()
            unit = kartochka_proekta_parse_metodika__cell_str(ws, f'F{row + 1}')
            calc_method = kartochka_proekta_parse_metodika__cell_str(ws, f'F{row + 2}') if row + 2 <= max_row else ''
            data_source = kartochka_proekta_parse_metodika__cell_str(ws, f'F{row + 3}') if row + 3 <= max_row else ''
            result['indicators'].append({'name': indicator_name, 'unit': unit if unit else None, 'calc_method': calc_method if calc_method else None, 'data_source': data_source if data_source else None, 'row_start': row})
            row += 4
            continue
        row += 1
    wb.close()
    output_file = output_dir / 'metodika.json'
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    return result

kartochka_proekta_parse_metodika_module = SimpleNamespace(_find_sheet=kartochka_proekta_parse_metodika__find_sheet, _cell_str=kartochka_proekta_parse_metodika__cell_str, _is_field_label=kartochka_proekta_parse_metodika__is_field_label, parse=kartochka_proekta_parse_metodika_parse)

# END_SOURCE_KARTOCHKA_PROEKTA_PARSE_METODIKA
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """Вызов OpenAI API"""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(model=os.environ.get('LLM_MODEL', 'openai/gpt-oss-120b'), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов КПСЦ. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    result_text = response.choices[0].message.content
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
    """LLM-проверка осмысленности названия проекта/потока."""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    prompt = f'Ты эксперт по проверке документов «Карточка проекта» в рамках бережливого производства.\n\nТРЕБОВАНИЕ:\n{rule['requirement_expert']}\n\nКРИТЕРИИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nНазвание проекта/потока: "{project_name}"\n\nЗАДАНИЕ:\nПроверь, является ли название проекта осмысленным текстом, описывающим реальный проект или поток.\nНе является осмысленным: placeholder ("Название проекта"), набор символов ("ааааа"), слишком общий текст ("тест").\nЯвляется осмысленным: конкретное описание проекта ("Оптимизация производства приборов учёта").\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{kartochka_proekta_validate_3_flow_name_RULE_INDEX}",\n  "rule_title": "{kartochka_proekta_validate_3_flow_name_RULE_TITLE}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n'
    response = client.chat.completions.create(model=resolve_runtime_llm_model(os.environ.get('LLM_MODEL', 'gpt-4.1-mini')), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    return json.loads(response.choices[0].message.content)

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
    """LLM-проверка осмысленности обоснования и ключевого риска."""
    base_url = os.environ.get('LLM_BASE_URL', 'http://localhost:8001/v1/')
    client = OpenAI(api_key=api_key, base_url=base_url)
    prompt = f'Ты эксперт по проверке документов «Карточка проекта» в рамках бережливого производства.\n\nТРЕБОВАНИЕ:\n{rule['requirement_expert']}\n\nКРИТЕРИИ:\n- Что проверять: {rule['validation_criteria']['what_to_check']}\n- Условие успеха: {rule['validation_criteria']['success_condition']}\n- Условие ошибки: {rule['validation_criteria']['error_condition']}\n\nФАКТИЧЕСКИЕ ДАННЫЕ:\nКлючевой риск (M11): "{key_risk}"\nОбоснование выбора потока (M13): "{justification}"\n\nЗАДАНИЕ:\nПроверь, содержат ли оба поля осмысленный текст:\n- Ключевой риск — должен описывать конкретный риск проекта (например: "Срыв сроков", "Потеря клиентов").\n  НЕ осмысленный: "-", "нет", "риск", набор символов.\n- Обоснование — должно содержать аргументацию выбора потока (например: "Наличие ожидания в потоке, несвоевременная подготовка").\n  НЕ осмысленный: "-", "обоснование", "тест", набор символов.\n\nФОРМАТ ОТВЕТА (строго JSON):\n{{\n  "rule_index": "{kartochka_proekta_validate_6_justification_RULE_INDEX}",\n  "rule_title": "{kartochka_proekta_validate_6_justification_RULE_TITLE}",\n  "status": "PASS или FAIL",\n  "discrepancy": "Описание проблемы если FAIL, иначе пустая строка"\n}}\n'
    response = client.chat.completions.create(model=resolve_runtime_llm_model(os.environ.get('LLM_MODEL', 'gpt-4.1-mini')), messages=[{'role': 'system', 'content': 'Ты эксперт по валидации документов. Отвечаешь строго в формате JSON.'}, {'role': 'user', 'content': prompt}], temperature=0, response_format={'type': 'json_object'})
    return json.loads(response.choices[0].message.content)

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

def resolve_runtime_llm_model(model_name: Optional[str]) -> str:
    """Map stale cloud model ids to the local model exposed by the repo LLM endpoint."""
    fallback = os.environ.get('LLM_MODEL', 'Qwen3.5-35B-A3B')
    if not model_name:
        return fallback
    lowered = model_name.lower()
    if lowered.startswith('openai/') or lowered.startswith('gpt-'):
        return fallback
    return model_name

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
        self.prompts_dir = log_dir / 'prompts'
        self.responses_dir = log_dir / 'responses'
        self.parsed_dir = log_dir / 'parsed_docs'
        self.prompts_dir.mkdir(exist_ok=True)
        self.responses_dir.mkdir(exist_ok=True)
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

    def log_rule_prompt(self, rule_index: int, system_prompt: str, user_prompt: str):
        """Сохраняет промпт для правила."""
        filepath = self.prompts_dir / f'rule_{rule_index:02d}_prompt.txt'
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f'{'=' * 80}\n')
            f.write(f'RULE #{rule_index} - PROMPT\n')
            f.write(f'{'=' * 80}\n\n')
            f.write(f'--- SYSTEM PROMPT ---\n')
            f.write(system_prompt)
            f.write(f'\n\n--- USER PROMPT ---\n')
            f.write(user_prompt)
        self.log(f'📝 Промпт для правила #{rule_index} сохранён: {filepath}')

    def log_rule_response(self, rule_index: int, raw_response: str, parsed_violations: List[Dict]):
        """Сохраняет ответ LLM для правила."""
        filepath = self.responses_dir / f'rule_{rule_index:02d}_response.txt'
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f'{'=' * 80}\n')
            f.write(f'RULE #{rule_index} - LLM RESPONSE\n')
            f.write(f'{'=' * 80}\n\n')
            f.write(f'--- RAW RESPONSE ---\n')
            f.write(raw_response)
            f.write(f'\n\n--- PARSED VIOLATIONS ---\n')
            f.write(json.dumps(parsed_violations, ensure_ascii=False, indent=2))
        self.log(f'✅ Ответ для правила #{rule_index} сохранён: {filepath}')

    def log_non_llm_result(self, rule_index: int, violations: List[Dict]):
        """Сохраняет результат non-LLM проверки."""
        filepath = self.responses_dir / f'rule_{rule_index:02d}_non_llm.txt'
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f'{'=' * 80}\n')
            f.write(f'RULE #{rule_index} - NON-LLM CHECK\n')
            f.write(f'{'=' * 80}\n\n')
            f.write(f'--- RESULT ---\n')
            f.write(json.dumps(violations, ensure_ascii=False, indent=2))
        self.log(f'✅ Результат non-LLM правила #{rule_index} сохранён: {filepath}')

    def log_final_results(self, violations: List[Dict]):
        """Сохраняет финальные результаты."""
        filepath = self.log_dir / 'final_results.json'
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(violations, f, ensure_ascii=False, indent=2)
        self.log(f'📊 Финальные результаты сохранены: {filepath}')

    def log_error(self, rule_index: int, error: str):
        """Логирует ошибку для правила."""
        filepath = self.responses_dir / f'rule_{rule_index:02d}_error.txt'
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f'ERROR for rule #{rule_index}:\n{error}')
        self.log(f'❌ ОШИБКА для правила #{rule_index}: {error}')

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

# START_SOURCE_CONTEXT_BUILDER
# PURPOSE: Inlined source from audit_engine/context_builder.py.
def extract_matching_paragraphs(text: str, patterns: List[str], headers_only: bool=False) -> str:
    """
    Извлекает из текста только строки/пункты, соответствующие паттернам.

    Используется для context_filter — сужение контекста перед отправкой в LLM.

    Args:
        text: исходный текст чанка
        patterns: список regex-паттернов для фильтрации
        headers_only: если True — только заголовки, иначе заголовок + тело до следующей секции
    """
    if not patterns:
        return text
    lines = text.split('\n')
    result_lines = []
    compiled_patterns = [re.compile(p) for p in patterns]
    if headers_only:
        for line in lines:
            stripped = line.strip()
            if any((p.match(stripped) for p in compiled_patterns)):
                result_lines.append(stripped)
    else:
        capturing = False
        new_section_pattern = re.compile('^(\\d+\\.|\\d+\\.\\d+\\.?)\\s')
        for line in lines:
            stripped = line.strip()
            matches_our_pattern = any((p.match(stripped) for p in compiled_patterns))
            if matches_our_pattern:
                capturing = True
                result_lines.append(line)
            elif capturing:
                if new_section_pattern.match(stripped) and (not matches_our_pattern):
                    capturing = False
                else:
                    result_lines.append(line)
    filtered_text = '\n'.join(result_lines).strip()
    if not filtered_text:
        return f'[Фильтр: не найдено пунктов по паттернам {patterns}]'
    return filtered_text

def build_context_for_rule(spec: RuleSpec, target_doc: Dict[str, Any], template_doc: Dict[str, Any], preprocessors: Optional[Dict[str, Callable[[str], str]]]=None) -> str:
    """
    Строит контекст для правила в зависимости от типа проверки.

    Args:
        spec: спецификация правила
        target_doc: распарсенный целевой документ
        template_doc: распарсенный шаблон
        preprocessors: словарь {scope: preprocess_fn} для нормализации текста

    Логика по compare:
    - target_only: только TARGET_<scope>
    - template: TARGET_<scope> + TEMPLATE_<scope> (с нормализацией)
    - cross_check: несколько TARGET_<scope> для сравнения между собой
    """
    context_parts = []
    scopes = [spec.scope] if isinstance(spec.scope, str) else spec.scope
    preprocessors = preprocessors or {}

    def apply_filter(content: str, scope: str) -> str:
        """Применяет context_filter если задан для данного scope."""
        if spec.context_filter and scope in spec.context_filter:
            patterns = spec.context_filter[scope]
            headers_only = spec.context_filter_mode == 'headers_only'
            return extract_matching_paragraphs(content, patterns, headers_only)
        return content

    def apply_preprocessor(content: str, scope: str) -> str:
        """Применяет препроцессор если зарегистрирован для scope."""
        fn = preprocessors.get(scope)
        if fn:
            return fn(content)
        return content

    def apply_max_chars(content: str) -> str:
        """Обрезает текст до max_chars если задано в правиле."""
        if spec.max_chars and len(content) > spec.max_chars:
            return content[:spec.max_chars] + '\n[...текст обрезан...]'
        return content
    if spec.compare == 'target_only':
        for scope in scopes:
            content = target_doc.get(scope, '')
            content = apply_max_chars(str(content))
            content = apply_filter(content, scope)
            context_parts.append(f'[TARGET_{scope}]')
            context_parts.append(content)
            context_parts.append(f'[/TARGET_{scope}]')
    elif spec.compare == 'template':
        for scope in scopes:
            target_content = apply_max_chars(str(target_doc.get(scope, '')))
            template_content = str(template_doc.get(scope, ''))
            target_content = apply_filter(target_content, scope)
            template_content = apply_filter(template_content, scope)
            target_content = apply_preprocessor(target_content, scope)
            template_content = apply_preprocessor(template_content, scope)
            context_parts.append(f'[TARGET_{scope}]')
            context_parts.append(target_content)
            context_parts.append(f'[/TARGET_{scope}]')
            context_parts.append(f'[TEMPLATE_{scope}]')
            context_parts.append(template_content)
            context_parts.append(f'[/TEMPLATE_{scope}]')
    elif spec.compare == 'cross_check':
        for scope in scopes:
            content = str(target_doc.get(scope, ''))
            content = apply_max_chars(content)
            content = apply_filter(content, scope)
            context_parts.append(f'[TARGET_{scope}]')
            context_parts.append(content)
            context_parts.append(f'[/TARGET_{scope}]')
    return '\n'.join(context_parts)

def build_user_prompt(spec: RuleSpec, target_doc: Dict[str, Any], template_doc: Dict[str, Any], preprocessors: Optional[Dict[str, Callable[[str], str]]]=None) -> str:
    """
    Формирует полный user prompt для LLM.

    Структура:
        RULE_INDEX: <номер>
        COMPARE: <тип>
        SCOPE: <чанки>
        RULE_TITLE: <заголовок>
        RULE_INSTRUCTIONS:
        - инструкция 1
        - инструкция 2

        CONTEXT:
        [TARGET_scope] ... [/TARGET_scope]
        [TEMPLATE_scope] ... [/TEMPLATE_scope]
    """
    scope_str = spec.scope if isinstance(spec.scope, str) else ', '.join(spec.scope)
    instructions_text = '\n'.join((f'- {instr}' for instr in spec.instructions))
    context = build_context_for_rule(spec, target_doc, template_doc, preprocessors)
    prompt = f'RULE_INDEX: {spec.index}\nCOMPARE: {spec.compare}\nSCOPE: {scope_str}\nRULE_TITLE: {spec.title}\nRULE_INSTRUCTIONS:\n{instructions_text}\n\nCONTEXT:\n{context}\n'
    return prompt

context_builder_module = SimpleNamespace(extract_matching_paragraphs=extract_matching_paragraphs, build_context_for_rule=build_context_for_rule, build_user_prompt=build_user_prompt)

# END_SOURCE_CONTEXT_BUILDER

# START_SOURCE_LLM_CLIENT
# PURPOSE: Inlined source from audit_engine/llm_client.py.
def sanitize_json_string(s: str) -> str:
    """
    Экранирует неэкранированные переносы строк внутри JSON-строк.

    LLM иногда возвращает JSON с реальными \\n внутри строковых значений,
    что ломает json.loads(). Эта функция проходит по тексту посимвольно,
    отслеживая состояние "внутри строки" / "вне строки", и заменяет
    сырые \\n, \\r, \\t на их escaped-версии.
    """
    result_chars = []
    in_string = False
    escape_next = False
    for char in s:
        if escape_next:
            result_chars.append(char)
            escape_next = False
            continue
        if char == '\\' and in_string:
            result_chars.append(char)
            escape_next = True
            continue
        if char == '"':
            in_string = not in_string
            result_chars.append(char)
            continue
        if in_string and char in '\n\r\t':
            if char == '\n':
                result_chars.append('\\n')
            elif char == '\r':
                result_chars.append('\\r')
            elif char == '\t':
                result_chars.append('\\t')
        else:
            result_chars.append(char)
    return ''.join(result_chars)

def parse_json_response(raw_response: str, spec_index: int, spec_title: str) -> List[Dict[str, Any]]:
    """
    Парсит JSON-ответ от LLM.

    Обрабатывает:
    - Markdown код-блоки (```json ... ```)
    - Битый JSON (переносы строк внутри строк)
    - Формат {"status": "ok"} → пустой список
    - Формат {"status": "fail", "нарушения": [...]} → список нарушений

    Args:
        raw_response: сырой ответ от LLM
        spec_index: номер правила (для fallback)
        spec_title: название правила (для fallback)

    Returns:
        Список нарушений. Пустой список если status=ok.
    """
    text = raw_response.strip()
    if text.startswith('```'):
        lines = text.split('\n')
        if lines[0].startswith('```'):
            lines = lines[1:]
        if lines and lines[-1].strip() == '```':
            lines = lines[:-1]
        text = '\n'.join(lines)
    text = sanitize_json_string(text)
    try:
        result = json.loads(text)
        if isinstance(result, dict):
            if result.get('status') == 'ok':
                return []
            violations = result.get('нарушения', [])
            for v in violations:
                v['rule_index'] = result.get('rule_index', spec_index)
                v['rule_title'] = result.get('rule_title', spec_title)
            return violations
        if isinstance(result, list):
            return result
        return []
    except json.JSONDecodeError as e:
        print(f'[WARN] Не удалось распарсить JSON: {e}', file=sys.stderr)
        print(f'[WARN] Ответ: {text[:200]}...', file=sys.stderr)
        return []

def call_llm(messages: List[Dict[str, str]], model: str, temperature: float=0.0, base_url: Optional[str]=None, max_tokens: Optional[int]=None, reasoning_effort: Optional[str]=None, seed: Optional[int]=None) -> str:
    """
    Вызывает LLM через OpenAI Chat API.

    Args:
        messages: список сообщений [{role, content}]
        model: идентификатор модели (gpt-4.1-mini, openai/gpt-oss-120b и т.д.)
        temperature: температура генерации (0.0 = детерминированный ответ)
        base_url: URL API (None = облачный OpenAI из env)
        max_tokens: лимит токенов генерации (None = по умолчанию провайдера)
        reasoning_effort: уровень reasoning для моделей gpt-oss ("low"/"medium"/"high")
        seed: фиксированный seed для воспроизводимости (особенно важен для MoE-моделей)

    Returns:
        Текст ответа LLM.
    """
    if base_url:
        model = resolve_runtime_llm_model(model)
        client = OpenAI(base_url=base_url, api_key='none')
    else:
        client = OpenAI()
    kwargs: Dict[str, Any] = {'model': model, 'messages': messages, 'temperature': temperature}
    if max_tokens is not None:
        kwargs['max_tokens'] = max_tokens
    if seed is not None:
        kwargs['seed'] = seed
    if reasoning_effort:
        kwargs['extra_body'] = {'reasoning_effort': reasoning_effort}
    response = client.chat.completions.create(**kwargs)
    return response.choices[0].message.content or ''

llm_client_module = SimpleNamespace(sanitize_json_string=sanitize_json_string, parse_json_response=parse_json_response, call_llm=call_llm)

# END_SOURCE_LLM_CLIENT

# START_SOURCE_MULTI_RULE
# PURPOSE: Inlined source from audit_engine/multi_rule.py.
multi_rule_logger = logging.getLogger(__name__)

SYSTEM_PROMPT = 'Ты — строгий аудитор документов.\n\nВ одном запросе тебе даётся ВЕСЬ документ и СПИСОК ПРАВИЛ для проверки.\n\nДокумент логически разделён на именованные СЕКЦИИ. Для каждой секции даны границы (маркеры начала/конца) и описание. Секции НЕ размечены в тексте — ты должен определить их сам по маркерам.\n\nДля каждого правила указана ЦЕЛЕВАЯ СЕКЦИЯ (target_section). Анализируй ТОЛЬКО эту секцию документа, даже если похожий текст встречается в других местах. Если у правила несколько target_sections — проверяй их совокупно (достаточно выполнения в любой из них, если явно не сказано иное).\n\nФормат ответа — СТРОГО JSON-массив, ПО ОДНОМУ объекту на правило, в том же порядке, что и в списке правил. У каждого объекта ДВА обязательных поля: «reasoning» и «verdict».\n\nПример для ok:\n{\n  "rule_index": 1,\n  "reasoning": "Краткое обоснование в 1-3 предложения.",\n  "verdict": {\n    "status": "ok"\n  }\n}\n\nПример для fail:\n{\n  "rule_index": 3,\n  "reasoning": "Краткое обоснование: где и что не так.",\n  "verdict": {\n    "status": "fail",\n    "нарушения": [\n      {"Целевой документ": "<цитата>", "Различие": "<что не так>"}\n    ]\n  }\n}\n\nПоля:\n- «rule_index» — номер правила из списка (обязательно)\n- «reasoning» — КРАТКОЕ рассуждение (1-3 предложения): какую секцию смотрел, что нашёл/не нашёл, почему такой вердикт. Без пересказа содержимого документа и правила.\n- «verdict» — объект с финальным вердиктом (ТОЛЬКО структурированный, никаких рассуждений):\n    - «status»: «ok» / «fail» / «error»\n    - «нарушения»: массив {«Целевой документ», «Различие»}, только при «fail»\n    - «Целевой документ»: цитата из документа или «отсутствует»\n    - «Различие»: что не так или чего не хватает\n- «error» — status для случая, когда правило технически невозможно проверить (нет нужной секции в документе). Добавь поле «reason».\n\nВАЖНО:\n- Ответ — ТОЛЬКО JSON-массив, без markdown-ограждения и без пояснений до/после массива.\n- В поле «verdict» — СТРОГО структурированные данные. Никаких «проверю ещё раз», «пересмотрю». Все такие мысли — в «reasoning».\n- В «reasoning» НЕ включай JSON-объекты и не пытайся там формировать ответ — это свободный текст для размышлений.\n- Не пропускай правила — в массиве должно быть ровно столько объектов, сколько правил.\n- Не смешивай правила между собой.\n- Игнорируй OCR-артефакты (пробелы между буквами, склейку строк, дублирование) — оценивай смысл.\n- Уважай target_section правила: не ищи нарушения вне указанной секции.'

def multi_rule_build_user_prompt(filename: str, doc_text: str, sections: dict, rules: list) -> str:
    """Собирает user-prompt из FILENAME + DOCUMENT + SECTIONS + RULES."""
    parts = ['## FILENAME', filename, '']
    parts += ['## DOCUMENT', doc_text, '']
    parts.append('## SECTIONS')
    for name, meta in sections.items():
        parts.append(f'- **{name}**: {meta['description']}')
        parts.append(f'    start: {meta['start']}')
        parts.append(f'    end: {meta['end']}')
    parts.append('')
    parts.append('## RULES')
    for rule in rules:
        parts.append(f'### RULE {rule['index']} — {rule['title']}')
        if 'target_section' in rule:
            parts.append(f'target_section: {rule['target_section']}')
        elif 'target_sections' in rule:
            parts.append(f'target_sections: {', '.join(rule['target_sections'])}')
        parts.append(f'check: {rule['check']}')
        if rule.get('exclusions'):
            parts.append(f'exclusions: {rule['exclusions']}')
        parts.append('')
    return '\n'.join(parts)

def _collect_doc_text(parsed: Dict[str, Any], include_scopes: Optional[List[str]]) -> str:
    """
    Собирает текст документа из parsed_docs для multi-rule.

    - Если парсер отдал `raw_text` (новый формат docx/pptx сервисов) — используем его
      целиком, игнорируя `include_scopes` и scope-ключи. Для парсера, который уже умеет
      отдавать полный текст, нарезка по скоупам теряет смысл.
    - Иначе (легаси-парсеры: Paddle, Vision, OCR и т.д.) — старое поведение: фильтр
      по `include_scopes` либо все непустые скоупы, с дедупликацией по значению.
    """
    # Новый путь: полный текст документа без scope-зависимости.
    raw_text = parsed.get('raw_text')
    if isinstance(raw_text, str) and raw_text:
        return raw_text

    # Легаси-путь: сбор по scope-ключам для парсеров, не переведённых на raw_text.
    if include_scopes:
        text_keys = [k for k in include_scopes if k in parsed]
    else:
        text_keys = [k for k in parsed.keys() if k not in ('filename', 'path')]
    unique_texts = list(dict.fromkeys((parsed[k] for k in text_keys if isinstance(parsed.get(k), str))))
    return '\n\n'.join(unique_texts) if unique_texts else ''

def _parse_response(response_text: str) -> List[Dict[str, Any]]:
    """
    Парсит JSON-массив вердиктов от LLM, снимая ```json обёртку если есть.

    Fallback: если LLM вернул невалидный JSON, пытаемся извлечь отдельные
    объекты верхнего уровня через regex — LLM иногда склеивает или ломает
    структуру между элементами массива.
    """
    import re
    clean = response_text.strip()
    if clean.startswith('```'):
        parts = clean.split('```', 2)
        if len(parts) >= 2:
            clean = parts[1]
            if clean.lstrip().startswith('json'):
                clean = clean.lstrip()[4:]
    clean = clean.strip()
    try:
        parsed = json.loads(clean)
        if isinstance(parsed, list):
            return parsed
        if isinstance(parsed, dict):
            return [parsed]
    except json.JSONDecodeError:
        pass
    verdicts = []
    i = 0
    while i < len(clean):
        if clean[i] != '{':
            i += 1
            continue
        depth = 0
        start = i
        in_str = False
        escape = False
        while i < len(clean):
            ch = clean[i]
            if escape:
                escape = False
            elif ch == '\\':
                escape = True
            elif ch == '"':
                in_str = not in_str
            elif not in_str:
                if ch == '{':
                    depth += 1
                elif ch == '}':
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
            i += 1
        obj_text = clean[start:i]
        try:
            verdicts.append(json.loads(obj_text))
        except json.JSONDecodeError:
            multi_rule_logger.warning(f'Пропущен невалидный JSON-объект: {obj_text[:100]}...')
    if verdicts:
        return verdicts
    rule_markers = [m.start() for m in re.finditer('"rule_index"\\s*:', clean)]
    if len(rule_markers) >= 2:
        boundaries = []
        for marker_pos in rule_markers:
            brace = clean.rfind('{', 0, marker_pos)
            if brace >= 0:
                boundaries.append(brace)
        boundaries.append(len(clean))
        for idx in range(len(boundaries) - 1):
            start = boundaries[idx]
            end = boundaries[idx + 1]
            chunk = clean[start:end]
            last_brace = chunk.rfind('}')
            if last_brace < 0:
                continue
            candidate = chunk[:last_brace + 1].strip().rstrip(',').strip()
            opens = candidate.count('{')
            closes = candidate.count('}')
            if closes > opens:
                extra = closes - opens
                for _ in range(extra):
                    candidate = candidate.rstrip()
                    if candidate.endswith('}'):
                        candidate = candidate[:-1].rstrip()
            elif closes < opens:
                candidate += '}' * (opens - closes)
            try:
                verdicts.append(json.loads(candidate))
            except json.JSONDecodeError:
                multi_rule_logger.warning(f'Пропущен чанк: {candidate[:80]}...')
    if not verdicts:
        raise json.JSONDecodeError('Не удалось распарсить ни одного объекта', clean, 0)
    return verdicts

def _verdict_to_violations(verdicts: List[Dict[str, Any]], rules: List[Dict[str, Any]], layer: str='base') -> List[Dict[str, Any]]:
    """
    Конвертирует массив verdicts из multi-rule формата в формат legacy violations.

    Args:
        verdicts: ответ LLM
        rules: список правил (для подстановки названий)
        layer: "base" (менеджерский слой) или "methodology" (методический слой из МР/МУ)

    Legacy format (как в старом engine.py):
        [{"правило": "...", "Целевой документ": "...", "Различие": "...", "layer": "base"}]

    Multi-rule format:
        [{"rule_index": 1, "verdict": {"status": "fail", "нарушения": [...]}, "reasoning": "..."}]
    """
    rules_by_idx = {r['index']: r for r in rules}
    violations = []
    for v in verdicts:
        idx = v.get('rule_index')
        rule = rules_by_idx.get(idx, {})
        rule_title = rule.get('title', f'Правило {idx}')
        verdict_obj = v.get('verdict', {}) if isinstance(v.get('verdict'), dict) else {}
        status = verdict_obj.get('status', '?')
        if status == 'fail':
            for violation in verdict_obj.get('нарушения', []):
                violations.append({'правило': rule_title, 'rule_index': idx, 'layer': layer, 'Целевой документ': violation.get('Целевой документ', 'отсутствует'), 'Различие': violation.get('Различие', '?')})
    return violations

def run_multi_rule_audit(*, parsed: Dict[str, Any], sections: Dict[str, Dict[str, str]], rules: List[Dict[str, Any]], include_scopes: Optional[List[str]], filename: str, llm_base_url: str, llm_model: str='Qwen3.5-35B-A3B', session_dir: Optional[Path]=None, layer: str='base') -> Dict[str, Any]:
    """
    Запускает multi-rule audit: сборка промпта → LLM → парсинг вердикта.

    Returns:
        {"verdicts": [...], "violations": [...], "usage": {...}, "prompt_chars": int}

    Side effects:
        Если указан session_dir — сохраняет system_prompt.txt, user_prompt.txt,
        response_raw.txt, response_parsed.json.
    """
    doc_text = _collect_doc_text(parsed, include_scopes)
    user_prompt = multi_rule_build_user_prompt(filename, doc_text, sections, rules)
    prefix = 'multi_rule' if layer == 'base' else f'multi_rule_{layer}'
    if session_dir:
        session_dir = Path(session_dir)
        session_dir.mkdir(parents=True, exist_ok=True)
        (session_dir / f'{prefix}_system_prompt.txt').write_text(SYSTEM_PROMPT, encoding='utf-8')
        (session_dir / f'{prefix}_user_prompt.txt').write_text(user_prompt, encoding='utf-8')
    client = OpenAI(base_url=llm_base_url, api_key='dummy')
    response = client.chat.completions.create(model=llm_model, messages=[{'role': 'system', 'content': SYSTEM_PROMPT}, {'role': 'user', 'content': user_prompt}], temperature=0.7, top_p=0.8, presence_penalty=1.5, max_tokens=8192, seed=42, extra_body={'chat_template_kwargs': {'enable_thinking': False}, 'top_k': 20, 'min_p': 0.0, 'repetition_penalty': 1.0})
    response_text = response.choices[0].message.content or ''
    if session_dir:
        (session_dir / f'{prefix}_response_raw.txt').write_text(response_text, encoding='utf-8')
    verdicts = _parse_response(response_text)
    if session_dir:
        (session_dir / f'{prefix}_response_parsed.json').write_text(json.dumps(verdicts, ensure_ascii=False, indent=2), encoding='utf-8')
    violations = _verdict_to_violations(verdicts, rules, layer=layer)
    return {'verdicts': verdicts, 'violations': violations, 'usage': {'prompt_tokens': response.usage.prompt_tokens, 'completion_tokens': response.usage.completion_tokens, 'total_tokens': response.usage.total_tokens}, 'prompt_chars': len(user_prompt)}

def load_multi_rule_config(doc_configs_dir: Path, doc_type: str) -> Optional[Dict[str, Any]]:
    """
    Загружает multi-rule конфиг (sections.json + rules_multi.json) для типа документа.

    Returns:
        {"sections": ..., "rules": ..., "include_scopes": ...} или None если multi-rule
        не настроен для этого типа.
    """
    base = Path(doc_configs_dir) / doc_type
    sections_path = base / 'sections.json'
    rules_path = base / 'rules_multi.json'
    if not sections_path.exists() or not rules_path.exists():
        return None
    sections_data = json.loads(sections_path.read_text(encoding='utf-8'))
    rules_data = json.loads(rules_path.read_text(encoding='utf-8'))
    return {'sections': sections_data['sections'], 'rules': rules_data['rules'], 'include_scopes': rules_data.get('include_scopes')}

def load_methodology_config(doc_configs_dir: Path, doc_type: str) -> Optional[Dict[str, Any]]:
    """
    Загружает методический конфиг (rules_methodology.json) для типа документа.

    Sections переиспользуются из sections.json (та же семантическая карта).

    Returns:
        {"rules": ..., "include_scopes": ..., "source": ...} или None если методический
        слой не настроен для этого типа.
    """
    base = Path(doc_configs_dir) / doc_type
    rules_path = base / 'rules_methodology.json'
    if not rules_path.exists():
        return None
    rules_data = json.loads(rules_path.read_text(encoding='utf-8'))
    return {'rules': rules_data['rules'], 'include_scopes': rules_data.get('include_scopes'), 'source': rules_data.get('source', '')}

multi_rule_module = SimpleNamespace(logger=multi_rule_logger, SYSTEM_PROMPT=SYSTEM_PROMPT, build_user_prompt=multi_rule_build_user_prompt, _collect_doc_text=_collect_doc_text, _parse_response=_parse_response, _verdict_to_violations=_verdict_to_violations, run_multi_rule_audit=run_multi_rule_audit, load_multi_rule_config=load_multi_rule_config, load_methodology_config=load_methodology_config)

# END_SOURCE_MULTI_RULE

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
    base_url = config.get('llm_base_url', 'http://localhost:8001/v1/')
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
    os.environ['LLM_BASE_URL'] = config.get('llm_base_url', 'http://localhost:8001/v1/')
    os.environ['LLM_MODEL'] = config.get('model', 'openai/gpt-oss-120b')
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
        self.system_prompt = self._load_system_prompt()
        self.preprocessors = get_preprocessors(doc_type)
        self.logger: Optional[PipelineLogger] = None

    def _load_system_prompt(self) -> str:
        """Загружает системный промпт из файла."""
        prompt_file = Path(__file__).parent / 'system_prompts' / self.config.system_prompt
        if prompt_file.exists():
            return prompt_file.read_text(encoding='utf-8')
        default_file = Path(__file__).parent / 'system_prompts' / 'default.txt'
        if default_file.exists():
            return default_file.read_text(encoding='utf-8')
        return 'Ты — строгий аудитор документов.'

    def run(self, target_path: str, template_path: Optional[str]=None, model: Optional[str]=None, temperature: Optional[float]=None, rule_filter: Optional[int]=None, parse_only: bool=False, no_cache: bool=False, print_prompts: bool=False, session_dir: Optional[str]=None, chunk_filter: Optional[str]=None, out_xlsx: Optional[str]=None, secondary_path: Optional[str]=None, progress_callback: Optional[callable]=None) -> AuditResult:
        """
        Запуск полного цикла аудита.

        Args:
            target_path: путь к целевому документу
            template_path: путь к шаблону (если не указан — из конфига)
            model: модель LLM (переопределяет конфиг)
            temperature: температура (переопределяет конфиг)
            rule_filter: проверить только одно правило
            parse_only: режим только парсинга
            no_cache: не использовать кэш шаблона
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
            self.logger.log(f'   Layout: {self.config.paddle_layout_model or 'Heron-101 (дефолт)'}')
            self.logger.log(f'   VLM: {self.config.paddle_vlm_model or 'PaddleOCR-VL-1.5 (авто)'}')
            self.logger.log(f'   VLM URL: {self.config.ocr_base_url}')
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
        template_doc: Dict[str, Any] = {}
        if parse_only:
            self.logger.log(f'✅ Режим --parse-only: парсинг завершён')
            print(f'\n{'=' * 60}')
            print('TARGET:')
            print(json.dumps(target_doc, ensure_ascii=False, indent=2))
            print(f'\n{'=' * 60}')
            print('TEMPLATE:')
            print(json.dumps(template_doc, ensure_ascii=False, indent=2))
            return AuditResult(doc_type=self.doc_type, session_dir=session_path, target_path=target_path, duration_sec=time.time() - start_time)
        available_keys = set(target_doc.keys())
        filtered_rules = []
        for r in rules:
            scopes = [r.scope] if isinstance(r.scope, str) else list(r.scope)
            if all((s == 'filename' or s in available_keys for s in scopes)):
                filtered_rules.append(r)
            else:
                missing = [s for s in scopes if s != 'filename' and s not in available_keys]
                self.logger.log(f'   ⏭️ Пропуск правила #{r.index} ({r.title}): нет чанков {missing}')
        if len(filtered_rules) < len(rules):
            self.logger.log(f'   📋 Доступно {len(filtered_rules)} из {len(rules)} правил (остальные пропущены)')
        rules = filtered_rules
        rules_summary = [{'index': r.index, 'title': r.title, 'llm': r.llm, 'compare': r.compare} for r in rules]
        with open(session_path / 'rules_summary.json', 'w', encoding='utf-8') as f:
            json.dump(rules_summary, f, ensure_ascii=False, indent=2)
        engine_mode = getattr(self.config, 'engine_mode', None) or 'legacy'
        mr_config = None
        if engine_mode == 'multi_rule':
            doc_configs_dir = Path(self.config.chunks_vision_path).parent.parent
            mr_config = load_multi_rule_config(doc_configs_dir, self.doc_type)
            if mr_config is None:
                self.logger.log(f'⚠️ engine_mode=multi_rule, но sections.json/rules_multi.json не найдены — fallback на legacy')
        if mr_config is None:
            needs_template = any((r.compare == 'template' for r in rules))
            if needs_template:
                _emit('parsing_template', {})
                template_doc = self._get_template(template_path=template_path, no_cache=no_cache, chunks_to_parse=chunks_to_parse, session_dir=session_path)
                _emit('parsing_template_done', {})
            else:
                self.logger.log('ℹ️ Template context не требуется: legacy-правила без compare=template')
        if mr_config is not None:
            self.logger.log(f'🔍 Запуск multi-rule аудита (базовый слой, {len(mr_config['rules'])} правил)...')
            _emit('checking_rules', {'total': len(mr_config['rules'])})
            mr_model = self.config.model if self.config.model and 'qwen' in self.config.model.lower() else 'Qwen3.5-35B-A3B'
            mr_result = run_multi_rule_audit(parsed=target_doc, sections=mr_config['sections'], rules=mr_config['rules'], include_scopes=mr_config.get('include_scopes'), filename=target_doc.get('filename', Path(target_path).name), llm_base_url=self.config.llm_base_url or self.config.ocr_base_url, llm_model=mr_model, session_dir=session_path, layer='base')
            violations = mr_result['violations']
            self.logger.log(f'   Base usage: prompt={mr_result['usage']['prompt_tokens']} completion={mr_result['usage']['completion_tokens']} total={mr_result['usage']['total_tokens']}')
            meth_config = load_methodology_config(doc_configs_dir, self.doc_type)
            if meth_config is not None:
                self.logger.log(f'🔍 Запуск методического слоя ({len(meth_config['rules'])} правил, источник: {meth_config.get('source', 'МР/МУ')[:80]})...')
                meth_result = run_multi_rule_audit(parsed=target_doc, sections=mr_config['sections'], rules=meth_config['rules'], include_scopes=meth_config.get('include_scopes') or mr_config.get('include_scopes'), filename=target_doc.get('filename', Path(target_path).name), llm_base_url=self.config.llm_base_url or self.config.ocr_base_url, llm_model=mr_model, session_dir=session_path, layer='methodology')
                violations.extend(meth_result['violations'])
                self.logger.log(f'   Methodology usage: prompt={meth_result['usage']['prompt_tokens']} completion={meth_result['usage']['completion_tokens']} total={meth_result['usage']['total_tokens']}')
                self.logger.log(f'   Methodology violations: {len(meth_result['violations'])}')
            _all_multi_rules = [{**r, 'layer': 'base'} for r in mr_config['rules']]
            if meth_config is not None:
                _all_multi_rules.extend(({**r, 'layer': 'methodology'} for r in meth_config['rules']))
            _emit('checking_rules_done', {'violations': len(violations)})
        else:
            _all_multi_rules = None
            self.logger.log(f'🔍 Запуск проверок...')
            violations = self._run_checks(rules=rules, target_doc=target_doc, template_doc=template_doc, model=model, temperature=temperature, print_prompts=print_prompts, progress_callback=progress_callback)
        self.logger.log_final_results(violations)
        print(json.dumps(violations, ensure_ascii=False, indent=2))
        if out_xlsx:
            xlsx_path = out_xlsx
        else:
            xlsx_path = str(session_path / 'audit_result.xlsx')
        save_to_excel(violations, xlsx_path, all_rules=rules if _all_multi_rules is None else None, multi_rules=_all_multi_rules)
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
        return AuditResult(violations=violations, doc_type=self.doc_type, session_dir=session_path, duration_sec=duration, rules_checked=len(rules), target_path=target_path)

    def _parse_document(self, file_path: str, chunks_to_parse: Optional[List[str]], vision_log_dir: Path) -> Dict[str, Any]:
        """
        Парсит документ выбранным парсером.

        Выбор парсера: `config.parser_by_ext[<расширение файла>]`. Никаких
        автоматических переопределений — один конфиг описывает всю цепочку.
        Если расширение не описано в карте, бросается ValueError.
        """
        chunk_filter_arg = chunks_to_parse[0] if chunks_to_parse and len(chunks_to_parse) == 1 else None
        file_ext = Path(file_path).suffix.lower()
        # Единый источник правды — parser_by_ext. Никаких скрытых override по формату.
        if file_ext not in self.config.parser_by_ext:
            raise ValueError(f"Для doc_type={self.doc_type} не сконфигурирован парсер для расширения {file_ext!r}. parser_by_ext={self.config.parser_by_ext}")
        parser_name = self.config.parser_by_ext[file_ext]
        if parser_name == 'docx':
            self.logger.log(f'📄 DOCX-парсинг: {file_path}...')
            doc = parse_docx(file_path, vlm_base_url=getattr(self.config, 'ocr_base_url', None))
        elif parser_name == 'pptx':
            self.logger.log(f'📄 PPTX-парсинг: {file_path}...')
            doc = parse_pptx(file_path)
        elif parser_name == 'paddle':
            self.logger.log(f'📄 Paddle-парсинг: {file_path}...')
            parser = PaddleParser(config=self.config, log_dir=str(vision_log_dir))
            doc = parser.parse(file_path, chunk_filter=chunk_filter_arg)
        else:
            # Единственные поддерживаемые парсеры сейчас — docx/pptx/paddle. Любое другое
            # значение — ошибка конфигурации и причина явно падать, а не молча продолжать.
            raise ValueError(f"Неподдерживаемый парсер {parser_name!r} для {file_ext!r}. Допустимы: docx, pptx, paddle.")
        self.logger.log_parsed_doc(doc, Path(file_path).stem)
        return doc

    def _get_template(self, template_path: Optional[str], no_cache: bool, chunks_to_parse: Optional[List[str]], session_dir: Path) -> Dict[str, Any]:
        """
        Получает шаблон: из кэша или через Vision-парсинг.

        Кэширование по SHA256:
        - Если template_cached.json существует и хэш совпадает — используем кэш
        - Иначе парсим через Vision и сохраняем кэш
        """
        if template_path:
            tpl_path = Path(template_path)
        elif self.config.template_path:
            tpl_path = self.config.template_path
        else:
            self.logger.log('⚠️ Шаблон не указан, пропускаем')
            return {}
        if not tpl_path.exists():
            self.logger.log(f'❌ Шаблон не найден: {tpl_path}')
            return {}
        self.logger.log(f'📄 Шаблон: {tpl_path}')
        file_hash = self._compute_file_hash(tpl_path)
        cache_path = tpl_path.parent / 'template_cached.json'
        # Имя парсера для шаблона резолвим из parser_by_ext по расширению файла шаблона.
        tpl_ext = tpl_path.suffix.lower()
        expected_parser = self.config.parser_by_ext.get(tpl_ext, '')
        if not no_cache and cache_path.exists():
            try:
                with open(cache_path, 'r', encoding='utf-8') as f:
                    cached = json.load(f)
                cache_parser = cached.get('_parser', '')
                if cached.get('_hash') == file_hash and cache_parser == expected_parser:
                    self.logger.log(f'✅ Используем кэш шаблона: {cache_path}')
                    cached.pop('_hash', None)
                    cached.pop('_cached_at', None)
                    cached.pop('_parser', None)
                    return cached
                elif cached.get('_hash') != file_hash:
                    self.logger.log(f'⚠️ Кэш устарел (хэш изменился), перепарсинг...')
                else:
                    self.logger.log(f'⚠️ Кэш от другого парсера ({cache_parser}→{expected_parser}), перепарсинг...')
            except (json.JSONDecodeError, KeyError):
                self.logger.log(f'⚠️ Кэш повреждён, перепарсинг...')
        self.logger.log(f'🔮 Vision-парсинг шаблона...')
        template_doc = self._parse_document(str(tpl_path), chunks_to_parse, session_dir / 'vision_template')
        cache_data = dict(template_doc)
        cache_data['_hash'] = file_hash
        cache_data['_parser'] = expected_parser
        cache_data['_cached_at'] = datetime.now().isoformat()
        with open(cache_path, 'w', encoding='utf-8') as f:
            json.dump(cache_data, f, ensure_ascii=False, indent=2)
        self.logger.log(f'💾 Кэш шаблона сохранён: {cache_path}')
        return template_doc

    @staticmethod
    def _compute_file_hash(file_path: Path) -> str:
        """Вычисляет SHA256 хэш файла."""
        sha256 = hashlib.sha256()
        with open(file_path, 'rb') as f:
            for chunk in iter(lambda: f.read(8192), b''):
                sha256.update(chunk)
        return sha256.hexdigest()

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

    def _run_checks(self, rules: List[RuleSpec], target_doc: Dict[str, Any], template_doc: Dict[str, Any], model: str, temperature: float, print_prompts: bool=False, progress_callback: Optional[callable]=None) -> List[Dict[str, Any]]:
        """
        Выполняет все проверки (non-LLM + LLM).

        Non-LLM правила выполняются синхронно.
        LLM правила выполняются параллельно через ThreadPoolExecutor.
        """
        all_violations: List[Dict[str, Any]] = []
        rules_done = 0
        total_rules = len(rules)
        for spec in rules:
            if not spec.llm:
                self.logger.log(f'🔧 Запуск non-LLM правила #{spec.index}: {spec.title}')
                check_fn = get_check(self.doc_type, spec.index)
                if check_fn:
                    try:
                        violations = check_fn(target_doc, self.config)
                    except Exception as e:
                        self.logger.log(f'   ❌ Ошибка в non-LLM проверке #{spec.index}: {e}')
                        violations = []
                else:
                    self.logger.log(f'   ⚠️ Нет зарегистрированной проверки для {self.doc_type}#{spec.index}')
                    violations = []
                all_violations.extend(violations)
                self.logger.log_non_llm_result(spec.index, violations)
                rules_done += 1
                if progress_callback:
                    progress_callback('rule_done', {'rule_index': spec.index, 'rule_title': spec.title, 'current': rules_done, 'total': total_rules, 'violations_count': len(violations)})
                if print_prompts:
                    print(f'\n{'=' * 60}')
                    print(f'[NON-LLM] Правило #{spec.index}: {spec.title}')
                    print(f'Результат: {violations}')
        llm_rules = [r for r in rules if r.llm]
        if print_prompts:
            for spec in llm_rules:
                user_prompt = build_user_prompt(spec, target_doc, template_doc, self.preprocessors)
                print(f'\n{'=' * 60}')
                print(f'[LLM] Правило #{spec.index}: {spec.title}')
                print(f'{'=' * 60}')
                print('\n--- SYSTEM PROMPT ---')
                print(self.system_prompt)
                print('\n--- USER PROMPT ---')
                print(user_prompt)
                self.logger.log_rule_prompt(spec.index, self.system_prompt, user_prompt)
            return all_violations
        max_workers = self.config.max_workers
        self.logger.log(f'🚀 Запуск {len(llm_rules)} LLM-проверок (max_workers={max_workers})')
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {executor.submit(self._run_single_llm_check, spec, target_doc, template_doc, model, temperature): spec for spec in llm_rules}
            for future in as_completed(futures):
                spec = futures[future]
                try:
                    violations = future.result()
                    all_violations.extend(violations)
                    self.logger.log(f'✅ Правило #{spec.index} проверено, нарушений: {len(violations)}')
                    rules_done += 1
                    if progress_callback:
                        progress_callback('rule_done', {'rule_index': spec.index, 'rule_title': spec.title, 'current': rules_done, 'total': total_rules, 'violations_count': len(violations)})
                except Exception as e:
                    error_msg = str(e)
                    self.logger.log_error(spec.index, error_msg)
                    print(f'[ERROR] Правило #{spec.index}: {e}', file=sys.stderr)
        return sorted(all_violations, key=lambda x: x.get('rule_index', 0))

    def _run_single_llm_check(self, spec: RuleSpec, target_doc: Dict[str, Any], template_doc: Dict[str, Any], model: str, temperature: float) -> List[Dict[str, Any]]:
        """Проверка одного правила через LLM."""
        user_prompt = build_user_prompt(spec, target_doc, template_doc, self.preprocessors)
        messages = [{'role': 'system', 'content': self.system_prompt}, {'role': 'user', 'content': user_prompt}]
        self.logger.log_rule_prompt(spec.index, self.system_prompt, user_prompt)
        raw_response = call_llm(messages, model, temperature, base_url=self.config.llm_base_url, max_tokens=self.config.llm_max_tokens, reasoning_effort=self.config.reasoning_effort, seed=self.config.llm_seed)
        violations = parse_json_response(raw_response, spec.index, spec.title)
        for obj in violations:
            obj.setdefault('rule_index', spec.index)
            obj.setdefault('rule_title', spec.title)
        self.logger.log_rule_response(spec.index, raw_response, violations)
        return violations

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
    parser.add_argument('--template', default=None, help='Путь к шаблону (опционально, по умолчанию — из конфига)')
    parser.add_argument('--model', default=None, help='Модель OpenAI (переопределяет конфиг)')
    parser.add_argument('--temperature', type=float, default=None, help='Температура генерации')
    parser.add_argument('--rule-filter', default=None, help='Проверить только указанное правило (номер: 3 для Vision, 1.1 для КПСЦ)')
    parser.add_argument('--parse-only', action='store_true', help='Только Vision-парсинг, без проверки правил')
    parser.add_argument('--no-cache', action='store_true', help='Не использовать кэш шаблона')
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
        result = engine.run(target_path=args.target, template_path=args.template, model=args.model, temperature=args.temperature, rule_filter=rule_filter, parse_only=args.parse_only, no_cache=args.no_cache, print_prompts=args.print_prompts, session_dir=args.session_dir, chunk_filter=args.chunk_filter, out_xlsx=args.out_xlsx, secondary_path=args.secondary)
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
    "grafik_obhod": parsers_grafik_obhod_parse,
}

KPSC_PARSER_MODULE_DISPATCH = {
    "parse_kpsc_header": kpsc_parse_kpsc_header_module,
    "parse_kpsc_table1": kpsc_parse_kpsc_table1_module,
    "parse_legend": kpsc_parse_legend_module,
    "parse_loss_digitization": kpsc_parse_loss_digitization_module,
    "parse_pa1_chart": kpsc_parse_pa1_chart_module,
    "parse_pa1_table": kpsc_parse_pa1_table_module,
    "parse_pokazateli": kpsc_parse_pokazateli_module,
    "parse_spaghetti_sheet": kpsc_parse_spaghetti_sheet_module,
    "parse_spaghetti_problems": kpsc_parse_spaghetti_problems_module,
}

KARTOCHKA_PARSER_MODULE_DISPATCH = {
    "parse_kartochka_main": kartochka_proekta_parse_kartochka_main_module,
    "parse_metodika": kartochka_proekta_parse_metodika_module,
    "parse_dropdown": kartochka_proekta_parse_dropdown_module,
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
    for short_name, module_ns in KPSC_PARSER_MODULE_DISPATCH.items():
        try:
            module_ns.parse(xlsx_path, output_dir)
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
    for short_name, module_ns in KARTOCHKA_PARSER_MODULE_DISPATCH.items():
        try:
            module_ns.parse(xlsx_path, output_dir)
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
    os.environ["LLM_BASE_URL"] = config.get("llm_base_url", "http://172.16.10.35:11437/v1/")
    os.environ["LLM_MODEL"] = resolve_runtime_llm_model(config.get("model", "openai/gpt-oss-120b"))
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
    self.system_prompt = self._load_system_prompt()
    self.preprocessors = get_preprocessors(doc_type)
    self.logger = None


def _audit_engine_load_system_prompt(self) -> str:
    prompt_file = SYSTEM_PROMPTS_DIR / self.config.system_prompt
    if prompt_file.exists():
        return prompt_file.read_text(encoding="utf-8")
    default_file = SYSTEM_PROMPTS_DIR / "default.txt"
    if default_file.exists():
        return default_file.read_text(encoding="utf-8")
    return "Ты — строгий аудитор документов."


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
AuditEngine._load_system_prompt = _audit_engine_load_system_prompt
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
    parser.add_argument("--template", default=None, help="Путь к шаблону (опционально, по умолчанию — из конфига)")
    parser.add_argument("--model", default=None, help="Модель OpenAI (переопределяет конфиг)")
    parser.add_argument("--temperature", type=float, default=None, help="Температура генерации")
    parser.add_argument("--rule-filter", default=None, help="Проверить только указанное правило")
    parser.add_argument("--parse-only", action="store_true", help="Только парсинг, без проверок")
    parser.add_argument("--no-cache", action="store_true", help="Не использовать кэш шаблона")
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
            template_path=args.template,
            model=args.model,
            temperature=args.temperature,
            rule_filter=rule_filter,
            parse_only=args.parse_only,
            no_cache=args.no_cache,
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
