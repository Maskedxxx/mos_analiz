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
from src.audit.excel_reporter import save_to_excel
from src.audit.logger import PipelineLogger
from src.audit.models import (
    AuditConfig,
    AuditResult,
    RuleSpec,
    SecondaryFileConfig,
    load_audit_config,
    load_rules,
)
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
    run_kartochka_proekta_special,
    run_kartochka_validator_module,
)
from src.doc_type_validators.kpsc import (
    KPSC_VALIDATOR_MODULE_DISPATCH,
    run_kpsc_special,
    run_kpsc_validator_module,
)
from src.doc_type_validators.drivers import run_drivers_special
from src.doc_type_validators.plan_grafik import run_plan_grafik_special
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

# END_REPORTING

# START_RUNTIME_SERVICES
# PURPOSE: Orchestrate the generic audit engine, special engines, prompt building, and LLM calls.
# INPUTS: doc_type, parsed docs, rule specs, templates, API requests, CLI args.
# OUTPUTS: AuditResult objects, session directories, and engine dispatch decisions.
# KEYWORDS: runtime, engine, multi-rule, drivers, kpsc, kartochka, plan-grafik.
# LINKS: audit_engine/engine.py, audit_engine/*/__init__.py, audit_engine/multi_rule.py.
# RATIONALE: The monolith should expose one runtime graph instead of scattered module entrypoints.






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



def get_parser(name: str) -> Callable[[str], Dict[str, str]]:
    if name not in SECONDARY_PARSER_DISPATCH:
        raise KeyError(f"Парсер '{name}' не зарегистрирован. Доступные: {sorted(SECONDARY_PARSER_DISPATCH)}")
    return SECONDARY_PARSER_DISPATCH[name]


















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




# Test-compat aliases — используются tests/test_kpsc_parsers.py.
build_kpsc_header_payload = kpsc_parse_kpsc_header_build_payload
from src.doc_type_validators.kpsc import _run_kpsc_parsers as run_kpsc_parsers

# END_RUNTIME_INTEGRATION

if __name__ == "__main__":
    main()
