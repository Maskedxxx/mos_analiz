#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Модели данных для аудит-движка.

Содержит датаклассы:
- RuleSpec — спецификация одного правила проверки
- AuditConfig — конфигурация аудита для типа документа
- AuditResult — результат аудита (нарушения, статистика)
"""

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Union


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
    context_filter_mode: str = "paragraphs"


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
    chunk_prefix: str = "xlsx_"


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
        config_dir — путь к папке с конфигами (автоматически)
        rules_path — путь к rules.json (автоматически)
        chunks_vision_path — путь к chunks_vision.json (автоматически)
        template_path — путь к файлу шаблона (автоматически)
        template_cached_path — путь к кэшу Vision-парсинга шаблона
    """
    doc_type: str
    doc_title: str = ""
    model: str = "gpt-4.1-mini"
    system_prompt: str = "default.txt"
    filename_pattern: str = ""
    filename_keywords: Optional[List[str]] = None
    max_workers: int = 1
    temperature: float = 0.0
    secondary_file: Optional[SecondaryFileConfig] = None
    # Выбор парсера: "vision" (облако GPT) | "ocr" (локальный HunyuanOCR)
    parser: str = "vision"
    # OCR-параметры (используются при parser="ocr")
    ocr_model: Optional[str] = None
    ocr_base_url: str = "http://localhost:8000/v1/"
    ocr_prompt: str = "提取文档图片中正文的所有信息用markdown格式表示，忽略页眉页脚。表格用html格式表达，公式用LaTeX格式表示，按照阅读顺序组织进行解析。特别注意：保留表格上方和下方的所有独立标题行和文本，不要将标题合并到表格中。"
    ocr_dpi: int = 200
    # LLM-параметры (для сверки правил через локальный vLLM)
    llm_base_url: Optional[str] = "http://localhost:8001/v1/"
    llm_max_tokens: int = 4096
    reasoning_effort: Optional[str] = None
    llm_seed: Optional[int] = None
    # Paddle-параметры (используются при parser="paddle")
    paddle_layout_model: Optional[str] = None      # HF repo Heron-101 (None = дефолт)
    paddle_layout_device: str = "cuda:0"            # GPU для layout detection
    paddle_vlm_model: Optional[str] = None          # Имя VLM (None = автоопределение)
    # OCR post-processing: имена чанков, где убирать аннотации Word-форм ("текст -> значение")
    strip_annotations_chunks: List[str] = field(default_factory=list)
    # Пути (заполняются автоматически при загрузке)
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
    doc_type: str = ""
    session_dir: Path = field(default_factory=Path)
    duration_sec: float = 0.0
    rules_checked: int = 0
    target_path: str = ""


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
    for rule in data.get("правила", []):
        spec = RuleSpec(
            index=rule["index"],
            title=rule["title"],
            scope=rule["scope"],
            compare=rule["compare"],
            llm=rule["llm"],
            instructions=rule.get("content", []),
            context_filter=rule.get("context_filter", {}),
            context_filter_mode=rule.get("context_filter_mode", "paragraphs")
        )
        rules.append(spec)

    return rules


def load_audit_config(config_dir: Path) -> AuditConfig:
    """
    Загружает AuditConfig из папки конфигов.

    Ищет config.json, rules.json, chunks_vision.json, template/*.
    """
    config_path = config_dir / "config.json"
    with open(config_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    config = AuditConfig(
        doc_type=data["doc_type"],
        doc_title=data.get("doc_title", data["doc_type"]),
        model=data.get("model", "gpt-4.1-mini"),
        system_prompt=data.get("system_prompt", "default.txt"),
        filename_pattern=data.get("filename_pattern", ""),
        filename_keywords=data.get("filename_keywords", None),
        max_workers=data.get("max_workers", 1),
        temperature=data.get("temperature", 0.0),
        # Парсер и OCR
        parser=data.get("parser", "vision"),
        ocr_model=data.get("ocr_model", None),
        ocr_base_url=data.get("ocr_base_url", "http://localhost:8000/v1/"),
        ocr_prompt=data.get("ocr_prompt", AuditConfig.ocr_prompt),
        ocr_dpi=data.get("ocr_dpi", 200),
        # Paddle-параметры
        paddle_layout_model=data.get("paddle_layout_model", None),
        paddle_layout_device=data.get("paddle_layout_device", "cuda:0"),
        paddle_vlm_model=data.get("paddle_vlm_model", None),
        # LLM-параметры для локального vLLM
        llm_base_url=data.get("llm_base_url", "http://localhost:8001/v1/"),
        llm_max_tokens=data.get("llm_max_tokens", 4096),
        reasoning_effort=data.get("reasoning_effort", None),
        llm_seed=data.get("llm_seed", None),
        strip_annotations_chunks=data.get("strip_annotations_chunks", []),
        config_dir=config_dir,
        rules_path=config_dir / "rules.json",
        chunks_vision_path=config_dir / "chunks_vision.json",
    )

    # Парсим конфиг вторичного файла (для multi-file аудитов)
    if "secondary_file" in data:
        sf = data["secondary_file"]
        config.secondary_file = SecondaryFileConfig(
            type=sf["type"],
            parser=sf["parser"],
            chunk_prefix=sf.get("chunk_prefix", "xlsx_")
        )

    # Ищем шаблон
    template_dir = config_dir / "template"
    if template_dir.exists():
        # Ищем файл шаблона (docx или pptx)
        for ext in ["*.docx", "*.pptx"]:
            templates = list(template_dir.glob(ext))
            if templates:
                config.template_path = templates[0]
                break

        # Проверяем наличие кэша
        cached = template_dir / "template_cached.json"
        if cached.exists():
            config.template_cached_path = cached

    return config
