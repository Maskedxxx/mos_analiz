# START_MODULE_CONTRACT
# PURPOSE: Публичный API пакета PDF-парсера. Внутренности (конфиг, клиенты, парсинг, оркестратор) — package-private и наружу не светятся.
# INPUTS: Импорт-время — подмодули пакета.
# OUTPUTS: Конфиги, клиенты и функции чистой логики. `parse_pdf` появится на B2.4.
# KEYWORDS: package, pdf-parser, public-api, paddle-internal.
# LINKS: src/format_parsers/pdf/_config.py, src/format_parsers/pdf/_clients.py, src/format_parsers/pdf/_parsing.py, config/parsers.json.
# RATIONALE: Снаружи пакет воспринимается как «PDF-парсер» независимо от внутренней реализации (сейчас Paddle).
# END_MODULE_CONTRACT

from __future__ import annotations

# START_REEXPORTS
from src.format_parsers.pdf._config import (
    PdfExtractorConfig,
    PdfLayoutConfig,
    PdfParserConfig,
    PdfParsingConfig,
    PdfVlmConfig,
    load_pdf_parser_config,
)
from src.format_parsers.pdf._clients import (
    LayoutDetector,
    RemoteLayoutDetector,
    VLMClient,
)
from src.format_parsers.pdf._parsing import (
    Cell,
    build_page_markdown,
    is_table_format,
    merge_text_and_tables,
    normalize_paddle_text,
    parse_paddle_table_html,
    parse_paddle_table_markdown,
    sort_by_reading_order,
)
from src.format_parsers.pdf.parse import PaddleExtractor, parse_pdf

__all__ = [
    # Публичный entry
    "parse_pdf",
    # Конфиг
    "PdfExtractorConfig",
    "PdfLayoutConfig",
    "PdfParserConfig",
    "PdfParsingConfig",
    "PdfVlmConfig",
    "load_pdf_parser_config",
    # Клиенты (IO)
    "LayoutDetector",
    "RemoteLayoutDetector",
    "VLMClient",
    # Чистая логика
    "Cell",
    "build_page_markdown",
    "is_table_format",
    "merge_text_and_tables",
    "normalize_paddle_text",
    "parse_paddle_table_html",
    "parse_paddle_table_markdown",
    "sort_by_reading_order",
    # Оркестратор (внутренний, но export для тестов/отладки)
    "PaddleExtractor",
]
# END_REEXPORTS
