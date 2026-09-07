# START_MODULE_CONTRACT
# PURPOSE: Единый конфиг парсеров проекта (Pydantic-модели + дефолтные значения). Адреса внешних сервисов (layout-детектор, OCR/VLM) читаются из окружения: LAYOUT_BASE_URL, OCR_BASE_URL, OCR_API_KEY (или `.env` в корне). Каждое поле снабжено `description` — при наведении в IDE или через `Model.model_json_schema()` видно, зачем поле нужно и как влияет.
# INPUTS: —
# OUTPUTS: Pydantic-классы (`PdfParserConfig`, `DocxHeaderOcrConfig`, `ParsersConfig`) и единственный экземпляр `CONFIG`, который используется во всём рантайме.
# KEYWORDS: config, pydantic, single-source-of-truth, no-json, annotated-fields.
# LINKS: src/format_parsers/pdf/, src/format_parsers/docx.py, main.py.
# RATIONALE: Конфиг — Python-модуль, а не внешний JSON. Плюсы: нативные docstring + description поля, автодополнение в IDE, статическая проверка, нулевая зависимость от парсеров формата. Смена окружения (другой стенд) — только через переменные окружения для адресов сервисов; параметры алгоритмов остаются в коде.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
from pathlib import Path
from typing import Dict, List, Optional

from pydantic import AliasChoices, BaseModel, ConfigDict, Field
from pydantic_settings import BaseSettings, SettingsConfigDict

# Корень проекта: config/parsers.py → parents[1]. Отсюда читается `.env`, если он есть.
_PROJECT_ROOT = Path(__file__).resolve().parents[1]
_ENV_SETTINGS = dict(
    frozen=True,
    extra="ignore",
    env_file=str(_PROJECT_ROOT / ".env"),
    env_file_encoding="utf-8",
)
# END_IMPORTS


# START_PDF_LAYOUT
class PdfLayoutConfig(BaseSettings):
    """Параметры layout-детектора — определяет регионы на странице PDF (текст, таблицы, картинки).

    `base_url` берётся из переменной окружения LAYOUT_BASE_URL (или `.env`), остальное — из кода.
    """

    model_config = SettingsConfigDict(**_ENV_SETTINGS)

    base_url: Optional[str] = Field(
        default="http://127.0.0.1:11439",
        validation_alias=AliasChoices("LAYOUT_BASE_URL"),
        description=(
            "URL удалённого Layout API (env: LAYOUT_BASE_URL). Используется `RemoteLayoutDetector` в "
            "`src/format_parsers/pdf/_clients.py`. Если `None` — запускается локальный "
            "`LayoutDetector` (Heron-101), требующий torch+GPU."
        ),
    )
    timeout: int = Field(
        default=60,
        description=(
            "Таймаут HTTP-запроса к remote Layout API в секундах. Увеличить, если сервис "
            "медленно отвечает на больших страницах."
        ),
    )
    model: str = Field(
        default="docling-project/docling-layout-heron-101",
        description=(
            "HuggingFace repo_id модели для локального layout-детектора. "
            "Игнорируется при remote-режиме (когда `base_url` задан)."
        ),
    )
    device: str = Field(
        default="cuda:0",
        description=(
            "Устройство для локального layout-детектора (`cuda:0`, `cuda:1`, `cpu`). "
            "Игнорируется при remote-режиме."
        ),
    )
    conf_threshold: float = Field(
        default=0.3,
        description=(
            "Минимальная уверенность модели при первичной детекции региона "
            "(`post_process_object_detection`). Ниже → больше шумных регионов и таблиц, "
            "выше → можно пропустить слабые сигналы."
        ),
    )
    post_conf_min: float = Field(
        default=0.1,
        description=(
            "Порог пост-фильтрации: регионы со score ниже этого значения выбрасываются "
            "перед dedup-проходом."
        ),
    )
    dedup_iou: float = Field(
        default=0.9,
        description=(
            "Порог IoU для дедупликации перекрывающихся регионов. Если два региона имеют "
            "IoU > dedup_iou, оставляем только с бОльшим score."
        ),
    )
    containment_thr: float = Field(
        default=0.7,
        description=(
            "Порог containment: если меньший регион на >X% находится внутри большего, "
            "меньший считается дубликатом и удаляется. Борется с кейсом `Text` + `List-item` "
            "одновременно."
        ),
    )
# END_PDF_LAYOUT


# START_PDF_VLM
class PdfVlmConfig(BaseSettings):
    """Параметры VLM-клиента (PaddleOCR-VL на vLLM). Используется и pdf-, и docx-парсером.

    `base_url` и `api_key` берутся из переменных окружения OCR_BASE_URL / OCR_API_KEY (или `.env`).
    """

    model_config = SettingsConfigDict(**_ENV_SETTINGS)

    base_url: str = Field(
        default="http://127.0.0.1:11438/v1/",
        validation_alias=AliasChoices("OCR_BASE_URL"),
        description=(
            "URL OpenAI-совместимого VLM/OCR-сервиса (env: OCR_BASE_URL). Используется `VLMClient` "
            "в `src/format_parsers/pdf/_clients.py`."
        ),
    )
    model: Optional[str] = Field(
        default=None,
        description=(
            "Имя VLM-модели. `None` → автоопределение через `GET /v1/models` на старте "
            "клиента."
        ),
    )
    api_key: str = Field(
        default="none",
        validation_alias=AliasChoices("OCR_API_KEY"),
        description=(
            "API-ключ (env: OCR_API_KEY). Для vLLM обычно любое непустое значение (не проверяется). "
            "Передаётся в `AsyncOpenAI(api_key=...)`."
        ),
    )
    max_tokens: int = Field(
        default=2000,
        description=(
            "Лимит токенов на ответ VLM для PDF-страниц. Full-page OCR обычно укладывается "
            "в 1000-1500; запас 2000 покрывает плотные таблицы."
        ),
    )
    temperature: float = Field(
        default=0.0,
        description=(
            "Температура генерации VLM. Для OCR всегда 0 — нам нужен детерминизм, не "
            "креативность."
        ),
    )
# END_PDF_VLM


# START_PDF_PARSING
class PdfParsingConfig(BaseModel):
    """Параметры пост-обработки результатов VLM: промпты по классам, markdown-обёртки, формат таблиц."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    table_output_format: str = Field(
        default="html",
        description=(
            "`html` или `markdown` — формат вывода таблиц из PaddleOCR-VL. "
            "`html` сохраняет colspan/rowspan (полезно для LLM-аудита), "
            "`markdown` — плоский, теряет merged-cells."
        ),
    )
    class_prompts: Dict[str, str] = Field(
        default_factory=lambda: {
            "Text": "OCR:",
            "Title": "OCR:",
            "Section-header": "OCR:",
            "Caption": "OCR:",
            "Footnote": "OCR:",
            "List-item": "OCR:",
            "Page-header": "OCR:",
            "Page-footer": "OCR:",
            "Document Index": "OCR:",
            "Code": "OCR:",
            "Form": "OCR:",
            "Key-Value Region": "OCR:",
            "Checkbox-Selected": "OCR:",
            "Checkbox-Unselected": "OCR:",
            "Table": "Table Recognition:",
            "Formula": "Formula Recognition:",
            "Picture": "OCR:",
        },
        description=(
            "Карта `class_name → prompt` для VLM. Регионы, для класса которых нет записи "
            "в этой карте, пропускаются при сборке страницы. `Table` и `Formula` имеют "
            "специальные промпты, активирующие специфичное поведение модели."
        ),
    )
    class_md_wrapper: Dict[str, str] = Field(
        default_factory=lambda: {
            "Text": "{text}",
            "Title": "## {text}",
            "Section-header": "### {text}",
            "Caption": "*{text}*",
            "Footnote": "*{text}*",
            "List-item": "{text}",
            "Page-header": "{text}",
            "Page-footer": "{text}",
            "Document Index": "{text}",
            "Code": "```\n{text}\n```",
            "Form": "{text}",
            "Key-Value Region": "{text}",
            "Checkbox-Selected": "[x] {text}",
            "Checkbox-Unselected": "[ ] {text}",
            "Table": "{text}",
            "Formula": "$$\n{text}\n$$",
            "Picture": "{text}",
        },
        description=(
            "Шаблоны markdown-обёртки для каждого класса региона. Placeholder `{text}` "
            "заменяется распознанным текстом. Например, `Title` → `## Заголовок`, "
            "`Formula` → `$$\\nF=ma\\n$$`."
        ),
    )
# END_PDF_PARSING


# START_PDF_EXTRACTOR
class PdfExtractorConfig(BaseModel):
    """Параметры оркестратора PDF-парсинга: рендер страниц, фильтрация, padding."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    max_page_pixels: int = Field(
        default=4000,
        description=(
            "Максимальный размер длинной стороны отрендеренной страницы в пикселях. "
            "Если страница крупнее после scale по DPI → рендер уменьшается, чтобы не "
            "перегрузить layout-детектор и VLM."
        ),
    )
    blank_page_white_threshold: float = Field(
        default=0.995,
        description=(
            "Доля «белых» пикселей (все каналы ≥ 250), при которой страница считается "
            "пустой и пропускается. 0.995 = 99.5% белого."
        ),
    )
    blank_page_detect_dpi: int = Field(
        default=72,
        description=(
            "Низкий DPI для быстрого рендеринга при детекции пустых страниц. 72 достаточно, "
            "чтобы отличить пустоту от контента, но в 3-4 раза быстрее рабочего DPI."
        ),
    )
    small_element_min_area_ratio: float = Field(
        default=0.001,
        description=(
            "Минимальная площадь региона как доля от площади страницы. Применяется только "
            "к классам из `small_element_classes`. Отсекает артефакты типа 1-пиксельных "
            "Page-header."
        ),
    )
    small_element_classes: List[str] = Field(
        default_factory=lambda: ["Page-header", "Page-footer"],
        description=(
            "Имена классов, к которым применяется фильтр по минимальной площади "
            "(`small_element_min_area_ratio`). Для основных классов (Text, Table, ...) "
            "фильтр не применяется."
        ),
    )
    table_mask_padding_px: int = Field(
        default=5,
        description=(
            "Отступ в пикселях при маскировании таблиц белым перед full-page OCR. "
            "Нужен, чтобы целиком скрыть границы таблицы."
        ),
    )
    crop_padding_px: int = Field(
        default=15,
        description=(
            "Отступ в пикселях при вырезании региона под отдельный VLM-запрос. "
            "Запас против ошибок bbox на 1-2 пикселя."
        ),
    )
# END_PDF_EXTRACTOR


# START_PDF_PARSER_CONFIG
class PdfParserConfig(BaseModel):
    """Корневой конфиг PDF-парсера — объединяет layout, VLM, пост-обработку и оркестратор."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    layout: PdfLayoutConfig = Field(
        default_factory=PdfLayoutConfig,
        description="Параметры layout-детектора (remote/local).",
    )
    vlm: PdfVlmConfig = Field(
        default_factory=PdfVlmConfig,
        description="Параметры VLM-клиента (PaddleOCR-VL).",
    )
    parsing: PdfParsingConfig = Field(
        default_factory=PdfParsingConfig,
        description="Параметры пост-обработки (промпты, markdown-обёртки).",
    )
    extractor: PdfExtractorConfig = Field(
        default_factory=PdfExtractorConfig,
        description="Параметры оркестратора (рендер, фильтры).",
    )
    dpi: int = Field(
        default=200,
        description=(
            "DPI для рендеринга PDF-страниц в изображения. 200 — баланс качества/скорости. "
            "Выше → лучше качество OCR, медленнее и больше память."
        ),
    )
# END_PDF_PARSER_CONFIG


# START_DOCX_HEADER_OCR
class DocxHeaderOcrConfig(BaseModel):
    """Параметры OCR-распознавания картинок (эмблема, штампы) в шапке DOCX."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    prompt: str = Field(
        default="OCR this image. Return all text exactly as written.",
        description=(
            "Текстовый промпт для VLM при распознавании header-картинок DOCX. "
            "Отличается от PDF-промптов: там class-specific промпты (`OCR:`, "
            "`Table Recognition:`), здесь — универсальный «верни текст как есть»."
        ),
    )
    min_image_width: int = Field(
        default=200,
        description=(
            "Минимальная ширина картинки в пикселях, чтобы считаться значимой. "
            "Меньше → вероятно декоративная (маркер списка, мини-иконка) — пропускаем."
        ),
    )
    min_image_height: int = Field(
        default=50,
        description=(
            "Минимальная высота картинки в пикселях. См. `min_image_width`."
        ),
    )
    max_paragraphs_scanned: int = Field(
        default=6,
        description=(
            "Сколько первых абзацев документа осматривать на наличие embedded-картинок. "
            "Шапка обычно укладывается в первые 3-4 абзаца, 6 — запас."
        ),
    )
    max_tokens: int = Field(
        default=256,
        description=(
            "Лимит токенов VLM-ответа для header-OCR. Шапка — это название организации, "
            "эмблема, реквизиты; 256 токенов достаточно. Передаётся как override в "
            "`VLMClient.recognize_sync(..., max_tokens=...)`."
        ),
    )
# END_DOCX_HEADER_OCR


# START_PARSERS_CONFIG
class ParsersConfig(BaseModel):
    """Корневой конфиг всех парсеров проекта."""

    model_config = ConfigDict(frozen=True, extra="forbid")

    pdf: PdfParserConfig = Field(
        default_factory=PdfParserConfig,
        description="Конфиг PDF-парсера (PaddleOCR via vLLM + layout detector).",
    )
    docx_header_ocr: DocxHeaderOcrConfig = Field(
        default_factory=DocxHeaderOcrConfig,
        description="Конфиг header-OCR для DOCX-парсера.",
    )
# END_PARSERS_CONFIG


# START_SINGLETON
# PURPOSE: Единственный экземпляр конфига на весь рантайм. Используется импортом `from config.parsers import PARSERS_CONFIG`.
# INPUTS: —
# OUTPUTS: `PARSERS_CONFIG` с полностью заполненными дефолтами.
# KEYWORDS: singleton, config, runtime.
PARSERS_CONFIG = ParsersConfig()
# END_SINGLETON
