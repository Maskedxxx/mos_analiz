# START_MODULE_CONTRACT
# PURPOSE: Конфиг PDF-парсера — инфраструктурные и алгоритмические параметры (URLs, модели, DPI, пороги). Все значения приходят из `config/parsers.json`; никаких литералов в Python.
# INPUTS: Путь к `config/parsers.json` (опционально — по умолчанию в корне проекта).
# OUTPUTS: Типизированные `PdfParserConfig` / `PdfLayoutConfig` / `PdfVlmConfig`.
# KEYWORDS: config, pdf-parser, infra, no-hardcode, single-source-of-truth, nested-config.
# LINKS: config/parsers.json, src/format_parsers/pdf/_clients.py, src/format_parsers/pdf/__init__.py.
# RATIONALE: Всё, что относится к PDF-парсингу (URLs сервисов, модель layout, пороги IoU, параметры VLM) — единая точка правды в JSON. Nested структура отражает клиентский слой (layout, vlm) и общий слой (dpi).
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional
# END_IMPORTS


# START_SUBCONFIGS
# PURPOSE: Вложенные конфиги на клиентский слой (layout, vlm). Каждый соответствует отдельному внешнему сервису/модели.
# INPUTS: —
# OUTPUTS: Dataclass-объекты для передачи в соответствующие клиенты `_clients.py`.
# KEYWORDS: subconfig, layout, vlm, per-client.
@dataclass
class PdfLayoutConfig:
    """
    Назначение:
        Контракт параметров layout-детектора. Одна и та же структура используется как
        remote (HTTP API), так и local (локальная Docling Heron-101) детекторами —
        оркестратор выбирает реализацию по наличию `base_url`.

    Поля:
        base_url: URL remote HTTP Layout API. `null` → используется локальный детектор.
        timeout: Таймаут для remote-запросов (сек). Игнорируется локальным детектором.
        model: HF repo_id для локального детектора (например, docling-project/docling-layout-heron-101).
            Игнорируется remote-детектором.
        device: Устройство для локального детектора (cuda:0 / cpu). Игнорируется remote.
        conf_threshold: Минимальный score от детектора при первичной выборке регионов.
        post_conf_min: Порог пост-фильтрации (регионы со score ниже отбрасываются).
        dedup_iou: Порог IoU для дедупликации регионов.
        containment_thr: Порог containment для удаления вложенных регионов.
    """

    base_url: Optional[str]
    timeout: int
    model: str
    device: str
    conf_threshold: float
    post_conf_min: float
    dedup_iou: float
    containment_thr: float


@dataclass
class PdfVlmConfig:
    """
    Назначение:
        Контракт параметров VLM-клиента (PaddleOCR-VL через vLLM-совместимый сервис).

    Поля:
        base_url: URL VLM-сервиса.
        model: Имя модели. `null` → автоопределение через `/v1/models`.
        api_key: API-ключ (для vLLM обычно любое непустое значение).
        max_tokens: Максимум токенов в ответе VLM.
        temperature: Температура генерации VLM.
    """

    base_url: str
    model: Optional[str]
    api_key: str
    max_tokens: int
    temperature: float


@dataclass
class PdfParsingConfig:
    """
    Назначение:
        Контракт параметров пост-обработки результатов VLM: какие классы регионов
        получают какие промпты, какими markdown-обёртками оборачиваются и в каком
        формате сериализуются таблицы.

    Поля:
        table_output_format: `html` или `markdown` — формат вывода таблиц из PaddleOCR-VL.
        class_prompts: Карта `class_name → prompt` для VLM. Регионы, для которых нет
            записи в этой карте, пропускаются при сборке страницы.
        class_md_wrapper: Карта `class_name → формат-строка`. В формат-строке должен быть
            placeholder `{text}`, который заменяется распознанным текстом региона.

    Логика:
        Используется на этапе сборки страницы в markdown (`build_page_markdown`) и
        при выборе промпта в экстракторе. Карты живут в конфиге, чтобы при смене
        номенклатуры классов или стиля вывода не трогать Python.
    """

    table_output_format: str
    class_prompts: Dict[str, str]
    class_md_wrapper: Dict[str, str]


@dataclass
class PdfExtractorConfig:
    """
    Назначение:
        Контракт параметров PDF-экстрактора (оркестратора: рендеринг → layout → VLM →
        сборка страницы).

    Поля:
        max_page_pixels: Макс. размер длинной стороны отрендеренной страницы в пикселях.
            Если страница крупнее — рендер уменьшается, чтобы не перегружать layout/VLM.
        blank_page_white_threshold: Доля «белых» пикселей (≥ 250 по каждому каналу), при
            которой страница считается пустой и пропускается.
        blank_page_detect_dpi: Низкий DPI для быстрого рендеринга при детекции пустых страниц.
        small_element_min_area_ratio: Минимальная площадь региона как доля от площади
            страницы. Меньшие — выкидываются (для классов из `small_element_classes`).
        small_element_classes: Имена классов, к которым применяется фильтр по площади.
        table_mask_padding_px: Отступ при маскировании таблиц белым (чтобы full-page OCR
            не читал таблицу).
        crop_padding_px: Отступ при вырезании региона под отдельный VLM-запрос.
    """

    max_page_pixels: int
    blank_page_white_threshold: float
    blank_page_detect_dpi: int
    small_element_min_area_ratio: float
    small_element_classes: List[str]
    table_mask_padding_px: int
    crop_padding_px: int
# END_SUBCONFIGS


# START_PDF_PARSER_CONFIG
# PURPOSE: Корневой конфиг PDF-парсера. Содержит все суб-конфиги плюс общие параметры.
# INPUTS: —
# OUTPUTS: Dataclass для передачи в `parse_pdf`.
# KEYWORDS: dataclass, pdf-config, required-fields, nested.
@dataclass
class PdfParserConfig:
    """
    Назначение:
        Единый контракт инфраструктурных параметров PDF-парсера.

    Поля:
        layout: Параметры layout-детектора.
        vlm: Параметры VLM-клиента.
        parsing: Параметры пост-обработки (промпты, markdown-обёртки, формат таблиц).
        extractor: Параметры оркестратора (рендер, пороги фильтрации, padding).
        dpi: DPI для рендеринга PDF-страниц в изображения.

    Логика:
        Все значения обязательны и приходят из `config/parsers.json`. Отсутствие файла
        или любого поля — явная ошибка при загрузке (`FileNotFoundError` / `ValueError`).
    """

    layout: PdfLayoutConfig
    vlm: PdfVlmConfig
    parsing: PdfParsingConfig
    extractor: PdfExtractorConfig
    dpi: int
# END_PDF_PARSER_CONFIG


# START_CONFIG_LOADER
# PURPOSE: Загрузчик `config/parsers.json` → `PdfParserConfig`. Без скрытых дефолтов: отсутствие файла или поля — явная ошибка.
# INPUTS: Опциональный путь к JSON-файлу. Если не указан — ищет в `<project_root>/config/parsers.json`.
# OUTPUTS: Экземпляр `PdfParserConfig`.
# KEYWORDS: loader, strict-validation, no-silent-defaults, nested-parse.
_LAYOUT_REQUIRED_FIELDS = (
    "base_url",
    "timeout",
    "model",
    "device",
    "conf_threshold",
    "post_conf_min",
    "dedup_iou",
    "containment_thr",
)
_VLM_REQUIRED_FIELDS = (
    "base_url",
    "model",
    "api_key",
    "max_tokens",
    "temperature",
)
_PARSING_REQUIRED_FIELDS = (
    "table_output_format",
    "class_prompts",
    "class_md_wrapper",
)
_EXTRACTOR_REQUIRED_FIELDS = (
    "max_page_pixels",
    "blank_page_white_threshold",
    "blank_page_detect_dpi",
    "small_element_min_area_ratio",
    "small_element_classes",
    "table_mask_padding_px",
    "crop_padding_px",
)
_PDF_TOP_REQUIRED_FIELDS = ("layout", "vlm", "parsing", "extractor", "dpi")


def _default_config_path() -> Path:
    """
    Назначение:
        Возвращает путь к `config/parsers.json` по умолчанию — от корня проекта.

    Вход:
        —

    Выход:
        Path: Путь `<project_root>/config/parsers.json`.

    Логика:
        1. Берёт директорию текущего файла (`src/format_parsers/pdf/_config.py`).
        2. Поднимается на 3 уровня: `pdf/` → `format_parsers/` → `src/` → корень проекта.
        3. Возвращает `<корень>/config/parsers.json`.
    """
    return Path(__file__).resolve().parents[3] / "config" / "parsers.json"


def _require_fields(section: dict, fields: tuple, ctx: str) -> None:
    """
    Назначение:
        Проверяет, что все `fields` присутствуют в `section`. Бросает `ValueError`
        со списком отсутствующих ключей и контекстом, где именно провалилось.

    Вход:
        section: Словарь-секция из JSON.
        fields: Обязательные имена ключей.
        ctx: Человекочитаемый путь до секции (для сообщения об ошибке).

    Выход:
        None. При отсутствии полей — `ValueError`.

    Логика:
        1. Собирает список ключей, которых нет в секции.
        2. Если пусто — возвращает управление.
        3. Иначе — бросает `ValueError` с перечнем.
    """
    missing = [k for k in fields if k not in section]
    if missing:
        raise ValueError(f"{ctx}: не хватает полей {missing}.")


def load_pdf_parser_config(path: Optional[Path] = None) -> PdfParserConfig:
    """
    Назначение:
        Загружает инфраструктурный конфиг PDF-парсера из `config/parsers.json`.

    Вход:
        path: Путь к JSON-файлу. Если `None` — дефолт в корне проекта.

    Выход:
        PdfParserConfig: Экземпляр с полностью заполненными суб-конфигами.

    Логика:
        1. Если файл не существует — `FileNotFoundError`.
        2. Если корневая секция `pdf` отсутствует/не dict — `ValueError`.
        3. Если в `pdf` не хватает `layout`/`vlm`/`dpi` — `ValueError` с перечнем.
        4. Для каждой подсекции (`layout`, `vlm`) проверяет обязательные поля.
        5. Возвращает `PdfParserConfig(PdfLayoutConfig(...), PdfVlmConfig(...), dpi)`.
    """
    cfg_path = path or _default_config_path()
    if not cfg_path.exists():
        raise FileNotFoundError(f"Конфиг парсеров не найден: {cfg_path}")
    with open(cfg_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    pdf_section: Any = data.get("pdf") if isinstance(data, dict) else None
    if not isinstance(pdf_section, dict):
        raise ValueError(f"В {cfg_path} отсутствует или некорректна секция 'pdf'.")
    _require_fields(pdf_section, _PDF_TOP_REQUIRED_FIELDS, f"{cfg_path}, секция 'pdf'")

    layout_section = pdf_section["layout"]
    if not isinstance(layout_section, dict):
        raise ValueError(f"{cfg_path}, 'pdf.layout': ожидался объект.")
    _require_fields(layout_section, _LAYOUT_REQUIRED_FIELDS, f"{cfg_path}, секция 'pdf.layout'")

    vlm_section = pdf_section["vlm"]
    if not isinstance(vlm_section, dict):
        raise ValueError(f"{cfg_path}, 'pdf.vlm': ожидался объект.")
    _require_fields(vlm_section, _VLM_REQUIRED_FIELDS, f"{cfg_path}, секция 'pdf.vlm'")

    parsing_section = pdf_section["parsing"]
    if not isinstance(parsing_section, dict):
        raise ValueError(f"{cfg_path}, 'pdf.parsing': ожидался объект.")
    _require_fields(parsing_section, _PARSING_REQUIRED_FIELDS, f"{cfg_path}, секция 'pdf.parsing'")

    extractor_section = pdf_section["extractor"]
    if not isinstance(extractor_section, dict):
        raise ValueError(f"{cfg_path}, 'pdf.extractor': ожидался объект.")
    _require_fields(extractor_section, _EXTRACTOR_REQUIRED_FIELDS, f"{cfg_path}, секция 'pdf.extractor'")

    # Значения берём из JSON как есть — никаких подмен и дефолтов здесь нет.
    layout = PdfLayoutConfig(
        base_url=layout_section["base_url"],
        timeout=layout_section["timeout"],
        model=layout_section["model"],
        device=layout_section["device"],
        conf_threshold=layout_section["conf_threshold"],
        post_conf_min=layout_section["post_conf_min"],
        dedup_iou=layout_section["dedup_iou"],
        containment_thr=layout_section["containment_thr"],
    )
    vlm = PdfVlmConfig(
        base_url=vlm_section["base_url"],
        model=vlm_section["model"],
        api_key=vlm_section["api_key"],
        max_tokens=vlm_section["max_tokens"],
        temperature=vlm_section["temperature"],
    )
    parsing = PdfParsingConfig(
        table_output_format=parsing_section["table_output_format"],
        class_prompts=parsing_section["class_prompts"],
        class_md_wrapper=parsing_section["class_md_wrapper"],
    )
    extractor = PdfExtractorConfig(
        max_page_pixels=extractor_section["max_page_pixels"],
        blank_page_white_threshold=extractor_section["blank_page_white_threshold"],
        blank_page_detect_dpi=extractor_section["blank_page_detect_dpi"],
        small_element_min_area_ratio=extractor_section["small_element_min_area_ratio"],
        small_element_classes=list(extractor_section["small_element_classes"]),
        table_mask_padding_px=extractor_section["table_mask_padding_px"],
        crop_padding_px=extractor_section["crop_padding_px"],
    )
    return PdfParserConfig(
        layout=layout,
        vlm=vlm,
        parsing=parsing,
        extractor=extractor,
        dpi=pdf_section["dpi"],
    )
# END_CONFIG_LOADER
