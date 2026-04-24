# START_MODULE_CONTRACT
# PURPOSE: Оркестратор PDF-парсинга и публичный entry `parse_pdf`. Связывает клиенты (IO) и чистую логику (парсинг) в один конвейер: открыть PDF → детектировать пустые страницы → layout + VLM на каждой странице → склеить всё в `raw_text`.
# INPUTS: Путь к PDF-файлу, `PdfParserConfig` (из `config/parsers.json`), опциональная директория для debug-логов.
# OUTPUTS: `ParsedDocument` с полями `filename`, `path`, `raw_text`. Полный текст всех страниц, разделитель — `[СТРАНИЦА N]`.
# KEYWORDS: pdf, orchestrator, paddle-extractor, parse-pdf, raw-text, public-api.
# LINKS: src/format_parsers/pdf/_config.py, src/format_parsers/pdf/_clients.py, src/format_parsers/pdf/_parsing.py, src/format_parsers/_types.py.
# RATIONALE: Публичный API пакета — одна функция `parse_pdf(file_path, config)`. Всё остальное в пакете — внутренности, скрытые от потребителя.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import asyncio
import json
import logging
import time
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import fitz
import numpy as np
from PIL import Image, ImageDraw

from src.format_parsers._types import ParsedDocument
from src.format_parsers.pdf._clients import (
    LayoutDetector,
    RemoteLayoutDetector,
    VLMClient,
)
from config.parsers import PdfParserConfig
from src.format_parsers.pdf._parsing import (
    is_table_format,
    merge_text_and_tables,
    parse_paddle_table_html,
)
# END_IMPORTS


# START_LOGGER
extractor_logger = logging.getLogger(__name__ + ".extractor")
parse_pdf_logger = logging.getLogger(__name__ + ".parse_pdf")
# END_LOGGER


# START_PADDLE_EXTRACTOR
# PURPOSE: Оркестратор low-level PDF pipeline: рендер страниц через PyMuPDF, layout-детекция (local/remote), VLM-OCR, сборка HTML-таблиц, возврат page_num → markdown.
# INPUTS: `PdfParserConfig` (все параметры), опциональная директория для debug-артефактов.
# OUTPUTS: `Dict[int, str]` — 1-based номер страницы → markdown страницы (с вставленными HTML-таблицами).
# KEYWORDS: extractor, paddle, layout, vlm, pymupdf, blank-pages.
class PaddleExtractor:
    """
    Назначение:
        Внутренний оркестратор Paddle-пайплайна. Используется `parse_pdf` — снаружи пакета
        не предназначен для прямого вызова.

    Вход:
        config: `PdfParserConfig` — все параметры из `config/parsers.json`.
        log_dir: Директория для debug-артефактов (визуализации layout, crop'ы таблиц,
            сырые JSON). `None` → логов-артефактов нет.

    Выход:
        Объект с методами `detect_blank_pages(pdf_path)` и `extract_pages(pdf_path, indices)`.

    Логика:
        Модели загружаются лениво при первом вызове `extract_pages()` — это даёт быструю
        детекцию пустых страниц без прогрева layout/VLM.
    """

    def __init__(self, config: PdfParserConfig, log_dir: Optional[Path] = None):
        """
        Назначение:
            Сохраняет конфиг и пути для будущих вызовов. Модели пока не загружает.

        Вход:
            config: `PdfParserConfig`.
            log_dir: Директория для debug-артефактов или `None`.

        Выход:
            None.

        Логика:
            Lazy init — клиенты создаются при первом `_ensure_models()`.
        """
        self.config = config
        self.log_dir = log_dir
        self._detector: Optional[Any] = None
        self._vlm: Optional[VLMClient] = None

    def _ensure_models(self) -> None:
        """
        Назначение:
            Лениво создаёт layout-детектор (remote или local) и VLM-клиент.

        Вход:
            —

        Выход:
            None.

        Логика:
            1. Если `layout.base_url` задан — remote HTTP детектор, иначе локальная Heron-101.
            2. VLM-клиент с параметрами из `vlm`-секции.
            3. Повторные вызовы — no-op.
        """
        if self._detector is None:
            if self.config.layout.base_url:
                extractor_logger.info(f"Использую RemoteLayoutDetector: {self.config.layout.base_url}")
                self._detector = RemoteLayoutDetector(
                    base_url=self.config.layout.base_url,
                    timeout=self.config.layout.timeout,
                )
            else:
                extractor_logger.info("Загрузка локального LayoutDetector (Heron-101)...")
                self._detector = LayoutDetector(
                    model_repo=self.config.layout.model,
                    device=self.config.layout.device,
                    conf_threshold=self.config.layout.conf_threshold,
                    post_conf_min=self.config.layout.post_conf_min,
                    dedup_iou=self.config.layout.dedup_iou,
                    containment_thr=self.config.layout.containment_thr,
                )
        if self._vlm is None:
            extractor_logger.info("Инициализация VLM-клиента...")
            self._vlm = VLMClient(
                base_url=self.config.vlm.base_url,
                model=self.config.vlm.model,
                api_key=self.config.vlm.api_key,
                max_tokens=self.config.vlm.max_tokens,
                temperature=self.config.vlm.temperature,
            )

    def _is_blank_page(self, image: Image.Image) -> bool:
        """
        Назначение:
            Проверяет, «пустая» ли страница по доле белых пикселей.

        Вход:
            image: PIL Image страницы.

        Выход:
            True если страница считается пустой.

        Логика:
            Считает долю пикселей со значением ≥ 250 по всем каналам. Порог — из конфига.
        """
        arr = np.array(image)
        white_ratio = (arr >= 250).sum() / arr.size
        return white_ratio >= self.config.extractor.blank_page_white_threshold

    def _render_page(self, doc: fitz.Document, page_idx: int, target_dpi: int) -> Image.Image:
        """
        Назначение:
            Рендерит одну страницу PDF в PIL Image с ограничением длинной стороны.

        Вход:
            doc: Открытый `fitz.Document`.
            page_idx: 0-based индекс страницы.
            target_dpi: Желаемый DPI.

        Выход:
            PIL Image RGB.

        Логика:
            1. Считает scale как `target_dpi / 72`.
            2. Если получившийся long_side превышает `max_page_pixels` — уменьшает scale,
               чтобы не рендерить слишком большие изображения.
            3. Вызывает PyMuPDF и конвертирует в PIL.
        """
        page = doc[page_idx]
        scale = target_dpi / 72.0
        long_side = max(page.rect.width, page.rect.height) * scale
        if long_side > self.config.extractor.max_page_pixels:
            scale = self.config.extractor.max_page_pixels / max(page.rect.width, page.rect.height)
        mat = fitz.Matrix(scale, scale)
        pix = page.get_pixmap(matrix=mat)
        return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    def detect_blank_pages(self, pdf_path: str) -> Tuple[int, Set[int]]:
        """
        Назначение:
            Определяет пустые страницы в PDF (быстрый low-DPI рендер + проверка на белое).

        Вход:
            pdf_path: Путь к PDF.

        Выход:
            `(total_pages, blank_indices)` — общее число страниц и множество 0-based индексов пустых.

        Логика:
            Проходит по всем страницам, рендерит каждую на `blank_page_detect_dpi`,
            проверяет `_is_blank_page`.
        """
        doc = fitz.open(pdf_path)
        total = len(doc)
        blanks: Set[int] = set()
        for i in range(total):
            img = self._render_page(doc, i, target_dpi=self.config.extractor.blank_page_detect_dpi)
            if self._is_blank_page(img):
                blanks.add(i)
                extractor_logger.info(f"  стр.{i + 1}: пустая (пропускаем)")
        if blanks:
            extractor_logger.info(f"Пустых страниц: {len(blanks)} из {total}")
        return (total, blanks)

    def extract_pages(self, pdf_path: str, page_indices: List[int]) -> Dict[int, str]:
        """
        Назначение:
            Layout-aware OCR для указанных страниц PDF.

        Вход:
            pdf_path: Путь к PDF.
            page_indices: Список 0-based индексов страниц для OCR.

        Выход:
            `{1-based page_num: markdown-текст страницы}`.

        Логика:
            1. Проверяет, что файл существует.
            2. Лениво инициализирует модели.
            3. Открывает PDF через PyMuPDF.
            4. Фильтрует невалидные индексы.
            5. Запускает async-конвейер обработки страниц.
        """
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF не найден: {pdf_path}")
        self._ensure_models()
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        if total_pages == 0:
            raise ValueError("PDF не содержит страниц")
        unique_indices = sorted(set(i for i in page_indices if 0 <= i < total_pages))
        if not unique_indices:
            extractor_logger.warning(f"Нет валидных страниц из {page_indices} (всего {total_pages})")
            return {}
        extractor_logger.info(f"Paddle OCR: {len(unique_indices)} страниц из {total_pages}")
        return asyncio.run(self._process_pages_from_pdf(doc, unique_indices))

    async def _process_pages_from_pdf(self, doc: fitz.Document, page_indices: List[int]) -> Dict[int, str]:
        """
        Назначение:
            Последовательно проходит страницы PDF: рендер → обработка.

        Вход:
            doc: Открытый PDF.
            page_indices: 0-based индексы страниц.

        Выход:
            `{1-based page_num: markdown}`.

        Логика:
            Для каждой страницы — рендер на рабочем DPI, затем `_process_single_page`.
            Ошибки не падают весь пайплайн — записываются как текст `[ОШИБКА Paddle OCR: ...]`.
        """
        page_texts: Dict[int, str] = {}
        for idx in page_indices:
            page_num = idx + 1
            try:
                img = self._render_page(doc, idx, target_dpi=self.config.dpi)
                md = await self._process_single_page(img, page_num)
                page_texts[page_num] = md
            except Exception as e:
                extractor_logger.error(f"Ошибка Paddle OCR стр.{page_num}: {e}")
                page_texts[page_num] = f"[ОШИБКА Paddle OCR: {e}]"
        return page_texts

    async def _process_single_page(self, page_image: Image.Image, page_num: int) -> str:
        """
        Назначение:
            Обрабатывает одну страницу: layout-детекция → full-page OCR + отдельные
            table-crop'ы → сборка.

        Вход:
            page_image: Отрендеренное изображение страницы.
            page_num: 1-based номер страницы (для логов).

        Выход:
            Текст страницы с вставленными HTML-таблицами.

        Логика:
            1. Layout-детекция.
            2. Фильтрация мелких регионов (Page-header/Page-footer ниже min_area).
            3. Визуализация layout в `log_dir` (если задан).
            4. Маскирование таблиц белым + full-page OCR на замаскированном изображении.
            5. Table crop'ы распознаются отдельно с промптом "Table Recognition:".
            6. VLM-запросы запускаются батчом параллельно.
            7. HTML-таблицы вставляются обратно в текст по y-позиции.
        """
        extractor_logger.info(f"=== Страница {page_num} ===")
        assert self._detector is not None and self._vlm is not None

        t0 = time.time()
        regions = self._detector.detect(page_image)
        layout_time = time.time() - t0
        class_counts = Counter(r["class_name"] for r in regions)
        extractor_logger.info(f"Layout: {len(regions)} регионов за {layout_time:.2f}s | {dict(class_counts)}")

        # Фильтр «слишком мелкие header/footer» — они чаще артефакты, чем полезный контент.
        page_w, page_h = page_image.size
        page_area = page_w * page_h
        min_area = page_area * self.config.extractor.small_element_min_area_ratio
        small_classes = set(self.config.extractor.small_element_classes)
        filtered: List[Dict[str, Any]] = []
        for r in regions:
            bbox = r["bbox"]
            box_area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
            if r["class_name"] in small_classes and box_area < min_area:
                extractor_logger.debug(f"  Пропуск мелкого {r['class_name']}: area={box_area:.0f} < {min_area:.0f}")
                continue
            filtered.append(r)
        if len(filtered) != len(regions):
            extractor_logger.info(f"Отфильтровано мелких: {len(regions) - len(filtered)}")
        regions = filtered

        if self.log_dir:
            self.log_dir.mkdir(parents=True, exist_ok=True)
            vis_path = self.log_dir / f"page_{page_num}_layout.png"
            self._detector.visualize(page_image, regions, str(vis_path))

        # Таблицы сортируем по y — порядок вставки в текст важен.
        table_regions = [r for r in regions if r["class_name"] == "Table"]
        table_regions.sort(key=lambda r: r["bbox"][1])
        n_tables = len(table_regions)
        extractor_logger.info(f"Таблиц: {n_tables}, full-page OCR + {n_tables} table crop(s)")

        # Формируем batch VLM-запросов: index=-1 для full-page OCR, 0..N для таблиц.
        vlm_items: List[Dict[str, Any]] = []
        if table_regions:
            ocr_page = self._mask_table_regions(page_image, table_regions)
        else:
            ocr_page = page_image
        vlm_items.append({"image": ocr_page, "prompt": "OCR:", "index": -1})
        for i, table_reg in enumerate(table_regions):
            crop = self._crop_region(page_image, table_reg["bbox"])
            vlm_items.append({"image": crop, "prompt": "Table Recognition:", "index": i})

        if self.log_dir:
            if table_regions:
                masked_path = self.log_dir / f"page_{page_num}_masked.png"
                ocr_page.save(str(masked_path))
            for i in range(n_tables):
                crop_path = self.log_dir / f"page_{page_num}_table_{i}.png"
                vlm_items[1 + i]["image"].save(str(crop_path))

        t0 = time.time()
        vlm_results = await self._vlm.recognize_batch(vlm_items)
        vlm_time = time.time() - t0
        extractor_logger.info(f"VLM: {len(vlm_results)} запросов за {vlm_time:.2f}s")

        vlm_map = {r["index"]: r["text"] for r in vlm_results}
        full_page_text = vlm_map.get(-1, "")
        table_entries: List[Dict[str, Any]] = []
        for i, table_reg in enumerate(table_regions):
            raw_table = vlm_map.get(i, "")
            html = parse_paddle_table_html(raw_table) if is_table_format(raw_table) else raw_table
            table_entries.append({"bbox": table_reg["bbox"], "html": html})

        if self.log_dir:
            raw_data = {
                "full_page_text": full_page_text,
                "tables": [
                    {
                        "bbox": [int(c) for c in table_regions[i]["bbox"]],
                        "score": round(table_regions[i]["score"], 3),
                        "raw_text": vlm_map.get(i, ""),
                        "parsed_html": entry["html"],
                    }
                    for i, entry in enumerate(table_entries)
                ],
            }
            raw_path = self.log_dir / f"page_{page_num}_raw.json"
            with open(raw_path, "w", encoding="utf-8") as f:
                json.dump(raw_data, f, ensure_ascii=False, indent=2)

        return merge_text_and_tables(full_page_text, table_entries, page_image.size)

    def _mask_table_regions(self, page_image: Image.Image, table_regions: List[Dict[str, Any]]) -> Image.Image:
        """
        Назначение:
            Закрашивает области таблиц белым на копии страницы. Full-page OCR не читает
            их содержимое — таблицы распознаются отдельно.

        Вход:
            page_image: Оригинал страницы.
            table_regions: Регионы с `class_name="Table"`.

        Выход:
            Копия страницы с замаскированными таблицами.

        Логика:
            Отступ маски — из `extractor.table_mask_padding_px`. Используется PIL ImageDraw.
        """
        padding = self.config.extractor.table_mask_padding_px
        masked = page_image.copy()
        draw = ImageDraw.Draw(masked)
        for region in table_regions:
            x1, y1, x2, y2 = [int(c) for c in region["bbox"]]
            x1 = max(0, x1 - padding)
            y1 = max(0, y1 - padding)
            x2 = min(masked.width, x2 + padding)
            y2 = min(masked.height, y2 + padding)
            draw.rectangle([x1, y1, x2, y2], fill="white")
        return masked

    def _crop_region(self, page_image: Image.Image, bbox: List[float]) -> Image.Image:
        """
        Назначение:
            Вырезает регион из изображения страницы с паддингом из конфига.

        Вход:
            page_image: Изображение страницы.
            bbox: `[x1, y1, x2, y2]`.

        Выход:
            Вырезанное изображение.

        Логика:
            Паддинг — из `extractor.crop_padding_px`, клампится по границам страницы.
        """
        padding = self.config.extractor.crop_padding_px
        x1, y1, x2, y2 = [int(c) for c in bbox]
        w, h = page_image.size
        x1c = max(0, x1 - padding)
        y1c = max(0, y1 - padding)
        x2c = min(w, x2 + padding)
        y2c = min(h, y2 + padding)
        return page_image.crop((x1c, y1c, x2c, y2c))
# END_PADDLE_EXTRACTOR


# START_PARSE_PDF
# PURPOSE: Публичный entry пакета. Принимает путь к PDF и конфиг, отдаёт ParsedDocument.
# INPUTS: Путь к PDF, `PdfParserConfig`, опциональная директория для debug-артефактов.
# OUTPUTS: `ParsedDocument{filename, path, raw_text}`. Все страницы склеены в один raw_text с маркерами `[СТРАНИЦА N]`.
# KEYWORDS: parse-pdf, public-api, raw-text, full-document.
def parse_pdf(
    file_path: str,
    config: PdfParserConfig,
    log_dir: Optional[Path] = None,
) -> ParsedDocument:
    """
    Назначение:
        Читает PDF целиком и возвращает полный текст документа без разбивки на чанки.

    Вход:
        file_path: Путь к PDF-файлу.
        config: `PdfParserConfig` (из `config/parsers.json`).
        log_dir: Директория для debug-артефактов (layout-визуализации, crop'ы, raw JSON).

    Выход:
        `ParsedDocument{filename, path, raw_text}`. Полный текст всех непустых страниц,
        разделитель — `[СТРАНИЦА N]` (N — 1-based номер).

    Логика:
        1. Открывает PDF, детектирует пустые страницы на низком DPI.
        2. Для всех непустых — layout + VLM OCR через `PaddleExtractor`.
        3. Склеивает `page_texts` в один `raw_text` с сохранением порядка и маркерами.
        4. Возвращает унифицированный `ParsedDocument`.
    """
    file_path_obj = Path(file_path)
    if not file_path_obj.exists():
        raise FileNotFoundError(f"PDF не найден: {file_path}")

    parse_pdf_logger.info(f"PDF: {file_path_obj.name}")
    extractor = PaddleExtractor(config=config, log_dir=log_dir)

    # Шаг 1: быстрое определение пустых страниц на низком DPI. Обрабатываем только непустые.
    total_pages, blank_pages = extractor.detect_blank_pages(str(file_path_obj))
    non_blank_indices = [i for i in range(total_pages) if i not in blank_pages]
    parse_pdf_logger.info(
        f"  Страниц всего: {total_pages}, непустых: {len(non_blank_indices)} (пустых: {len(blank_pages)})"
    )

    # Шаг 2: layout-aware OCR для всех непустых страниц.
    page_texts = extractor.extract_pages(str(file_path_obj), non_blank_indices) if non_blank_indices else {}

    # Шаг 3: склеиваем страницы в raw_text. Порядок — по возрастанию номера.
    parts: List[str] = []
    for page_num in sorted(page_texts.keys()):
        text = page_texts[page_num].strip()
        if not text:
            continue
        # Маркер `[СТРАНИЦА N]` сохраняет порядок и визуальные границы в сыром тексте.
        parts.append(f"[СТРАНИЦА {page_num}]\n{text}")
    raw_text = "\n\n".join(parts)

    parse_pdf_logger.info(f"  raw_text: {len(raw_text)} символов из {len(parts)} страниц")

    return {
        "filename": file_path_obj.name,
        "path": str(file_path_obj.resolve()),
        "raw_text": raw_text,
    }
# END_PARSE_PDF
