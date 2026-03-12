#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Paddle Extractor — layout-aware постраничное OCR через Heron-101 + PaddleOCR-VL-1.5.

Для каждой страницы:
1. Layout detection (Docling Heron-101) → регионы
2. Фильтрация мелких элементов
3. Reading order (XY-cut)
4. Crop каждого региона с padding
5. VLM OCR (PaddleOCR-VL-1.5) — параллельно через asyncio
6. Сборка в markdown (build_page_markdown)

Результат: Dict[int, str] — {1-based page_num: markdown_text}.
"""

import asyncio
import json
import logging
import time
from collections import Counter
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import numpy as np
from PIL import Image
from pdf2image import convert_from_path

from . import config
from .layout_detector import LayoutDetector
from .vlm_client import VLMClient
from .reading_order import sort_by_reading_order
from .normalizer import merge_text_and_tables
from .table_parser import parse_paddle_table_html, is_table_format

logger = logging.getLogger(__name__)


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

    def __init__(
        self,
        vlm_base_url: str = "http://localhost:8010/v1",
        vlm_model: Optional[str] = None,
        layout_model_repo: Optional[str] = None,
        layout_device: str = "cuda:0",
        dpi: int = 200,
        log_dir: Optional[str] = None,
    ):
        self.vlm_base_url = vlm_base_url
        self.vlm_model = vlm_model
        self.layout_model_repo = layout_model_repo
        self.layout_device = layout_device
        self.dpi = dpi
        self.log_dir = Path(log_dir) if log_dir else None

        # Lazy init — модели загружаются при первом extract_pages()
        self._detector: Optional[LayoutDetector] = None
        self._vlm: Optional[VLMClient] = None

    def _ensure_models(self) -> None:
        """Загружает модели если ещё не загружены."""
        if self._detector is None:
            logger.info("Загрузка Heron-101 layout detector...")
            self._detector = LayoutDetector(
                model_name=self.layout_model_repo,
                device=self.layout_device,
            )
        if self._vlm is None:
            logger.info("Инициализация VLM-клиента...")
            self._vlm = VLMClient(
                base_url=self.vlm_base_url,
                model_name=self.vlm_model,
            )

    @staticmethod
    def _is_blank_page(image: Image.Image, threshold: float = 0.995) -> bool:
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

    def detect_blank_pages(self, pdf_path: str) -> Tuple[int, Set[int]]:
        """
        Определяет пустые страницы в PDF (быстрый рендер 72 DPI).

        Args:
            pdf_path: путь к PDF

        Returns:
            (total_pages, blank_indices): общее число страниц и множество
            0-based индексов пустых страниц
        """
        images = convert_from_path(pdf_path, dpi=72, fmt='PNG')
        total = len(images)
        blanks = set()
        for i, img in enumerate(images):
            if self._is_blank_page(img):
                blanks.add(i)
                logger.info(f"  стр.{i+1}: пустая (пропускаем)")

        if blanks:
            logger.info(f"Пустых страниц: {len(blanks)} из {total}")

        return total, blanks

    def extract_pages(
        self,
        pdf_path: str,
        page_indices: List[int]
    ) -> Dict[int, str]:
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
            raise FileNotFoundError(f"PDF не найден: {pdf_path}")

        self._ensure_models()

        # PDF → PNG[]
        images = convert_from_path(pdf_path, dpi=self.dpi, fmt='PNG')
        total_pages = len(images)

        if total_pages == 0:
            raise ValueError("PDF не содержит страниц")

        # Фильтруем валидные индексы
        unique_indices = sorted(set(
            i for i in page_indices if 0 <= i < total_pages
        ))

        if not unique_indices:
            logger.warning(f"Нет валидных страниц из {page_indices} (всего {total_pages})")
            return {}

        logger.info(f"Paddle OCR: {len(unique_indices)} страниц из {total_pages}")

        # Обработка каждой страницы
        page_texts = asyncio.run(
            self._process_pages_async(images, unique_indices)
        )

        return page_texts

    async def _process_pages_async(
        self,
        images: List[Image.Image],
        page_indices: List[int]
    ) -> Dict[int, str]:
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
                logger.error(f"Ошибка Paddle OCR стр.{page_num}: {e}")
                page_texts[page_num] = f"[ОШИБКА Paddle OCR: {e}]"

        return page_texts

    async def _process_single_page(
        self,
        page_image: Image.Image,
        page_num: int
    ) -> str:
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
        logger.info(f"=== Страница {page_num} ===")

        # Этап 1: Layout detection
        t0 = time.time()
        regions = self._detector.detect(page_image)
        layout_time = time.time() - t0

        class_counts = Counter(r["class_name"] for r in regions)
        logger.info(f"Layout: {len(regions)} регионов за {layout_time:.2f}s | {dict(class_counts)}")

        # Фильтрация мелких элементов (Page-header/Page-footer мусорного размера)
        page_w, page_h = page_image.size
        page_area = page_w * page_h
        min_area = page_area * config.SMALL_ELEMENT_MIN_AREA_RATIO

        filtered = []
        for r in regions:
            bbox = r["bbox"]
            box_area = (bbox[2] - bbox[0]) * (bbox[3] - bbox[1])
            if r["class_name"] in config.SMALL_ELEMENT_CLASSES and box_area < min_area:
                logger.debug(f"  Пропуск мелкого {r['class_name']}: area={box_area:.0f} < {min_area:.0f}")
                continue
            filtered.append(r)

        if len(filtered) != len(regions):
            logger.info(f"Отфильтровано мелких: {len(regions) - len(filtered)}")
        regions = filtered

        # Визуализация layout
        if self.log_dir:
            self.log_dir.mkdir(parents=True, exist_ok=True)
            vis_path = self.log_dir / f"page_{page_num}_layout.png"
            self._detector.visualize(page_image, regions, str(vis_path))

        # Этап 2: Разделение на таблицы и всё остальное
        table_regions = [r for r in regions if r["class_name"] == "Table"]
        table_regions.sort(key=lambda r: r["bbox"][1])  # сортировка по y сверху вниз

        n_tables = len(table_regions)
        logger.info(f"Таблиц: {n_tables}, full-page OCR + {n_tables} table crop(s)")

        # Этап 3: Подготовка VLM-запросов
        vlm_items = []

        # 3a. Full-page OCR (с маскировкой таблиц если есть)
        if table_regions:
            ocr_page = self._mask_table_regions(page_image, table_regions)
        else:
            ocr_page = page_image

        vlm_items.append({
            "image": ocr_page,
            "prompt": "OCR:",
            "index": -1,  # спец-индекс для full-page
        })

        # 3b. Кропы таблиц
        for i, table_reg in enumerate(table_regions):
            crop = self._crop_region(page_image, table_reg["bbox"])
            vlm_items.append({
                "image": crop,
                "prompt": "Table Recognition:",
                "index": i,
            })

        # Сохраняем для отладки
        if self.log_dir:
            if table_regions:
                masked_path = self.log_dir / f"page_{page_num}_masked.png"
                ocr_page.save(str(masked_path))
            for i in range(n_tables):
                crop_path = self.log_dir / f"page_{page_num}_table_{i}.png"
                vlm_items[1 + i]["image"].save(str(crop_path))

        # Этап 4: VLM batch (параллельно)
        t0 = time.time()
        vlm_results = await self._vlm.recognize_batch(vlm_items)
        vlm_time = time.time() - t0
        logger.info(f"VLM: {len(vlm_results)} запросов за {vlm_time:.2f}s")

        # Этап 5: Разбор результатов
        vlm_map = {r["index"]: r["text"] for r in vlm_results}
        full_page_text = vlm_map.get(-1, "")

        table_entries = []
        for i, table_reg in enumerate(table_regions):
            raw_table = vlm_map.get(i, "")
            if is_table_format(raw_table):
                html = parse_paddle_table_html(raw_table)
            else:
                html = raw_table
            table_entries.append({
                "bbox": table_reg["bbox"],
                "html": html,
            })

        # Сохраняем сырые данные для сверки
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

        # Этап 6: Склейка текста и таблиц
        md = merge_text_and_tables(full_page_text, table_entries, page_image.size)
        return md

    @staticmethod
    def _mask_table_regions(
        page_image: Image.Image,
        table_regions: list,
        padding: int = None,
    ) -> Image.Image:
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

        padding = padding if padding is not None else config.TABLE_MASK_PADDING_PX
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

    @staticmethod
    def _crop_region(
        page_image: Image.Image,
        bbox: list,
        padding: int = None
    ) -> Image.Image:
        """
        Вырезает регион из изображения страницы с padding.

        Args:
            page_image: изображение страницы.
            bbox: [x1, y1, x2, y2].
            padding: отступ в пикселях.

        Returns:
            Вырезанное изображение.
        """
        padding = padding if padding is not None else config.CROP_PADDING_PX
        x1, y1, x2, y2 = [int(c) for c in bbox]

        w, h = page_image.size
        x1c = max(0, x1 - padding)
        y1c = max(0, y1 - padding)
        x2c = min(w, x2 + padding)
        y2c = min(h, y2 + padding)

        return page_image.crop((x1c, y1c, x2c, y2c))
