#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Paddle Parser — layout-aware OCR pipeline для извлечения текста из документов.

Архитектура:
    DOCX/PDF -> DocumentConverter -> PDF -> PaddleExtractor (layout + VLM постранично)
                                                  |
                                          Dict[page_num, markdown_text]
                                                  |
                                      ChunkAssembler (по chunks_vision.json,
                                                      normalizer_fn=normalize_paddle_text)
                                                  |
                                          Dict[chunk_name, text]
                                                  |
                                      ChunkAggregator (метаданные)
                                                  |
                                          Dict[str, Any] (legacy-формат)

Отличие от OcrParser:
- Layout detection (Heron-101) → crop → VLM (PaddleOCR-VL-1.5) по классу региона
- Таблицы сохраняются в HTML с colspan/rowspan (не конвертируются в markdown)
- Нормализация без LaTeX-артефактов и html_table_to_markdown
"""

import logging
from pathlib import Path
from typing import Any, Dict, Optional

from ..models import AuditConfig
from ..vision_parser.config import load_chunks_config
from ..vision_parser.converter import DocumentConverter
from ..vision_parser.aggregator import ChunkAggregator
from ..ocr_parser.assembler import ChunkAssembler, collect_unique_pages
from .extractor import PaddleExtractor
from .normalizer import normalize_paddle_text

log = logging.getLogger(__name__)

__all__ = ['PaddleParser']


class PaddleParser:
    """
    Главный класс Paddle OCR Pipeline.

    Объединяет:
    1. DocumentConverter — DOCX → PDF
    2. PaddleExtractor — PDF → layout + VLM OCR → постраничный markdown
    3. ChunkAssembler — сборка чанков из страниц по конфигу
    4. ChunkAggregator — добавление метаданных (имя файла, путь)

    Возвращает тот же формат что VisionParser и OcrParser — Dict[str, Any].
    """

    def __init__(self, config: AuditConfig, log_dir: Optional[str] = None):
        """
        Инициализация Paddle-парсера.

        Args:
            config: конфигурация аудита (содержит paddle_* и ocr_* параметры)
            log_dir: директория для логов (layout-визуализации, crop'ы, raw.json)
        """
        self.config = config

        # Загружаем конфигурацию чанков
        self.chunks = load_chunks_config(str(config.chunks_vision_path))

        # Компоненты pipeline
        self.converter = DocumentConverter()
        self.extractor = PaddleExtractor(
            vlm_base_url=config.ocr_base_url,   # переиспользуем — тот же формат API
            vlm_model=config.paddle_vlm_model,
            layout_model_repo=config.paddle_layout_model,
            layout_device=config.paddle_layout_device,
            layout_base_url=config.paddle_layout_base_url,  # удалённый Layout API (опционально)
            dpi=config.ocr_dpi,
            log_dir=log_dir,
        )
        self.assembler = ChunkAssembler()
        self.aggregator = ChunkAggregator()

        self._temp_pdf: Optional[str] = None

    def parse(
        self,
        file_path: str,
        chunk_filter: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Парсинг документа через Paddle OCR Pipeline.

        Args:
            file_path: путь к документу (DOCX, PDF)
            chunk_filter: имя чанка (если указан — собираем только его)

        Returns:
            Dict[str, Any] в формате legacy-парсера:
            {
                "имя_файла": str,
                "путь": str,
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
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        try:
            # 1. Конвертация в PDF (если нужно)
            pdf_path = self.converter.convert(str(input_path))
            if pdf_path != str(input_path.absolute()):
                self._temp_pdf = pdf_path

            # 2. Определяем какие чанки парсить
            chunks = self.chunks
            if chunk_filter:
                chunks = [c for c in self.chunks if c.name == chunk_filter]
                if not chunks:
                    raise ValueError(f"Чанк '{chunk_filter}' не найден в конфигурации")

            # 3. Детекция пустых страниц (LibreOffice добавляет при смене ориентации)
            total_pages, blank_pages = self.extractor.detect_blank_pages(pdf_path)

            # effective_total — для разрешения отрицательных индексов ([-1] = последняя непустая)
            non_blank = [i for i in range(total_pages) if i not in blank_pages]
            effective_total = (max(non_blank) + 1) if non_blank else total_pages
            if effective_total != total_pages:
                log.info(f"Страниц в PDF: {total_pages}, эффективных: {effective_total} "
                         f"(пустые: {sorted(i+1 for i in blank_pages)})")

            # 4. Собираем уникальные страницы ([-1] разрешается по effective_total)
            unique_page_indices = collect_unique_pages(chunks, effective_total)

            # Убираем пустые страницы из списка
            unique_page_indices = [i for i in unique_page_indices if i not in blank_pages]
            log.info(f"Уникальных страниц для Paddle OCR: {len(unique_page_indices)} из {total_pages}")

            # 5. Layout-aware OCR всех уникальных страниц
            page_texts = self.extractor.extract_pages(pdf_path, unique_page_indices)

            # 6. Сборка чанков из страниц (с paddle-нормализацией)
            chunk_texts = self.assembler.assemble(
                page_texts, chunks, effective_total,
                strip_annotations_chunks=self.config.strip_annotations_chunks,
                normalizer_fn=normalize_paddle_text,
            )

            # 7. Агрегация в legacy-формат (добавляет имя_файла, путь)
            result = self.aggregator.aggregate(chunk_texts, str(input_path))

            return result

        finally:
            self._cleanup()

    def _cleanup(self):
        """Очистка временных файлов."""
        if self._temp_pdf:
            self.converter.cleanup(self._temp_pdf)
            self._temp_pdf = None
