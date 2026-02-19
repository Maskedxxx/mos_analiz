#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OCR Parser — локальный OCR-pipeline для извлечения текста из документов.

Архитектура:
    DOCX/PDF -> DocumentConverter -> PDF -> OcrExtractor (постранично)
                                                  |
                                          Dict[page_num, text]
                                                  |
                                      ChunkAssembler (по chunks_vision.json)
                                                  |
                                          Dict[chunk_name, text]
                                                  |
                                      ChunkAggregator (метаданные)
                                                  |
                                          Dict[str, Any] (legacy-формат)

Отличие от VisionParser:
- OCR-модель (HunyuanOCR) работает постранично с ЕДИНЫМ промптом
- Каждая страница OCR'ится ОДИН раз (дедупликация)
- Сборка в чанки — по номерам страниц из chunks_vision.json
- Работает через локальный vLLM, не облачный GPT
"""

import logging
from pathlib import Path
from typing import Any, Dict, Optional

from ..models import AuditConfig
from ..vision_parser.config import ChunkConfig, load_chunks_config
from ..vision_parser.converter import DocumentConverter
from ..vision_parser.aggregator import ChunkAggregator
from .extractor import OcrExtractor
from .assembler import ChunkAssembler, collect_unique_pages

log = logging.getLogger(__name__)

__all__ = ['OcrParser']


class OcrParser:
    """
    Главный класс OCR Pipeline.

    Объединяет:
    1. DocumentConverter — DOCX -> PDF
    2. OcrExtractor — PDF -> постраничный OCR через vLLM
    3. ChunkAssembler — сборка чанков из страниц по конфигу
    4. ChunkAggregator — добавление метаданных (имя файла, путь)

    Возвращает тот же формат что и VisionParser — Dict[str, Any].
    """

    def __init__(self, config: AuditConfig, log_dir: Optional[str] = None):
        """
        Инициализация OCR-парсера.

        Args:
            config: конфигурация аудита (содержит OCR-параметры)
            log_dir: директория для логов OCR
        """
        self.config = config

        # Загружаем конфигурацию чанков
        self.chunks = load_chunks_config(str(config.chunks_vision_path))

        # Компоненты pipeline
        self.converter = DocumentConverter()
        self.extractor = OcrExtractor(
            base_url=config.ocr_base_url,
            model=config.ocr_model,
            prompt=config.ocr_prompt,
            dpi=config.ocr_dpi,
            log_dir=log_dir
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
        Парсинг документа через OCR Pipeline.

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
            ConnectionError: vLLM OCR недоступен
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

            # Убираем пустые страницы из списка OCR (на случай явного указания)
            unique_page_indices = [i for i in unique_page_indices if i not in blank_pages]
            log.info(f"Уникальных страниц для OCR: {len(unique_page_indices)} из {total_pages}")

            # 5. OCR всех уникальных страниц (одним батчем)
            page_texts = self.extractor.extract_pages(pdf_path, unique_page_indices)

            # 6. Сборка чанков из страниц (тоже с effective_total для [-1])
            chunk_texts = self.assembler.assemble(
                page_texts, chunks, effective_total,
                strip_annotations_chunks=self.config.strip_annotations_chunks
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
