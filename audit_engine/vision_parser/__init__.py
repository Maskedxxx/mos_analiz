#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vision Parser — Universal Vision Pipeline для извлечения текста из документов.

Главный модуль, предоставляющий VisionParser — единый интерфейс для
извлечения структурированного текста из DOCX, PDF и сканов
с использованием Vision LLM (gpt-4.1-mini).

Архитектура:
    DOCX/PDF → DocumentConverter → PDF → VisionExtractor → raw_chunks
                                                              ↓
                                        ChunkAggregator ← Dict[str, str]
                                              ↓
                                    Dict[str, Any] (совместим с legacy)

Использование:
    from shared.vision_parser import VisionParser

    parser = VisionParser("config/chunks_vision.json")
    result = parser.parse("document.docx")

    # result = {
    #     "имя_файла": "document.docx",
    #     "путь": "/path/to/document.docx",
    #     "шапка": "...",
    #     "преамбула": "...",
    #     ...
    # }
"""

from pathlib import Path
from typing import Any, Dict, List, Optional

from .config import ChunkConfig, VisionConfig, load_chunks_config, load_vision_config
from .converter import DocumentConverter, ConversionError
from .extractor import VisionExtractor, ExtractionError
from .aggregator import ChunkAggregator


__all__ = [
    'VisionParser',
    'ChunkConfig',
    'VisionConfig',
    'load_chunks_config',
    'load_vision_config',
    'DocumentConverter',
    'VisionExtractor',
    'ChunkAggregator',
    'ConversionError',
    'ExtractionError',
]

__version__ = '0.1.0'


class VisionParser:
    """
    Главный класс Vision Pipeline.

    Объединяет все компоненты пайплайна:
    1. DocumentConverter — DOCX → PDF
    2. VisionExtractor — PDF → chunks через Vision LLM
    3. ChunkAggregator — нормализация и форматирование результата

    Атрибуты:
        chunks_config: Список конфигураций чанков
        vision_config: Параметры Vision API
        converter: Конвертер документов
        extractor: Vision экстрактор
        aggregator: Агрегатор результатов
    """

    def __init__(
        self,
        chunks_config: str,
        vision_config: VisionConfig = None,
        temp_dir: str = None,
        log_dir: str = None
    ):
        """
        Инициализация Vision Parser.

        Args:
            chunks_config: Путь к JSON с конфигурацией чанков
            vision_config: Конфигурация Vision API (опционально)
            temp_dir: Директория для временных файлов (опционально)
            log_dir: Директория для логирования промптов Vision (опционально)
        """
        # Загружаем конфигурации
        self.chunks_config = load_chunks_config(chunks_config)
        # Загружаем vision_config из JSON если не передан явно
        if vision_config is None:
            self.vision_config = load_vision_config(chunks_config)
        else:
            self.vision_config = vision_config

        # Создаём компоненты
        self.converter = DocumentConverter(temp_dir)
        self.extractor = VisionExtractor(self.vision_config, log_dir)
        self.aggregator = ChunkAggregator()

        # Флаг для отслеживания временных файлов
        self._temp_pdf: Optional[str] = None

    def parse(
        self,
        file_path: str,
        chunk_filter: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Главный метод — парсинг документа через Vision Pipeline.

        Args:
            file_path: Путь к документу (DOCX, PDF или скан)
            chunk_filter: Имя чанка для отладки (если указан — парсим только его)

        Returns:
            Словарь с извлечёнными чанками в формате legacy-парсера:
            {
                "имя_файла": str,
                "путь": str,
                "шапка": str,
                "преамбула": str,
                "текст_приказа": str,
                ...
            }

        Raises:
            FileNotFoundError: Если файл не найден
            ConversionError: Если конвертация не удалась
            ExtractionError: Если извлечение не удалось
        """
        input_path = Path(file_path)

        if not input_path.exists():
            raise FileNotFoundError(f"Файл не найден: {file_path}")

        try:
            # 1. Конвертируем в PDF (если нужно)
            pdf_path = self.converter.convert(str(input_path))

            # Запоминаем, если был создан временный PDF
            if pdf_path != str(input_path.absolute()):
                self._temp_pdf = pdf_path

            # 2. Извлекаем чанки через Vision API
            raw_chunks = self.extractor.extract(
                pdf_path,
                self.chunks_config,
                chunk_filter=chunk_filter
            )

            # 3. Агрегируем результат
            result = self.aggregator.aggregate(raw_chunks, str(input_path))

            return result

        finally:
            # Очищаем временный PDF
            self._cleanup()

    def parse_pages(
        self,
        file_path: str,
        pages: List[int],
        prompt: str
    ) -> str:
        """
        Извлекает текст с указанных страниц по произвольному промпту.

        Полезно для ad-hoc извлечения без предварительной конфигурации.

        Args:
            file_path: Путь к документу
            pages: Список номеров страниц (1-based)
            prompt: Промпт для Vision LLM

        Returns:
            Извлечённый текст
        """
        # Создаём временный ChunkConfig
        temp_chunk = ChunkConfig(
            name="_temp",
            pages=pages,
            prompt=prompt
        )

        input_path = Path(file_path)

        try:
            # Конвертируем в PDF
            pdf_path = self.converter.convert(str(input_path))

            if pdf_path != str(input_path.absolute()):
                self._temp_pdf = pdf_path

            # Извлекаем
            raw_chunks = self.extractor.extract(
                pdf_path,
                [temp_chunk]
            )

            return raw_chunks.get("_temp", "")

        finally:
            self._cleanup()

    def _cleanup(self):
        """Очищает временные файлы."""
        if self._temp_pdf:
            self.converter.cleanup(self._temp_pdf)
            self._temp_pdf = None

    def __enter__(self):
        """Поддержка context manager."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Автоматическая очистка."""
        self._cleanup()
        return False


def parse_document(
    file_path: str,
    chunks_config: str,
    vision_config: VisionConfig = None
) -> Dict[str, Any]:
    """
    Утилитарная функция для быстрого парсинга.

    Args:
        file_path: Путь к документу
        chunks_config: Путь к JSON с конфигурацией чанков
        vision_config: Конфигурация Vision API (опционально)

    Returns:
        Словарь с извлечёнными чанками
    """
    with VisionParser(chunks_config, vision_config) as parser:
        return parser.parse(file_path)
