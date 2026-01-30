#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Агрегатор чанков — нормализация и объединение извлечённых данных.

Модуль преобразует сырой вывод Vision LLM в формат,
совместимый с legacy-парсером (python-docx).
"""

import re
from pathlib import Path
from typing import Any, Dict


class ChunkAggregator:
    """
    Агрегатор для объединения и нормализации чанков.

    Преобразует сырые результаты Vision API в формат,
    совместимый с legacy-парсером для обратной совместимости.
    """

    def aggregate(
        self,
        raw_chunks: Dict[str, str],
        file_path: str
    ) -> Dict[str, Any]:
        """
        Агрегирует и нормализует извлечённые чанки.

        Args:
            raw_chunks: Словарь {chunk_name: raw_text} от VisionExtractor
            file_path: Путь к исходному файлу

        Returns:
            Словарь в формате legacy-парсера:
            {
                "имя_файла": str,
                "путь": str,
                "шапка": str,
                "преамбула": str,
                ...
            }
        """
        path = Path(file_path)

        # Базовые метаданные (как в legacy-парсере)
        result = {
            "имя_файла": path.name,
            "путь": str(path.absolute()),
        }

        # Добавляем нормализованные чанки
        for chunk_name, raw_text in raw_chunks.items():
            normalized = self._normalize(raw_text)
            result[chunk_name] = normalized

        return result

    def _normalize(self, text: str) -> str:
        """
        Нормализует текст чанка.

        Операции:
        1. Убирает markdown-разметку (если LLM её добавил)
        2. Нормализует пробелы и переносы строк
        3. Убирает артефакты OCR

        Args:
            text: Сырой текст от Vision LLM

        Returns:
            Нормализованный текст
        """
        if not text:
            return ""

        # Убираем markdown code blocks
        text = re.sub(r'```[\w]*\n?', '', text)
        text = re.sub(r'```', '', text)

        # Убираем markdown headers (# Header)
        text = re.sub(r'^#+\s*', '', text, flags=re.MULTILINE)

        # Убираем markdown bold/italic
        text = re.sub(r'\*\*([^*]+)\*\*', r'\1', text)
        text = re.sub(r'\*([^*]+)\*', r'\1', text)
        text = re.sub(r'__([^_]+)__', r'\1', text)
        text = re.sub(r'_([^_]+)_', r'\1', text)

        # Нормализуем пробелы
        # Множественные пробелы → один (кроме начала строки)
        text = re.sub(r'([^\n]) {2,}', r'\1 ', text)

        # Множественные пустые строки → одна
        text = re.sub(r'\n{3,}', '\n\n', text)

        # Убираем пробелы в конце строк
        text = re.sub(r' +\n', '\n', text)

        # Убираем артефакты OCR (типичные мусорные символы)
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)

        # Нормализуем тире и дефисы
        text = text.replace('–', '-')  # en-dash
        text = text.replace('—', '-')  # em-dash

        # Нормализуем кавычки
        text = text.replace('«', '"')
        text = text.replace('»', '"')
        text = text.replace('„', '"')
        text = text.replace('"', '"')
        text = text.replace('"', '"')

        # Убираем пробелы в начале и конце
        text = text.strip()

        return text

    def merge_chunks(
        self,
        chunks: Dict[str, str],
        order: list = None
    ) -> str:
        """
        Объединяет несколько чанков в один текст.

        Args:
            chunks: Словарь чанков
            order: Порядок объединения (если не указан, используется порядок словаря)

        Returns:
            Объединённый текст
        """
        if order is None:
            order = list(chunks.keys())

        parts = []
        for name in order:
            if name in chunks and chunks[name]:
                parts.append(chunks[name])

        return "\n\n".join(parts)


def aggregate_results(
    raw_chunks: Dict[str, str],
    file_path: str
) -> Dict[str, Any]:
    """
    Утилитарная функция для быстрой агрегации.

    Args:
        raw_chunks: Сырые чанки от VisionExtractor
        file_path: Путь к исходному файлу

    Returns:
        Агрегированный результат
    """
    aggregator = ChunkAggregator()
    return aggregator.aggregate(raw_chunks, file_path)
