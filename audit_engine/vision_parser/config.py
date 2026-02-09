#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конфигурация для Vision Parser.

Модуль содержит dataclass-ы для конфигурации чанков и параметров Vision API,
а также функции загрузки конфигурации из JSON.
"""

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Union


@dataclass
class ChunkConfig:
    """
    Конфигурация одного чанка документа.

    Attributes:
        name: Имя чанка (ключ в результирующем словаре)
        pages: Страницы для обработки:
               - List[int]: конкретные номера страниц [1, 2, 3]
               - str "all": все страницы
               - str "1-3": диапазон страниц
               - List с отрицательными: [-1] = последняя страница
        prompt: Промпт для Vision LLM с инструкциями по извлечению
    """
    name: str
    pages: Union[List[int], str]
    prompt: str

    def get_page_indices(self, total_pages: int) -> List[int]:
        """
        Преобразует pages в список индексов страниц (0-based).

        Args:
            total_pages: Общее количество страниц в документе

        Returns:
            Список индексов страниц (0-based)
        """
        if self.pages == "all":
            return list(range(total_pages))

        if isinstance(self.pages, str):
            # Парсим диапазон "1-3" -> [0, 1, 2]
            if "-" in self.pages:
                parts = self.pages.split("-")
                start = int(parts[0]) - 1  # 1-based -> 0-based
                end = int(parts[1])        # включительно
                return list(range(start, end))
            else:
                # Одна страница как строка
                return [int(self.pages) - 1]

        if isinstance(self.pages, list):
            result = []
            for p in self.pages:
                if p < 0:
                    # Отрицательные индексы: -1 = последняя
                    result.append(total_pages + p)
                else:
                    # 1-based -> 0-based
                    result.append(p - 1)
            return result

        return []


@dataclass
class VisionConfig:
    """
    Конфигурация Vision Pipeline.

    Attributes:
        model: Модель OpenAI для Vision API
        dpi: Разрешение для рендеринга PDF (200 = хороший баланс качества/размера)
        max_workers: Максимум параллельных запросов к API
        timeout: Таймаут одного запроса в секундах
        max_retries: Максимум повторных попыток при ошибке
        detail: Уровень детализации для Vision API ("low", "high", "auto")
    """
    model: str = "gpt-4.1-mini"
    dpi: int = 200
    max_workers: int = 4
    timeout: int = 60
    max_retries: int = 3
    detail: str = "high"


def load_chunks_config(config_path: str) -> List[ChunkConfig]:
    """
    Загружает конфигурацию чанков из JSON-файла.

    Args:
        config_path: Путь к JSON-файлу с конфигурацией

    Returns:
        Список ChunkConfig

    Raises:
        FileNotFoundError: Если файл не найден
        json.JSONDecodeError: Если JSON невалидный
    """
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"Конфигурация чанков не найдена: {config_path}")

    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    chunks = []
    for chunk_data in data.get("chunks", []):
        chunk = ChunkConfig(
            name=chunk_data["name"],
            pages=chunk_data["pages"],
            prompt=chunk_data["prompt"]
        )
        chunks.append(chunk)

    return chunks


def load_vision_config(config_path: str = None) -> VisionConfig:
    """
    Загружает конфигурацию Vision Pipeline.

    Args:
        config_path: Путь к JSON-файлу (опционально).
                    Если не указан, возвращает дефолтную конфигурацию.

    Returns:
        VisionConfig с параметрами
    """
    if config_path is None:
        return VisionConfig()

    path = Path(config_path)
    if not path.exists():
        return VisionConfig()

    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    vision_data = data.get("vision", {})
    return VisionConfig(
        model=vision_data.get("model", "gpt-4.1-mini"),
        dpi=vision_data.get("dpi", 200),
        max_workers=vision_data.get("max_workers", 4),
        timeout=vision_data.get("timeout", 60),
        max_retries=vision_data.get("max_retries", 3),
        detail=vision_data.get("detail", "high")
    )
