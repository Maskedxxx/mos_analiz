#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vision Extractor — извлечение текста из PDF через Vision LLM.

Модуль конвертирует страницы PDF в изображения и отправляет их
в Vision API (gpt-4.1-mini) для извлечения структурированного текста.
"""

import base64
import io
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Dict, List, Optional

from PIL import Image
from openai import OpenAI
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

from .config import ChunkConfig, VisionConfig


class ExtractionError(Exception):
    """Ошибка извлечения текста."""
    pass


class VisionExtractor:
    """
    Извлекает текст из PDF через Vision LLM.

    Пайплайн:
    1. PDF → PNG[] (pdf2image)
    2. Для каждого чанка параллельно:
       - Выбор страниц по ChunkConfig.pages
       - Кодирование в base64
       - Вызов Vision API с промптом
       - Retry при ошибках
    3. Возврат Dict[chunk_name → extracted_text]

    Атрибуты:
        config: Параметры Vision API
        client: Клиент OpenAI
        log_dir: Директория для логирования промптов (опционально)
    """

    # Системный промпт для Vision LLM
    SYSTEM_PROMPT = """Ты — OCR система. Твоя ЕДИНСТВЕННАЯ задача — читать текст с изображения.

КРИТИЧЕСКИ ВАЖНО:
- Выводи ТОЛЬКО тот текст, который РЕАЛЬНО НАПИСАН на изображении
- НИКОГДА не выдумывай и не дополняй текст
- НИКОГДА не заменяй слова на похожие
- Если не можешь прочитать — напиши "НЕ ЧИТАЕТСЯ"
- Читай ДОСЛОВНО, символ в символ

Формат: только текст с изображения, без комментариев."""

    def __init__(self, config: VisionConfig = None, log_dir: str = None):
        """
        Инициализация экстрактора.

        Args:
            config: Конфигурация Vision API. Если не указана, используется дефолтная.
            log_dir: Директория для логирования промптов и ответов.
        """
        self.config = config or VisionConfig()
        self.client = OpenAI()  # Использует OPENAI_API_KEY из env
        self.log_dir = Path(log_dir) if log_dir else None

        # Создаём директорию для логов Vision
        if self.log_dir:
            self.vision_prompts_dir = self.log_dir / "vision_prompts"
            self.vision_prompts_dir.mkdir(parents=True, exist_ok=True)

    def extract(
        self,
        pdf_path: str,
        chunks: List[ChunkConfig],
        parallel: bool = True
    ) -> Dict[str, str]:
        """
        Главный метод извлечения текста из PDF.

        Args:
            pdf_path: Путь к PDF файлу
            chunks: Список конфигураций чанков для извлечения
            parallel: Использовать параллельную обработку

        Returns:
            Словарь {chunk_name: extracted_text}

        Raises:
            FileNotFoundError: Если PDF не найден
            ExtractionError: Если извлечение не удалось
        """
        # Проверяем файл
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF не найден: {pdf_path}")

        # Конвертируем PDF в изображения
        images = self._pdf_to_images(pdf_path)
        total_pages = len(images)

        if total_pages == 0:
            raise ExtractionError("PDF не содержит страниц")

        # Извлекаем чанки
        results = {}

        if parallel and len(chunks) > 1:
            # Параллельное извлечение
            results = self._extract_parallel(images, chunks, total_pages)
        else:
            # Последовательное извлечение
            for chunk in chunks:
                text = self._extract_chunk(images, chunk, total_pages)
                results[chunk.name] = text

        return results

    def _pdf_to_images(self, pdf_path: str) -> List[Image.Image]:
        """
        Конвертирует PDF в список изображений.

        Args:
            pdf_path: Путь к PDF файлу

        Returns:
            Список PIL Image объектов (по одному на страницу)

        Raises:
            ExtractionError: Если конвертация не удалась
        """
        try:
            # Импортируем pdf2image здесь, чтобы ошибка была понятнее
            from pdf2image import convert_from_path
        except ImportError:
            raise ExtractionError(
                "pdf2image не установлен. Установите:\n"
                "  pip install pdf2image\n"
                "  # И poppler:\n"
                "  macOS: brew install poppler\n"
                "  Linux: apt install poppler-utils"
            )

        try:
            images = convert_from_path(
                pdf_path,
                dpi=self.config.dpi,
                fmt='PNG'
            )
            return images
        except Exception as e:
            raise ExtractionError(f"Ошибка конвертации PDF: {e}")

    def _extract_parallel(
        self,
        images: List[Image.Image],
        chunks: List[ChunkConfig],
        total_pages: int
    ) -> Dict[str, str]:
        """
        Параллельное извлечение чанков.

        Args:
            images: Список изображений страниц
            chunks: Конфигурации чанков
            total_pages: Общее количество страниц

        Returns:
            Словарь {chunk_name: extracted_text}
        """
        results = {}

        with ThreadPoolExecutor(max_workers=self.config.max_workers) as executor:
            # Создаём задачи
            futures = {
                executor.submit(
                    self._extract_chunk,
                    images,
                    chunk,
                    total_pages
                ): chunk
                for chunk in chunks
            }

            # Собираем результаты
            for future in as_completed(futures):
                chunk = futures[future]
                try:
                    text = future.result()
                    results[chunk.name] = text
                except Exception as e:
                    # Записываем ошибку в результат
                    results[chunk.name] = f"[ОШИБКА ИЗВЛЕЧЕНИЯ: {e}]"

        return results

    def _extract_chunk(
        self,
        images: List[Image.Image],
        chunk: ChunkConfig,
        total_pages: int
    ) -> str:
        """
        Извлекает один чанк из указанных страниц.

        Args:
            images: Список изображений страниц
            chunk: Конфигурация чанка
            total_pages: Общее количество страниц

        Returns:
            Извлечённый текст
        """
        # Получаем индексы страниц для чанка
        page_indices = chunk.get_page_indices(total_pages)

        # Фильтруем валидные индексы
        valid_indices = [i for i in page_indices if 0 <= i < total_pages]

        if not valid_indices:
            return "[НЕТ СТРАНИЦ ДЛЯ ИЗВЛЕЧЕНИЯ]"

        # Выбираем нужные страницы
        chunk_images = [images[i] for i in valid_indices]

        # Вызываем Vision API
        result = self._call_vision_api(chunk_images, chunk.prompt, chunk.name, valid_indices)

        return result

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=2, max=10),
        retry=retry_if_exception_type((Exception,)),
        reraise=True
    )
    def _call_vision_api(
        self,
        images: List[Image.Image],
        prompt: str,
        chunk_name: str = None,
        page_indices: List[int] = None
    ) -> str:
        """
        Вызывает Vision API для извлечения текста.

        Args:
            images: Список изображений для анализа
            prompt: Инструкция для извлечения
            chunk_name: Имя чанка (для логирования)
            page_indices: Индексы страниц (для логирования)

        Returns:
            Извлечённый текст

        Raises:
            ExtractionError: Если API вернул ошибку
        """
        # Формируем content с изображениями
        content = []

        # Добавляем изображения
        for img in images:
            base64_image = self._image_to_base64(img)
            content.append({
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/png;base64,{base64_image}",
                    "detail": self.config.detail
                }
            })

        # Добавляем промпт
        content.append({
            "type": "text",
            "text": prompt
        })

        # Логируем промпт если указана директория
        if self.log_dir and chunk_name:
            self._log_vision_prompt(chunk_name, page_indices or [], prompt)

        # Вызываем API
        try:
            response = self.client.chat.completions.create(
                model=self.config.model,
                messages=[
                    {"role": "system", "content": self.SYSTEM_PROMPT},
                    {"role": "user", "content": content}
                ],
                max_tokens=8096,
                temperature=0.0
            )

            result = response.choices[0].message.content.strip()

            # Логируем ответ
            if self.log_dir and chunk_name:
                self._log_vision_response(chunk_name, result)

            return result

        except Exception as e:
            raise ExtractionError(f"Ошибка Vision API: {e}")

    def _log_vision_prompt(self, chunk_name: str, page_indices: List[int], user_prompt: str):
        """Логирует промпт Vision API."""
        if not self.vision_prompts_dir:
            return

        filepath = self.vision_prompts_dir / f"chunk_{chunk_name}_prompt.txt"
        pages_str = ", ".join(str(i + 1) for i in page_indices)  # 1-based для читаемости

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"VISION PROMPT - Чанк: {chunk_name}\n")
            f.write(f"{'='*80}\n\n")
            f.write(f"Модель: {self.config.model}\n")
            f.write(f"Страницы (1-based): {pages_str}\n")
            f.write(f"Количество изображений: {len(page_indices)}\n")
            f.write(f"Detail: {self.config.detail}\n\n")
            f.write(f"--- SYSTEM PROMPT ---\n")
            f.write(self.SYSTEM_PROMPT)
            f.write(f"\n\n--- USER PROMPT ---\n")
            f.write(user_prompt)

    def _log_vision_response(self, chunk_name: str, response: str):
        """Логирует ответ Vision API."""
        if not self.vision_prompts_dir:
            return

        filepath = self.vision_prompts_dir / f"chunk_{chunk_name}_response.txt"

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"VISION RESPONSE - Чанк: {chunk_name}\n")
            f.write(f"{'='*80}\n\n")
            f.write(f"Длина ответа: {len(response)} символов\n\n")
            f.write(f"--- ОТВЕТ ---\n")
            f.write(response)

    def _image_to_base64(self, image: Image.Image) -> str:
        """
        Конвертирует PIL Image в base64 строку.

        Args:
            image: PIL Image объект

        Returns:
            Base64-encoded строка
        """
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        buffer.seek(0)
        return base64.b64encode(buffer.read()).decode("utf-8")


def extract_from_pdf(
    pdf_path: str,
    chunks_config_path: str,
    vision_config: VisionConfig = None
) -> Dict[str, str]:
    """
    Утилитарная функция для быстрого извлечения.

    Args:
        pdf_path: Путь к PDF
        chunks_config_path: Путь к JSON с конфигурацией чанков
        vision_config: Конфигурация Vision API (опционально)

    Returns:
        Словарь {chunk_name: extracted_text}
    """
    from .config import load_chunks_config

    chunks = load_chunks_config(chunks_config_path)
    extractor = VisionExtractor(vision_config)

    return extractor.extract(pdf_path, chunks)
