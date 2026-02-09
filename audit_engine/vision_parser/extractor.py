#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Vision Extractor — извлечение текста из PDF через Vision LLM.

Модуль конвертирует страницы PDF в изображения и отправляет их
в Vision API (gpt-4.1-mini) для извлечения структурированного текста.

Особенности:
- Каждая страница обрабатывается ОТДЕЛЬНЫМ запросом (предотвращает галлюцинации)
- Параллельная обработка через AsyncOpenAI + asyncio.gather()
- Поддержка фильтрации чанков для отладки
"""

import asyncio
import base64
import io
from pathlib import Path
from typing import Dict, List, Optional

from PIL import Image
from openai import AsyncOpenAI
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
    2. Для каждого чанка:
       - Каждая страница → отдельный async запрос к Vision API
       - Все страницы чанка обрабатываются параллельно (asyncio.gather)
       - Результаты склеиваются в порядке страниц
    3. Возврат Dict[chunk_name → extracted_text]
    """

    # Системный промпт для Vision LLM — строгий OCR режим
    SYSTEM_PROMPT = """Ты — OCR система. Твоя ЕДИНСТВЕННАЯ задача — читать текст с изображения.

АБСОЛЮТНЫЕ ПРАВИЛА:
1. Выводи ТОЛЬКО тот текст, который РЕАЛЬНО ВИДЕН на изображении
2. НИКОГДА не выдумывай текст, которого нет на изображении
3. НИКОГДА не дополняй, не продолжай, не "улучшай" текст
4. НИКОГДА не добавляй нумерацию (1.1, 3.5 и т.п.) если её нет на изображении
5. НИКОГДА не структурируй текст по своему усмотрению
6. Если текст обрезан или неполный — выводи только видимую часть
7. Если не можешь прочитать символ — пиши [?]
8. Читай СТРОГО слева направо, сверху вниз

ЗАПРЕЩЕНО:
- Додумывать что "должно быть" в документе
- Добавлять пункты которых нет
- Менять порядок слов
- Исправлять "ошибки" в тексте

Формат: ТОЛЬКО текст с изображения, без комментариев и пояснений."""

    def __init__(self, config: Optional[VisionConfig] = None, log_dir: Optional[str] = None):
        """
        Инициализация экстрактора.

        Args:
            config: Конфигурация Vision API. Если не указана, используется дефолтная.
            log_dir: Директория для логирования промптов и ответов.
        """
        self.config = config or VisionConfig()
        self.client = AsyncOpenAI()  # Асинхронный клиент OpenAI
        self.log_dir = Path(log_dir) if log_dir else None

        # Создаём директорию для логов Vision
        if self.log_dir:
            self.vision_prompts_dir = self.log_dir / "vision_prompts"
            self.vision_prompts_dir.mkdir(parents=True, exist_ok=True)
        else:
            self.vision_prompts_dir = None

    def extract(
        self,
        pdf_path: str,
        chunks: List[ChunkConfig],
        chunk_filter: Optional[str] = None
    ) -> Dict[str, str]:
        """
        Главный метод извлечения текста из PDF.

        Args:
            pdf_path: Путь к PDF файлу
            chunks: Список конфигураций чанков для извлечения
            chunk_filter: Имя чанка для отладки (если указан — парсим только его)

        Returns:
            Словарь {chunk_name: extracted_text}
        """
        # Проверяем файл
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF не найден: {pdf_path}")

        # Конвертируем PDF в изображения
        images = self._pdf_to_images(pdf_path)
        total_pages = len(images)

        if total_pages == 0:
            raise ExtractionError("PDF не содержит страниц")

        # Фильтруем чанки если указан фильтр
        if chunk_filter:
            chunks = [c for c in chunks if c.name == chunk_filter]
            if not chunks:
                raise ExtractionError(f"Чанк '{chunk_filter}' не найден в конфигурации")

        # Запускаем async извлечение
        return asyncio.run(self._extract_all_chunks(images, chunks, total_pages))

    async def _extract_all_chunks(
        self,
        images: List[Image.Image],
        chunks: List[ChunkConfig],
        total_pages: int
    ) -> Dict[str, str]:
        """
        Асинхронное извлечение всех чанков.

        Args:
            images: Список изображений страниц
            chunks: Конфигурации чанков
            total_pages: Общее количество страниц

        Returns:
            Словарь {chunk_name: extracted_text}
        """
        # Создаём задачи для всех чанков
        tasks = []
        chunk_names = []

        for chunk in chunks:
            tasks.append(self._extract_chunk_async(images, chunk, total_pages))
            chunk_names.append(chunk.name)

        # Выполняем все задачи параллельно
        results_list = await asyncio.gather(*tasks, return_exceptions=True)

        # Собираем результаты
        results = {}
        for name, result in zip(chunk_names, results_list):
            if isinstance(result, Exception):
                results[name] = f"[ОШИБКА: {result}]"
            else:
                results[name] = result

        return results

    async def _extract_chunk_async(
        self,
        images: List[Image.Image],
        chunk: ChunkConfig,
        total_pages: int
    ) -> str:
        """
        Асинхронное извлечение одного чанка.

        Каждая страница обрабатывается ОТДЕЛЬНЫМ запросом к API,
        затем результаты склеиваются. Это предотвращает галлюцинации.

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

        # Если одна страница — простой вызов
        if len(valid_indices) == 1:
            page_idx = valid_indices[0]
            return await self._call_vision_api_async(
                images[page_idx],
                chunk.prompt,
                f"{chunk.name}",
                page_idx
            )

        # Несколько страниц — параллельная обработка каждой
        # Добавляем тег страницы к промпту для контекста
        tasks = []
        for i, page_idx in enumerate(valid_indices):
            # Формируем промпт с указанием номера страницы
            page_num = page_idx + 1
            page_prompt = f"[СТРАНИЦА {page_num}]\n\n{chunk.prompt}"

            tasks.append(
                self._call_vision_api_async(
                    images[page_idx],
                    page_prompt,
                    f"{chunk.name}_p{page_num}",
                    page_idx
                )
            )

        # Выполняем все запросы параллельно
        page_results = await asyncio.gather(*tasks, return_exceptions=True)

        # Склеиваем результаты в порядке страниц
        combined_parts = []
        for i, result in enumerate(page_results):
            if isinstance(result, Exception):
                combined_parts.append(f"[ОШИБКА СТРАНИЦЫ {valid_indices[i] + 1}: {result}]")
            elif result and result.strip():
                combined_parts.append(result.strip())

        return "\n\n".join(combined_parts)

    def _pdf_to_images(self, pdf_path: str) -> List[Image.Image]:
        """
        Конвертирует PDF в список изображений.

        Args:
            pdf_path: Путь к PDF файлу

        Returns:
            Список PIL Image объектов (по одному на страницу)
        """
        try:
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

    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=2, max=10),
        retry=retry_if_exception_type((Exception,)),
        reraise=True
    )
    async def _call_vision_api_async(
        self,
        image: Image.Image,
        prompt: str,
        chunk_name: str,
        page_idx: int
    ) -> str:
        """
        Асинхронный вызов Vision API для одной страницы.

        Args:
            image: Изображение страницы
            prompt: Инструкция для извлечения
            chunk_name: Имя чанка (для логирования)
            page_idx: Индекс страницы (0-based)

        Returns:
            Извлечённый текст
        """
        # Кодируем изображение в base64
        base64_image = self._image_to_base64(image)

        # Формируем content
        content = [
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/png;base64,{base64_image}",
                    "detail": self.config.detail
                }
            },
            {
                "type": "text",
                "text": prompt
            }
        ]

        # Логируем промпт
        if self.vision_prompts_dir:
            self._log_vision_prompt(chunk_name, page_idx, prompt)

        # Вызываем API
        try:
            response = await self.client.chat.completions.create(
                model=self.config.model,
                messages=[
                    {"role": "system", "content": self.SYSTEM_PROMPT},
                    {"role": "user", "content": content}
                ],
                max_tokens=8096,
                temperature=0.0
            )

            result = response.choices[0].message.content
            if result:
                result = result.strip()
            else:
                result = ""

            # Логируем ответ
            if self.vision_prompts_dir:
                self._log_vision_response(chunk_name, result)

            return result

        except Exception as e:
            raise ExtractionError(f"Ошибка Vision API: {e}")

    def _log_vision_prompt(self, chunk_name: str, page_idx: int, user_prompt: str):
        """Логирует промпт Vision API."""
        if not self.vision_prompts_dir:
            return

        filepath = self.vision_prompts_dir / f"chunk_{chunk_name}_prompt.txt"

        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"{'='*80}\n")
            f.write(f"VISION PROMPT - Чанк: {chunk_name}\n")
            f.write(f"{'='*80}\n\n")
            f.write(f"Модель: {self.config.model}\n")
            f.write(f"Страница (1-based): {page_idx + 1}\n")
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
    vision_config: Optional[VisionConfig] = None,
    chunk_filter: Optional[str] = None
) -> Dict[str, str]:
    """
    Утилитарная функция для быстрого извлечения.

    Args:
        pdf_path: Путь к PDF
        chunks_config_path: Путь к JSON с конфигурацией чанков
        vision_config: Конфигурация Vision API (опционально)
        chunk_filter: Имя чанка для отладки (опционально)

    Returns:
        Словарь {chunk_name: extracted_text}
    """
    from .config import load_chunks_config

    chunks = load_chunks_config(chunks_config_path)
    extractor = VisionExtractor(vision_config)

    return extractor.extract(pdf_path, chunks, chunk_filter)
