#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
VLM-клиент для PaddleOCR-VL-1.5 через OpenAI-совместимый API.

Отправляет crop изображения + промпт, получает текст.
Параметры задаются через конструктор (из AuditConfig).
"""

import asyncio
import base64
import io
import logging
import time
from typing import List, Dict, Any, Optional

from PIL import Image
from openai import AsyncOpenAI

from . import config as paddle_config

logger = logging.getLogger(__name__)


class VLMClient:
    """
    Клиент к VLM-серверу (PaddleOCR-VL-1.5 через vLLM).

    Args:
        base_url: URL vLLM-сервера.
        model_name: имя модели (None = автоопределение через /v1/models).
        api_key: API-ключ (vLLM обычно не требует).
        max_tokens: максимум токенов в ответе.
        temperature: температура генерации.
    """

    def __init__(
        self,
        base_url: str = "http://localhost:8010/v1",
        model_name: Optional[str] = None,
        api_key: Optional[str] = None,
        max_tokens: Optional[int] = None,
        temperature: Optional[float] = None,
    ):
        self.base_url = base_url
        self.max_tokens = max_tokens or paddle_config.VLM_MAX_TOKENS
        self.temperature = temperature if temperature is not None else paddle_config.VLM_TEMPERATURE
        self.client = AsyncOpenAI(
            base_url=self.base_url,
            api_key=api_key or paddle_config.VLM_API_KEY,
        )
        # Автоопределение модели если не указана
        self.model_name = model_name or self._detect_model()

    def _detect_model(self) -> str:
        """
        Автоопределение модели на vLLM через /v1/models.

        Returns:
            str: имя модели

        Raises:
            ConnectionError: если vLLM не отвечает
        """
        from openai import OpenAI
        sync_client = OpenAI(base_url=self.base_url, api_key=paddle_config.VLM_API_KEY)
        try:
            models = sync_client.models.list()
            if models.data:
                model_id = models.data[0].id
                logger.info(f"VLM модель (авто): {model_id}")
                return model_id
            raise ConnectionError("vLLM не вернул ни одной модели")
        except Exception as e:
            raise ConnectionError(f"Не удалось подключиться к VLM ({self.base_url}): {e}")

    async def recognize(self, image: Image.Image, prompt: str) -> Dict[str, Any]:
        """
        Отправляет изображение + промпт в VLM и возвращает текст.

        Args:
            image: PIL Image (crop региона).
            prompt: текстовый промпт (напр. "OCR:", "Table Recognition:").

        Returns:
            {"text": str, "tokens": int, "time_sec": float}
        """
        b64 = self._image_to_base64(image)

        content = [
            {
                "type": "image_url",
                "image_url": {"url": f"data:image/png;base64,{b64}"}
            },
            {
                "type": "text",
                "text": prompt,
            }
        ]

        t0 = time.time()

        response = await self.client.chat.completions.create(
            model=self.model_name,
            messages=[{"role": "user", "content": content}],
            max_tokens=self.max_tokens,
            temperature=self.temperature,
        )

        elapsed = time.time() - t0
        text = response.choices[0].message.content or ""
        tokens = response.usage.completion_tokens if response.usage else 0

        if elapsed > 0:
            logger.info(f"VLM ответ: {tokens} tok за {elapsed:.1f}s ({tokens/elapsed:.0f} tok/s)")

        return {
            "text": text.strip(),
            "tokens": tokens,
            "time_sec": round(elapsed, 2),
        }

    async def recognize_batch(
        self,
        items: List[Dict[str, Any]]
    ) -> List[Dict[str, Any]]:
        """
        Параллельное распознавание нескольких регионов.

        Args:
            items: список [{"image": PIL.Image, "prompt": str, "index": int}, ...]

        Returns:
            Список результатов [{"text": str, "tokens": int, "time_sec": float, "index": int}, ...]
        """
        tasks = []
        for item in items:
            tasks.append(self._recognize_with_index(item))

        results = await asyncio.gather(*tasks, return_exceptions=True)

        output = []
        for i, result in enumerate(results):
            if isinstance(result, Exception):
                logger.error(f"Ошибка VLM для региона {items[i].get('index', i)}: {result}")
                output.append({
                    "text": f"[ОШИБКА: {result}]",
                    "tokens": 0,
                    "time_sec": 0,
                    "index": items[i].get("index", i),
                })
            else:
                output.append(result)

        return output

    async def _recognize_with_index(self, item: Dict[str, Any]) -> Dict[str, Any]:
        """Обёртка для сохранения index в результате."""
        result = await self.recognize(item["image"], item["prompt"])
        result["index"] = item.get("index", 0)
        return result

    @staticmethod
    def _image_to_base64(image: Image.Image) -> str:
        """
        Конвертирует PIL Image в base64 строку.

        Args:
            image: PIL Image.

        Returns:
            Base64 строка.
        """
        buf = io.BytesIO()
        image.save(buf, format="PNG")
        buf.seek(0)
        return base64.b64encode(buf.read()).decode("utf-8")
