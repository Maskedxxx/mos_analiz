#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OCR Extractor — постраничное распознавание через Vision LLM (HunyuanOCR).

Конвертирует PDF в изображения и отправляет каждую уникальную страницу
в OCR-модель ОДИН раз. Результат — словарь {page_num: raw_text}.

Отличие от VisionExtractor:
- Единый промпт для всех страниц (без chunk-промптов)
- Дедупликация: каждая страница OCR'ится ровно один раз
- Работает через локальный vLLM (HunyuanOCR), не облачный GPT
"""

import asyncio
import base64
import io
import logging
import time
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import numpy as np
from PIL import Image
from openai import AsyncOpenAI

log = logging.getLogger(__name__)


class OcrExtractor:
    """
    Постраничный OCR через Vision LLM.

    Каждая уникальная страница обрабатывается ОДНИМ async-запросом к vLLM.
    Результат: {номер_страницы_1based: текст}.

    Attributes:
        model: имя модели на vLLM (автоопределение если None)
        base_url: URL vLLM API
        prompt: OCR-промпт (единый для всех страниц)
        dpi: разрешение рендеринга PDF -> PNG
        log_dir: директория для логов промптов/ответов
    """

    def __init__(
        self,
        base_url: str = "http://localhost:8000/v1/",
        model: Optional[str] = None,
        prompt: str = "",
        dpi: int = 200,
        log_dir: Optional[str] = None
    ):
        """
        Инициализация OCR-экстрактора.

        Args:
            base_url: URL vLLM API
            model: имя модели (None = автоопределение через /v1/models)
            prompt: OCR-промпт для всех страниц
            dpi: разрешение PDF -> PNG (200 = баланс качества и размера)
            log_dir: директория для логов
        """
        self.base_url = base_url
        self.prompt = prompt
        self.dpi = dpi
        self.log_dir = Path(log_dir) if log_dir else None

        # Async-клиент для OCR vLLM
        self.client = AsyncOpenAI(base_url=base_url, api_key="none")

        # Автоопределение модели если не указана
        self.model = model or self._detect_model()

        # Директория логов
        if self.log_dir:
            self.ocr_log_dir = self.log_dir / "ocr_pages"
            self.ocr_log_dir.mkdir(parents=True, exist_ok=True)
        else:
            self.ocr_log_dir = None

    def _detect_model(self) -> str:
        """
        Автоопределение модели на vLLM через /v1/models.

        Returns:
            str: имя модели

        Raises:
            ConnectionError: если vLLM не отвечает
        """
        from openai import OpenAI
        sync_client = OpenAI(base_url=self.base_url, api_key="none")
        try:
            models = sync_client.models.list()
            if models.data:
                model_id = models.data[0].id
                log.info(f"OCR модель (авто): {model_id}")
                return model_id
            raise ConnectionError("vLLM OCR не вернул ни одной модели")
        except Exception as e:
            raise ConnectionError(f"Не удалось подключиться к OCR vLLM ({self.base_url}): {e}")

    @staticmethod
    def _is_blank_page(image: Image.Image, threshold: float = 0.995) -> bool:
        """
        Проверяет пустая ли страница (почти полностью белая).

        LibreOffice при конвертации DOCX→PDF вставляет пустые страницы
        при смене ориентации (портретная↔альбомная). Такие страницы
        вызывают галлюцинации у OCR-моделей.

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
        from pdf2image import convert_from_path

        images = convert_from_path(pdf_path, dpi=72, fmt='PNG')
        total = len(images)
        blanks = set()
        for i, img in enumerate(images):
            if self._is_blank_page(img):
                blanks.add(i)
                log.info(f"  стр.{i+1}: пустая (пропускаем)")

        if blanks:
            log.info(f"Пустых страниц: {len(blanks)} из {total} (LibreOffice артефакты)")

        return total, blanks

    def extract_pages(
        self,
        pdf_path: str,
        page_indices: List[int]
    ) -> Dict[int, str]:
        """
        OCR указанных страниц PDF.

        Каждая уникальная страница из page_indices OCR'ится один раз.

        Args:
            pdf_path: путь к PDF файлу
            page_indices: список 0-based индексов страниц для OCR

        Returns:
            Dict[int, str]: {1-based номер страницы: OCR-текст}
        """
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF не найден: {pdf_path}")

        # PDF -> PNG[]
        images = self._pdf_to_images(pdf_path)
        total_pages = len(images)

        if total_pages == 0:
            raise ValueError("PDF не содержит страниц")

        # Фильтруем валидные индексы и дедуплицируем
        unique_indices = sorted(set(
            i for i in page_indices if 0 <= i < total_pages
        ))

        if not unique_indices:
            log.warning(f"Нет валидных страниц из {page_indices} (всего {total_pages})")
            return {}

        log.info(f"OCR: {len(unique_indices)} страниц из {total_pages}, модель={self.model}")

        # Async OCR всех уникальных страниц
        result = asyncio.run(self._ocr_pages_async(images, unique_indices))

        return result

    async def _ocr_pages_async(
        self,
        images: List[Image.Image],
        page_indices: List[int]
    ) -> Dict[int, str]:
        """
        Асинхронный OCR страниц (параллельно через asyncio.gather).

        Args:
            images: все изображения PDF
            page_indices: 0-based индексы страниц для OCR

        Returns:
            Dict[int, str]: {1-based page_num: text}
        """
        tasks = []
        for idx in page_indices:
            tasks.append(self._ocr_single_page(images[idx], idx))

        results = await asyncio.gather(*tasks, return_exceptions=True)

        page_texts = {}
        for idx, result in zip(page_indices, results):
            page_num = idx + 1  # 1-based
            if isinstance(result, Exception):
                log.error(f"OCR страница {page_num}: {result}")
                page_texts[page_num] = f"[ОШИБКА OCR: {result}]"
            else:
                page_texts[page_num] = result

        return page_texts

    async def _ocr_single_page(self, image: Image.Image, page_idx: int) -> str:
        """
        OCR одной страницы через Vision API vLLM.

        Args:
            image: PIL Image страницы
            page_idx: 0-based индекс страницы

        Returns:
            str: распознанный текст
        """
        page_num = page_idx + 1
        base64_image = self._image_to_base64(image)

        content = [
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/png;base64,{base64_image}",
                    "detail": "high"
                }
            },
            {
                "type": "text",
                "text": self.prompt
            }
        ]

        t0 = time.time()
        try:
            response = await self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": content}],
                max_tokens=8096,
                temperature=0.0
            )
            text = response.choices[0].message.content or ""
            text = text.strip()
            elapsed = time.time() - t0

            # Статистика
            tokens = getattr(response.usage, 'completion_tokens', 0) if response.usage else 0
            tps = tokens / elapsed if elapsed > 0 else 0
            log.info(f"  стр.{page_num}: {tokens} tok, {elapsed:.1f}s, {tps:.0f} tok/s")

            # Логируем ответ
            self._log_page(page_num, text)

            return text

        except Exception as e:
            elapsed = time.time() - t0
            log.error(f"  стр.{page_num}: ошибка за {elapsed:.1f}s — {e}")
            raise

    def _pdf_to_images(self, pdf_path: str) -> List[Image.Image]:
        """
        PDF -> список PNG-изображений.

        Args:
            pdf_path: путь к PDF

        Returns:
            List[Image.Image]: страницы как PIL Image
        """
        from pdf2image import convert_from_path

        images = convert_from_path(pdf_path, dpi=self.dpi, fmt='PNG')
        log.info(f"PDF -> {len(images)} страниц ({self.dpi} dpi)")
        return images

    def _image_to_base64(self, image: Image.Image) -> str:
        """PIL Image -> base64 строка."""
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        buffer.seek(0)
        return base64.b64encode(buffer.read()).decode("utf-8")

    def _log_page(self, page_num: int, text: str) -> None:
        """Сохраняет OCR-результат страницы в лог-файл."""
        if not self.ocr_log_dir:
            return
        filepath = self.ocr_log_dir / f"page_{page_num:03d}.txt"
        filepath.write_text(text, encoding="utf-8")
