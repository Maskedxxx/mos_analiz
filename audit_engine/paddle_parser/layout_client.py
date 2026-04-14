#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Удалённый клиент для Layout Detection API.

Замена локального LayoutDetector — отправляет PNG на удалённый сервис,
получает layout-регионы. Для серверов без GPU.

Интерфейс совпадает с LayoutDetector.detect() — drop-in замена.
"""

import io
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import requests
from PIL import Image

logger = logging.getLogger(__name__)


class RemoteLayoutDetector:
    """
    Удалённый layout detector через HTTP API.

    Совместим по интерфейсу с LayoutDetector — метод detect() принимает
    PIL Image и возвращает список регионов.

    Args:
        base_url: URL layout-сервиса (например http://172.16.10.35:8012)
        timeout: таймаут запроса в секундах
    """

    def __init__(self, base_url: str, timeout: int = 60):
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        logger.info(f"RemoteLayoutDetector: {self.base_url}")

    def detect(self, image: Union[str, Path, Image.Image]) -> List[Dict[str, Any]]:
        """
        Отправляет изображение на удалённый layout-сервис.

        Args:
            image: путь к файлу или PIL Image

        Returns:
            [{"class_name": str, "bbox": [x1,y1,x2,y2], "score": float}, ...]
        """
        # Загрузка если путь
        if isinstance(image, (str, Path)):
            image = Image.open(image).convert("RGB")
        elif image.mode != "RGB":
            image = image.convert("RGB")

        # PNG в байты
        buf = io.BytesIO()
        image.save(buf, format="PNG")
        buf.seek(0)

        # HTTP POST
        try:
            resp = requests.post(
                f"{self.base_url}/detect",
                files={"file": ("page.png", buf, "image/png")},
                timeout=self.timeout,
            )
            resp.raise_for_status()
            data = resp.json()
            regions = data.get("regions", [])
            time_ms = data.get("time_ms", 0)
            logger.info(f"Remote layout: {len(regions)} регионов за {time_ms}ms")
            return regions
        except requests.exceptions.RequestException as e:
            logger.error(f"Layout API ошибка: {e}")
            raise ConnectionError(f"Layout Detection API недоступен: {self.base_url} — {e}") from e

    def visualize(self, image, regions, output_path):
        """Визуализация — делегируем в PIL (без torch)."""
        from PIL import ImageDraw

        if isinstance(image, (str, Path)):
            img = Image.open(image).convert("RGB")
        else:
            img = image.copy().convert("RGB")

        draw = ImageDraw.Draw(img)
        colors = {
            "Text": "blue", "Title": "red", "Section-header": "darkred",
            "Table": "green", "Picture": "purple", "Caption": "pink",
        }
        for i, region in enumerate(regions):
            bbox = region["bbox"]
            cls = region["class_name"]
            color = colors.get(cls, "yellow")
            draw.rectangle(bbox, outline=color, width=3)
            draw.text((bbox[0]+2, bbox[1]+2), f"[{i}] {cls} {region['score']:.2f}", fill=color)

        img.save(output_path)
