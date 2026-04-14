#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Обёртка над Docling Heron-101 (RT-DETRv2) для детекции layout документов.

Вход: изображение страницы (PIL Image или путь).
Выход: список регионов [{class_name, bbox, score}], отсортированных по score.
"""

import logging
from pathlib import Path
from typing import List, Dict, Any, Optional, Union

try:
    import torch  # требуется только для локального детектора
except ImportError:
    torch = None
from PIL import Image

from . import config

logger = logging.getLogger(__name__)


class LayoutDetector:
    """
    Детектор layout документов на базе Docling Heron-101 (RT-DETRv2).

    Args:
        model_name: HF repo_id модели. По умолчанию — из config.
        device: устройство для инференса ("cuda:0" / "cpu").
    """

    def __init__(
        self,
        model_name: Optional[str] = None,
        device: str = "cuda:0"
    ):
        if torch is None:
            raise ImportError("torch не установлен — используйте RemoteLayoutDetector через layout_base_url")
        from transformers import RTDetrV2ForObjectDetection, RTDetrImageProcessor

        repo = model_name or config.LAYOUT_MODEL_REPO
        logger.info(f"Загрузка Docling Heron-101: {repo} (device={device})")

        self.processor = RTDetrImageProcessor.from_pretrained(repo)
        torch.cuda.empty_cache()
        self.device = device
        self.model = RTDetrV2ForObjectDetection.from_pretrained(
            repo, torch_dtype=torch.float16,
        ).to(device)
        self.model.eval()

        # Маппинг id → class_name из модели
        self.id2label = self.model.config.id2label
        logger.info(f"Docling Heron-101 загружена, классов: {len(self.id2label)}")
        logger.info(f"Классы: {self.id2label}")

    def detect(self, image: Union[str, Path, Image.Image]) -> List[Dict[str, Any]]:
        """
        Детектирует layout-элементы на изображении.

        Args:
            image: путь к файлу или PIL Image.

        Returns:
            Список словарей:
            [{"class_name": str, "bbox": [x1,y1,x2,y2], "score": float}, ...]
            Отсортировано по score desc.
        """
        # Загрузка изображения если путь
        if isinstance(image, (str, Path)):
            image = Image.open(image).convert("RGB")
        elif image.mode != "RGB":
            image = image.convert("RGB")

        # Препроцессинг (приводим к dtype модели — float16 для экономии памяти)
        inputs = self.processor(images=[image], return_tensors="pt")
        model_dtype = next(self.model.parameters()).dtype
        inputs = {
            k: v.to(device=self.device, dtype=model_dtype if v.is_floating_point() else None)
            for k, v in inputs.items()
        }

        # Инференс
        with torch.no_grad():
            outputs = self.model(**inputs)

        # Постобработка — получаем bbox в абсолютных координатах
        target_sizes = torch.tensor([image.size[::-1]], device=self.device)  # (h, w)
        results = self.processor.post_process_object_detection(
            outputs,
            target_sizes=target_sizes,
            threshold=config.LAYOUT_CONF_THRESHOLD,
        )[0]

        regions = []
        for score, label_id, box in zip(
            results["scores"], results["labels"], results["boxes"]
        ):
            class_name = self.id2label[label_id.item()]
            regions.append({
                "class_name": class_name,
                "bbox": box.tolist(),  # [x1, y1, x2, y2]
                "score": score.item(),
            })

        logger.info(f"Детекция: {len(regions)} регионов (до фильтрации)")

        # Пост-фильтрация
        regions = self._post_filter(regions)
        logger.info(f"После фильтрации: {len(regions)} регионов")

        return regions

    def _post_filter(self, regions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Пост-фильтрация: убрать слабые, дедупликация по IoU и containment NMS.

        Containment NMS удаляет вложенные регионы: если маленький регион на >70%
        внутри большого (или наоборот), убираем тот, у которого ниже score.
        Это решает проблему, когда модель детектирует и крупный Text-блок,
        и отдельные List-item/Section-header внутри него.

        Args:
            regions: сырые детекции.

        Returns:
            Отфильтрованный список.
        """
        # 1. Убрать совсем слабые
        regions = [r for r in regions if r["score"] >= config.LAYOUT_POST_CONF_MIN]

        # 2. Дедупликация перекрывающихся bbox (оставляем с большим score)
        regions.sort(key=lambda r: r["score"], reverse=True)
        kept = []
        for region in regions:
            is_dup = False
            for existing in kept:
                if self._iou(region["bbox"], existing["bbox"]) > config.LAYOUT_DEDUP_IOU:
                    is_dup = True
                    break
            if not is_dup:
                kept.append(region)

        # 3. Containment NMS: удаляем вложенные регионы с меньшим score
        #    intersection / min(area_A, area_B) > threshold → убрать слабый
        kept.sort(key=lambda r: r["score"], reverse=True)
        final = []
        for region in kept:
            is_contained = False
            for existing in final:
                if self._containment(region["bbox"], existing["bbox"]) > config.LAYOUT_CONTAINMENT_THR:
                    is_contained = True
                    break
            if not is_contained:
                final.append(region)

        return final

    @staticmethod
    def _iou(box1: list, box2: list) -> float:
        """
        Вычисляет IoU двух bbox [x1,y1,x2,y2].

        Args:
            box1: первый bbox.
            box2: второй bbox.

        Returns:
            IoU значение (0.0 - 1.0).
        """
        x1 = max(box1[0], box2[0])
        y1 = max(box1[1], box2[1])
        x2 = min(box1[2], box2[2])
        y2 = min(box1[3], box2[3])

        inter = max(0, x2 - x1) * max(0, y2 - y1)
        area1 = (box1[2] - box1[0]) * (box1[3] - box1[1])
        area2 = (box2[2] - box2[0]) * (box2[3] - box2[1])
        union = area1 + area2 - inter

        return inter / union if union > 0 else 0.0

    @staticmethod
    def _containment(box1: list, box2: list) -> float:
        """
        Вычисляет степень вложенности: intersection / min(area1, area2).

        Высокое значение означает, что меньший bbox почти целиком внутри большего.
        Используется для удаления дублирующих вложенных регионов.

        Args:
            box1: первый bbox [x1,y1,x2,y2].
            box2: второй bbox [x1,y1,x2,y2].

        Returns:
            Containment ratio (0.0 - 1.0).
        """
        x1 = max(box1[0], box2[0])
        y1 = max(box1[1], box2[1])
        x2 = min(box1[2], box2[2])
        y2 = min(box1[3], box2[3])

        inter = max(0, x2 - x1) * max(0, y2 - y1)
        area1 = (box1[2] - box1[0]) * (box1[3] - box1[1])
        area2 = (box2[2] - box2[0]) * (box2[3] - box2[1])
        min_area = min(area1, area2)

        return inter / min_area if min_area > 0 else 0.0

    def visualize(
        self,
        image: Union[str, Path, Image.Image],
        regions: List[Dict[str, Any]],
        output_path: str
    ) -> None:
        """
        Рисует bbox на изображении и сохраняет результат.

        Args:
            image: исходное изображение.
            regions: список регионов с bbox.
            output_path: путь для сохранения аннотированного изображения.
        """
        from PIL import ImageDraw

        if isinstance(image, (str, Path)):
            img = Image.open(image).convert("RGB")
        else:
            img = image.copy().convert("RGB")

        draw = ImageDraw.Draw(img)

        # Цвета по классам Heron-101
        colors = {
            "Text": "blue",
            "Title": "red",
            "Section-header": "darkred",
            "Table": "green",
            "Picture": "purple",
            "Formula": "orange",
            "Caption": "pink",
            "Footnote": "brown",
            "Page-header": "gray",
            "Page-footer": "gray",
            "List-item": "cyan",
            "Document Index": "teal",
            "Code": "darkblue",
            "Checkbox-Selected": "lime",
            "Checkbox-Unselected": "olive",
            "Form": "magenta",
            "Key-Value Region": "gold",
        }

        for i, region in enumerate(regions):
            bbox = region["bbox"]
            cls = region["class_name"]
            score = region["score"]
            color = colors.get(cls, "yellow")

            draw.rectangle(bbox, outline=color, width=3)
            label = f"[{i}] {cls} {score:.2f}"
            draw.text((bbox[0] + 2, bbox[1] + 2), label, fill=color)

        img.save(output_path)
        logger.info(f"Визуализация сохранена: {output_path}")
