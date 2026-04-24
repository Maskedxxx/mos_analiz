# START_MODULE_CONTRACT
# PURPOSE: Клиенты к внешним сервисам/моделям, нужные PDF-парсеру. Три класса — `RemoteLayoutDetector` (HTTP Layout API), `LayoutDetector` (локальная Heron-101), `VLMClient` (async PaddleOCR-VL через vLLM). Все параметры — явные, через DI, без литералов.
# INPUTS: Значения из `PdfLayoutConfig` / `PdfVlmConfig`, PIL Image'ы, текстовые промпты.
# OUTPUTS: Списки регионов `[{class_name, bbox, score}, ...]` от layout-детекторов; словари `{text, tokens, time_sec}` от VLM.
# KEYWORDS: clients, layout-detector, vlm-client, remote-http, local-model, async-openai.
# LINKS: src/format_parsers/pdf/_config.py, src/format_parsers/pdf/parse.py.
# RATIONALE: Клиенты — это IO-слой PDF-парсера. Они не знают про doc_types или оркестрацию, принимают ровно свои параметры. Чистая логика (парсинг таблиц, reading order, нормализация) живёт отдельно в `_parsing.py`.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import asyncio
import base64
import io
import logging
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import requests
from PIL import Image

try:
    # torch нужен только для локального LayoutDetector. Если torch не установлен —
    # работа идёт исключительно через RemoteLayoutDetector по HTTP.
    import torch
except ImportError:
    torch = None

try:
    from openai import AsyncOpenAI, OpenAI
except ImportError:
    AsyncOpenAI = None  # type: ignore
    OpenAI = None  # type: ignore
# END_IMPORTS


# START_LOGGERS
# PURPOSE: Отдельные логгеры на каждый клиент для прозрачной трассировки.
layout_remote_logger = logging.getLogger(__name__ + ".layout_remote")
layout_local_logger = logging.getLogger(__name__ + ".layout_local")
vlm_client_logger = logging.getLogger(__name__ + ".vlm")
# END_LOGGERS


# START_REMOTE_LAYOUT_DETECTOR
# PURPOSE: HTTP-клиент к remote Layout API сервису.
# INPUTS: base_url сервиса, timeout; затем изображения для detect().
# OUTPUTS: Список регионов `[{class_name, bbox, score}, ...]`.
# KEYWORDS: remote, http, layout-api, lightweight.
class RemoteLayoutDetector:
    """
    Назначение:
        Удалённый layout-детектор через HTTP API. Используется, когда `PdfLayoutConfig.base_url`
        задан. По контракту эквивалентен локальному `LayoutDetector.detect()`.

    Вход:
        base_url: URL layout-сервиса.
        timeout: Таймаут запроса в секундах.

    Выход:
        Объект с методом `detect(image) -> List[{class_name, bbox, score}]`.

    Логика:
        Отправляет PNG-изображение на `POST {base_url}/detect` и возвращает массив
        регионов из ответа. Никаких локальных моделей не загружает — используется, когда
        торч/ML-стек разворачивать нецелесообразно.
    """

    def __init__(self, base_url: str, timeout: int):
        """
        Назначение:
            Инициализирует клиент.

        Вход:
            base_url: URL layout-сервиса (например, http://172.16.10.35:11439).
            timeout: Таймаут запроса в секундах.

        Выход:
            None.

        Логика:
            1. Нормализует URL (срезает хвостовой `/`).
            2. Сохраняет timeout для последующих запросов.
            3. Логирует инициализацию.
        """
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        layout_remote_logger.info(f"RemoteLayoutDetector: {self.base_url}")

    def detect(self, image: Union[str, Path, Image.Image]) -> List[Dict[str, Any]]:
        """
        Назначение:
            Отправляет изображение на удалённый сервис и возвращает обнаруженные регионы.

        Вход:
            image: Путь к файлу, `Path`, либо PIL Image.

        Выход:
            Список словарей `[{class_name, bbox, score}, ...]`, отсортирован по `score`.

        Логика:
            1. Приводит вход к RGB PIL Image.
            2. Сериализует в PNG в памяти.
            3. POST-запросом отправляет на `/detect` сервиса.
            4. При ошибке сети — бросает `ConnectionError` с контекстом URL.
        """
        if isinstance(image, (str, Path)):
            image = Image.open(image).convert("RGB")
        elif image.mode != "RGB":
            image = image.convert("RGB")
        buf = io.BytesIO()
        image.save(buf, format="PNG")
        buf.seek(0)
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
            layout_remote_logger.info(f"Remote layout: {len(regions)} регионов за {time_ms}ms")
            return regions
        except requests.exceptions.RequestException as e:
            layout_remote_logger.error(f"Layout API ошибка: {e}")
            raise ConnectionError(f"Layout Detection API недоступен: {self.base_url} — {e}") from e

    def visualize(self, image: Union[str, Path, Image.Image], regions: List[Dict[str, Any]], output_path: str) -> None:
        """
        Назначение:
            Рисует bbox регионов на изображении и сохраняет PNG — для отладки/логирования.

        Вход:
            image: Исходное изображение (путь или PIL Image).
            regions: Список регионов с полем `bbox`.
            output_path: Куда сохранить результат.

        Выход:
            None.

        Логика:
            1. Загружает изображение как RGB.
            2. Рисует каждый bbox цветом, зависящим от `class_name`.
            3. Подписывает индекс, класс и score.
            4. Сохраняет на диск.
        """
        from PIL import ImageDraw
        if isinstance(image, (str, Path)):
            img = Image.open(image).convert("RGB")
        else:
            img = image.copy().convert("RGB")
        draw = ImageDraw.Draw(img)
        # Палитра цветов выровнена с локальным LayoutDetector.visualize — одни и те же
        # классы подсвечены одним цветом, визуальные логи сравнимы между remote/local.
        colors = {"Text": "blue", "Title": "red", "Section-header": "darkred", "Table": "green", "Picture": "purple", "Caption": "pink"}
        for i, region in enumerate(regions):
            bbox = region["bbox"]
            cls = region["class_name"]
            color = colors.get(cls, "yellow")
            draw.rectangle(bbox, outline=color, width=3)
            draw.text((bbox[0] + 2, bbox[1] + 2), f"[{i}] {cls} {region['score']:.2f}", fill=color)
        img.save(output_path)
# END_REMOTE_LAYOUT_DETECTOR


# START_LOCAL_LAYOUT_DETECTOR
# PURPOSE: Локальный layout-детектор на базе Docling Heron-101 (RT-DETRv2).
# INPUTS: model_repo, device, пороги IoU/containment/score; затем изображения для detect().
# OUTPUTS: Список регионов `[{class_name, bbox, score}, ...]` после пост-фильтрации.
# KEYWORDS: local, torch, transformers, rt-detr, heron-101.
class LayoutDetector:
    """
    Назначение:
        Локальный layout-детектор на Docling Heron-101. Используется, когда
        `PdfLayoutConfig.base_url` не задан (remote недоступен).

    Вход:
        model_repo: HF repo_id модели.
        device: Устройство инференса (cuda:0 / cpu).
        conf_threshold, post_conf_min, dedup_iou, containment_thr: пороги алгоритма.

    Выход:
        Объект с методом `detect(image) -> List[{class_name, bbox, score}]`.

    Логика:
        Загружает модель и процессор через transformers; каждый `detect` прогоняет
        изображение через модель, пост-фильтрует результат от слабых/дублированных/
        вложенных регионов и возвращает итог.
    """

    def __init__(
        self,
        model_repo: str,
        device: str,
        conf_threshold: float,
        post_conf_min: float,
        dedup_iou: float,
        containment_thr: float,
    ):
        """
        Назначение:
            Загружает модель layout-детектора и сохраняет пороги алгоритма.

        Вход:
            model_repo: HF repo_id (например, docling-project/docling-layout-heron-101).
            device: Устройство инференса.
            conf_threshold: Порог уверенности при первичной post-process-выборке.
            post_conf_min: Порог пост-фильтрации (слабые регионы выкидываются).
            dedup_iou: Порог IoU для дедупликации пересекающихся регионов.
            containment_thr: Порог containment для удаления вложенных регионов.

        Выход:
            None.

        Логика:
            1. Проверяет, что torch установлен (иначе — `ImportError`).
            2. Загружает процессор и модель с весами fp16.
            3. Переводит модель в eval-режим.
            4. Сохраняет пороги для дальнейших вызовов `detect`.
        """
        if torch is None:
            raise ImportError("torch не установлен — используйте RemoteLayoutDetector через layout.base_url")
        from transformers import RTDetrImageProcessor, RTDetrV2ForObjectDetection

        layout_local_logger.info(f"Загрузка Docling Heron-101: {model_repo} (device={device})")
        self.processor = RTDetrImageProcessor.from_pretrained(model_repo)
        torch.cuda.empty_cache()
        self.device = device
        self.model = RTDetrV2ForObjectDetection.from_pretrained(model_repo, torch_dtype=torch.float16).to(device)
        self.model.eval()
        self.id2label = self.model.config.id2label
        # Сохраняем пороги — все операции пост-фильтрации берут их отсюда.
        self.conf_threshold = conf_threshold
        self.post_conf_min = post_conf_min
        self.dedup_iou = dedup_iou
        self.containment_thr = containment_thr
        layout_local_logger.info(f"Docling Heron-101 загружена, классов: {len(self.id2label)}")

    def detect(self, image: Union[str, Path, Image.Image]) -> List[Dict[str, Any]]:
        """
        Назначение:
            Детектирует layout-регионы на изображении.

        Вход:
            image: Путь к файлу или PIL Image.

        Выход:
            Отфильтрованный список `[{class_name, bbox, score}, ...]` по убыванию score.

        Логика:
            1. Приводит вход к RGB.
            2. Прогоняет через процессор/модель, post_process с `conf_threshold`.
            3. Применяет `_post_filter` (score-дроп → dedup по IoU → containment NMS).
        """
        if isinstance(image, (str, Path)):
            image = Image.open(image).convert("RGB")
        elif image.mode != "RGB":
            image = image.convert("RGB")
        inputs = self.processor(images=[image], return_tensors="pt")
        model_dtype = next(self.model.parameters()).dtype
        inputs = {k: v.to(device=self.device, dtype=model_dtype if v.is_floating_point() else None) for k, v in inputs.items()}
        with torch.no_grad():
            outputs = self.model(**inputs)
        target_sizes = torch.tensor([image.size[::-1]], device=self.device)
        results = self.processor.post_process_object_detection(outputs, target_sizes=target_sizes, threshold=self.conf_threshold)[0]
        regions: List[Dict[str, Any]] = []
        for score, label_id, box in zip(results["scores"], results["labels"], results["boxes"]):
            class_name = self.id2label[label_id.item()]
            regions.append({"class_name": class_name, "bbox": box.tolist(), "score": score.item()})
        layout_local_logger.info(f"Детекция: {len(regions)} регионов (до фильтрации)")
        regions = self._post_filter(regions)
        layout_local_logger.info(f"После фильтрации: {len(regions)} регионов")
        return regions

    def _post_filter(self, regions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Назначение:
            Пост-фильтрация: выкидывает слабые, дедупликация по IoU, containment NMS.

        Вход:
            regions: Сырые детекции от модели.

        Выход:
            Отфильтрованный список.

        Логика:
            1. Убирает регионы со `score < post_conf_min`.
            2. Сортирует по score убыв., оставляет один из каждой пары с IoU > dedup_iou.
            3. Удаляет регионы, вложенные в уже принятые (containment > containment_thr).
        """
        regions = [r for r in regions if r["score"] >= self.post_conf_min]
        regions.sort(key=lambda r: r["score"], reverse=True)
        kept: List[Dict[str, Any]] = []
        for region in regions:
            if any(self._iou(region["bbox"], e["bbox"]) > self.dedup_iou for e in kept):
                continue
            kept.append(region)
        kept.sort(key=lambda r: r["score"], reverse=True)
        final: List[Dict[str, Any]] = []
        for region in kept:
            if any(self._containment(region["bbox"], e["bbox"]) > self.containment_thr for e in final):
                continue
            final.append(region)
        return final

    @staticmethod
    def _iou(box1: list, box2: list) -> float:
        """
        Назначение:
            Вычисляет IoU двух bbox в формате `[x1, y1, x2, y2]`.

        Вход:
            box1, box2: Координаты прямоугольников.

        Выход:
            float в диапазоне [0.0, 1.0].

        Логика:
            Стандартная формула: пересечение / объединение.
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
        Назначение:
            Вычисляет долю пересечения относительно меньшего из двух bbox.

        Вход:
            box1, box2: Координаты прямоугольников `[x1, y1, x2, y2]`.

        Выход:
            float в диапазоне [0.0, 1.0] — насколько меньший bbox «внутри» большего.

        Логика:
            intersection / min(area1, area2). Используется для удаления вложенных
            дублирующих регионов (Text vs List-item и т.п.).
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

    def visualize(self, image: Union[str, Path, Image.Image], regions: List[Dict[str, Any]], output_path: str) -> None:
        """
        Назначение:
            Рисует bbox регионов на изображении и сохраняет PNG — для отладки.

        Вход:
            image: Исходное изображение.
            regions: Список регионов с bbox.
            output_path: Куда сохранить.

        Выход:
            None.

        Логика:
            Те же шаги, что в RemoteLayoutDetector.visualize, но палитра классов шире —
            локальная модель возвращает больше уникальных классов.
        """
        from PIL import ImageDraw
        if isinstance(image, (str, Path)):
            img = Image.open(image).convert("RGB")
        else:
            img = image.copy().convert("RGB")
        draw = ImageDraw.Draw(img)
        colors = {
            "Text": "blue", "Title": "red", "Section-header": "darkred",
            "Table": "green", "Picture": "purple", "Formula": "orange",
            "Caption": "pink", "Footnote": "brown",
            "Page-header": "gray", "Page-footer": "gray",
            "List-item": "cyan", "Document Index": "teal", "Code": "darkblue",
            "Checkbox-Selected": "lime", "Checkbox-Unselected": "olive",
            "Form": "magenta", "Key-Value Region": "gold",
        }
        for i, region in enumerate(regions):
            bbox = region["bbox"]
            cls = region["class_name"]
            score = region["score"]
            color = colors.get(cls, "yellow")
            draw.rectangle(bbox, outline=color, width=3)
            draw.text((bbox[0] + 2, bbox[1] + 2), f"[{i}] {cls} {score:.2f}", fill=color)
        img.save(output_path)
        layout_local_logger.info(f"Визуализация сохранена: {output_path}")
# END_LOCAL_LAYOUT_DETECTOR


# START_VLM_CLIENT
# PURPOSE: Async-клиент к PaddleOCR-VL через vLLM-совместимый сервис.
# INPUTS: base_url, model (optional — автоопределение), api_key, max_tokens, temperature. Потом — изображения + промпты.
# OUTPUTS: Словарь `{text, tokens, time_sec, [index]}` для каждого запроса.
# KEYWORDS: vlm, async, paddle-ocr-vl, openai-compatible, batch.
class VLMClient:
    """
    Назначение:
        Асинхронный клиент к VLM-серверу для распознавания текстовых регионов на
        изображениях. Совместим с OpenAI-протоколом (vLLM).

    Вход:
        base_url: URL сервиса.
        model: Имя модели или `None` (для автоопределения через `/v1/models`).
        api_key: API-ключ (для vLLM обычно любое непустое значение).
        max_tokens: Максимум токенов в ответе.
        temperature: Температура генерации.

    Выход:
        Объект с async-методами `recognize(image, prompt)` и `recognize_batch(items)`.

    Логика:
        Использует AsyncOpenAI SDK для отправки multimodal-запросов (image_url + text).
        Batch-метод запускает задачи параллельно через `asyncio.gather` и сохраняет
        индекс для корреляции ответов с входными элементами.
    """

    def __init__(
        self,
        base_url: str,
        model: Optional[str],
        api_key: str,
        max_tokens: int,
        temperature: float,
    ):
        """
        Назначение:
            Инициализирует VLM-клиент.

        Вход:
            base_url, model, api_key, max_tokens, temperature: см. описание класса.

        Выход:
            None.

        Логика:
            1. Проверяет, что openai SDK установлен.
            2. Сохраняет параметры.
            3. Создаёт AsyncOpenAI-клиент.
            4. Если model не задана — автоопределяет через sync-клиент.
        """
        if AsyncOpenAI is None:
            raise ImportError("openai SDK не установлен — `pip install openai`")
        self.base_url = base_url
        self.api_key = api_key
        self.max_tokens = max_tokens
        self.temperature = temperature
        self.client = AsyncOpenAI(base_url=self.base_url, api_key=self.api_key)
        self.model_name = model or self._detect_model()

    def _detect_model(self) -> str:
        """
        Назначение:
            Автоопределение имени модели через `/v1/models` сервиса.

        Вход:
            —

        Выход:
            str: ID первой модели, которую отдал сервис.

        Логика:
            1. Создаёт синхронный OpenAI-клиент для одного запроса.
            2. Запрашивает список моделей.
            3. Возвращает `models.data[0].id` или бросает ConnectionError.
        """
        sync_client = OpenAI(base_url=self.base_url, api_key=self.api_key)
        try:
            models = sync_client.models.list()
            if models.data:
                model_id = models.data[0].id
                vlm_client_logger.info(f"VLM модель (авто): {model_id}")
                return model_id
            raise ConnectionError("vLLM не вернул ни одной модели")
        except Exception as e:
            raise ConnectionError(f"Не удалось подключиться к VLM ({self.base_url}): {e}")

    async def recognize(
        self,
        image: Image.Image,
        prompt: str,
        max_tokens: Optional[int] = None,
    ) -> Dict[str, Any]:
        """
        Назначение:
            Отправляет изображение + текстовый промпт в VLM и возвращает распознанный текст.

        Вход:
            image: PIL Image (обычно — crop региона).
            prompt: Текстовый промпт (например, `OCR:`, `Table Recognition:`).
            max_tokens: Если задан — переопределяет `self.max_tokens` для этого вызова.
                Нужно, когда у одного клиента несколько use-case'ов с разным масштабом
                ответа (например, docx header OCR — 256 токенов, pdf full-page OCR — 2000).

        Выход:
            Dict `{text, tokens, time_sec}`.

        Логика:
            1. Сериализует изображение в base64 (PNG).
            2. Формирует multimodal-сообщение (image_url + text).
            3. Await на chat.completions.create.
            4. Возвращает текст, количество токенов и время ответа.
        """
        b64 = self._image_to_base64(image)
        content = [
            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            {"type": "text", "text": prompt},
        ]
        t0 = time.time()
        response = await self.client.chat.completions.create(
            model=self.model_name,
            messages=[{"role": "user", "content": content}],
            max_tokens=max_tokens if max_tokens is not None else self.max_tokens,
            temperature=self.temperature,
        )
        elapsed = time.time() - t0
        text = response.choices[0].message.content or ""
        tokens = response.usage.completion_tokens if response.usage else 0
        if elapsed > 0:
            vlm_client_logger.info(f"VLM ответ: {tokens} tok за {elapsed:.1f}s ({tokens / elapsed:.0f} tok/s)")
        return {"text": text.strip(), "tokens": tokens, "time_sec": round(elapsed, 2)}

    def recognize_sync(
        self,
        image: Image.Image,
        prompt: str,
        max_tokens: Optional[int] = None,
    ) -> str:
        """
        Назначение:
            Синхронная обёртка над `recognize` для кода, не поднимающего event loop
            (например, docx-парсер делает 1–6 OCR-вызовов на шапку документа).

        Вход:
            image: PIL Image.
            prompt: Текстовый промпт.
            max_tokens: Опциональный override лимита токенов для этого вызова.

        Выход:
            Текст распознавания (strip).

        Логика:
            Оборачивает `asyncio.run(self.recognize(...))` и возвращает только поле `text`.
            Блокирует поток на время одного запроса — это приемлемо для точечных sync-вызовов.
            Для батч-работы используйте `recognize_batch`.
        """
        result = asyncio.run(self.recognize(image, prompt, max_tokens=max_tokens))
        return result["text"]

    async def recognize_batch(self, items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Назначение:
            Параллельно распознаёт несколько регионов, сохраняя их индексы.

        Вход:
            items: Список `[{image, prompt, index}, ...]`.

        Выход:
            Список `[{text, tokens, time_sec, index}, ...]`.

        Логика:
            1. Запускает `_recognize_with_index` для каждого элемента через asyncio.gather.
            2. Исключения конвертирует в записи с `text=[ОШИБКА: ...]`, чтобы пайплайн не падал.
        """
        tasks = [self._recognize_with_index(item) for item in items]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        output: List[Dict[str, Any]] = []
        for i, result in enumerate(results):
            if isinstance(result, Exception):
                vlm_client_logger.error(f"Ошибка VLM для региона {items[i].get('index', i)}: {result}")
                output.append({"text": f"[ОШИБКА: {result}]", "tokens": 0, "time_sec": 0, "index": items[i].get("index", i)})
            else:
                output.append(result)
        return output

    async def _recognize_with_index(self, item: Dict[str, Any]) -> Dict[str, Any]:
        """
        Назначение:
            Обёртка над `recognize`, которая переносит поле `index` в результат.

        Вход:
            item: Dict с ключами `image`, `prompt`, `index`.

        Выход:
            Dict от `recognize` с добавленным ключом `index`.

        Логика:
            Простое делегирование + перенос index.
        """
        result = await self.recognize(item["image"], item["prompt"])
        result["index"] = item.get("index", 0)
        return result

    @staticmethod
    def _image_to_base64(image: Image.Image) -> str:
        """
        Назначение:
            Сериализует PIL Image в base64-строку (PNG).

        Вход:
            image: PIL Image.

        Выход:
            Base64-строка без префикса data:...

        Логика:
            Сохраняет в in-memory BytesIO, затем base64-кодирует.
        """
        buf = io.BytesIO()
        image.save(buf, format="PNG")
        buf.seek(0)
        return base64.b64encode(buf.read()).decode("utf-8")
# END_VLM_CLIENT
