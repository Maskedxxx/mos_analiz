#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Сборка чанков из постраничного OCR-текста.

Берёт результат OcrExtractor (Dict[page_num, text]) и конфигурацию чанков
из chunks_vision.json, собирает текст чанков по номерам страниц.
"""

import logging
from typing import Dict, List, Optional

from ..vision_parser.config import ChunkConfig
from .normalizer import normalize_ocr_text, strip_arrow_annotations

log = logging.getLogger(__name__)


class ChunkAssembler:
    """
    Собирает чанки документа из постраничного OCR-текста.

    Для каждого чанка: берёт pages из конфига -> собирает текст из page_texts ->
    нормализует (HTML-таблицы -> Markdown) -> возвращает {chunk_name: text}.
    """

    def assemble(
        self,
        page_texts: Dict[int, str],
        chunks: List[ChunkConfig],
        total_pages: int,
        strip_annotations_chunks: Optional[List[str]] = None,
        normalizer_fn=None
    ) -> Dict[str, str]:
        """
        Собирает чанки из постраничных OCR-текстов.

        Args:
            page_texts: {1-based page_num: OCR-текст} от OcrExtractor
            chunks: конфигурации чанков из chunks_vision.json
            total_pages: общее количество страниц в PDF
            strip_annotations_chunks: имена чанков, где убирать
                аннотации Word-форм (паттерн 'текст -> значение')
            normalizer_fn: функция нормализации текста (по умолчанию normalize_ocr_text).
                Для Paddle pipeline передаётся normalize_paddle_text,
                которая сохраняет HTML-таблицы с colspan/rowspan.

        Returns:
            Dict[str, str]: {chunk_name: собранный_и_нормализованный_текст}
        """
        _normalize = normalizer_fn or normalize_ocr_text
        _strip_chunks = set(strip_annotations_chunks or [])
        result = {}

        for chunk in chunks:
            # 0-based индексы -> 1-based номера страниц
            page_indices = chunk.get_page_indices(total_pages)
            page_nums = [idx + 1 for idx in page_indices]

            # Нужна ли очистка аннотаций для этого чанка
            need_strip = chunk.name in _strip_chunks

            # Собираем текст по страницам
            parts = []
            for pn in page_nums:
                if pn in page_texts:
                    text = page_texts[pn]
                    # Очистка аннотаций Word-форм только для указанных чанков
                    if need_strip:
                        text = strip_arrow_annotations(text)
                    if text and text.strip():
                        # Для многостраничных чанков — маркер страницы
                        if len(page_nums) > 1:
                            parts.append(f"[СТРАНИЦА {pn}]\n{text.strip()}")
                        else:
                            parts.append(text.strip())
                else:
                    log.warning(f"Чанк '{chunk.name}': страница {pn} отсутствует в OCR")

            if not parts:
                result[chunk.name] = ""
                log.warning(f"Чанк '{chunk.name}': пустой результат (страницы {page_nums})")
                continue

            # Склеиваем и нормализуем
            raw_text = "\n\n".join(parts)
            normalized = _normalize(raw_text)
            result[chunk.name] = normalized

            log.info(f"Чанк '{chunk.name}': страницы {page_nums}, {len(normalized)} символов")

        return result


def collect_unique_pages(chunks: List[ChunkConfig], total_pages: int) -> List[int]:
    """
    Собирает уникальные 0-based индексы страниц из ВСЕХ чанков.

    Используется для дедупликации: каждую страницу OCR'им один раз.

    Args:
        chunks: список конфигураций чанков
        total_pages: общее количество страниц

    Returns:
        List[int]: отсортированные уникальные 0-based индексы
    """
    all_indices = set()
    for chunk in chunks:
        indices = chunk.get_page_indices(total_pages)
        all_indices.update(indices)

    return sorted(all_indices)
