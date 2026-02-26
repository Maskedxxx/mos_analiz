#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Алгоритм XY-cut для определения reading order.

Рекурсивно разбивает bbox по горизонтальным и вертикальным разрывам,
упорядочивая блоки сверху вниз, слева направо. Заимствован из MinerU.
"""

import logging
from typing import List, Dict, Any

logger = logging.getLogger(__name__)


def sort_by_reading_order(regions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Сортирует регионы по reading order (XY-cut).

    Args:
        regions: список [{"class_name": str, "bbox": [x1,y1,x2,y2], ...}, ...]

    Returns:
        Тот же список, отсортированный в порядке чтения.
    """
    if len(regions) <= 1:
        return regions

    indices = list(range(len(regions)))
    ordered_indices = _xycut(indices, regions)

    result = [regions[i] for i in ordered_indices]
    logger.info(f"Reading order: {len(result)} элементов отсортировано")
    return result


def _xycut(indices: List[int], regions: List[Dict[str, Any]]) -> List[int]:
    """
    Рекурсивный XY-cut.

    1. Проецирует bbox на Y-ось → ищет горизонтальные разрывы.
    2. Разбивает на горизонтальные полосы.
    3. Внутри каждой полосы проецирует на X-ось → ищет колонки.
    4. Рекурсивно повторяет.

    Args:
        indices: индексы регионов для сортировки.
        regions: все регионы (для доступа к bbox).

    Returns:
        Отсортированный список индексов.
    """
    if len(indices) <= 1:
        return indices

    # Получаем bbox для текущих индексов
    boxes = [(i, regions[i]["bbox"]) for i in indices]

    # Пробуем разделить по Y (горизонтальные полосы)
    y_groups = _split_axis(boxes, axis="y")
    if len(y_groups) > 1:
        result = []
        for group in y_groups:
            group_indices = [idx for idx, _ in group]
            result.extend(_xycut(group_indices, regions))
        return result

    # Пробуем разделить по X (колонки)
    x_groups = _split_axis(boxes, axis="x")
    if len(x_groups) > 1:
        result = []
        for group in x_groups:
            group_indices = [idx for idx, _ in group]
            result.extend(_xycut(group_indices, regions))
        return result

    # Не удалось разделить — сортируем по Y, затем по X
    boxes.sort(key=lambda b: (b[1][1], b[1][0]))
    return [idx for idx, _ in boxes]


def _split_axis(
    boxes: List[tuple],
    axis: str,
    gap_ratio: float = 0.02
) -> List[List[tuple]]:
    """
    Разбивает bbox по одной оси при наличии разрывов.

    Проецирует bbox на ось, ищет промежутки между проекциями.
    Порог разрыва = gap_ratio * общий диапазон оси.

    Args:
        boxes: список (index, [x1,y1,x2,y2]).
        axis: "x" или "y".
        gap_ratio: минимальный разрыв как доля от диапазона.

    Returns:
        Список групп. Каждая группа — список (index, bbox).
    """
    if axis == "y":
        # Проекция на Y: (y_start, y_end, box_tuple)
        projections = [(b[1][1], b[1][3], b) for b in boxes]
    else:
        # Проекция на X: (x_start, x_end, box_tuple)
        projections = [(b[1][0], b[1][2], b) for b in boxes]

    # Сортировка по началу проекции
    projections.sort(key=lambda p: p[0])

    # Общий диапазон
    total_min = min(p[0] for p in projections)
    total_max = max(p[1] for p in projections)
    total_range = total_max - total_min
    if total_range <= 0:
        return [boxes]

    min_gap = total_range * gap_ratio

    # Ищем разрывы между проекциями (merge overlapping)
    groups = []
    current_group = [projections[0][2]]
    current_end = projections[0][1]

    for i in range(1, len(projections)):
        start_i = projections[i][0]
        end_i = projections[i][1]

        if start_i - current_end > min_gap:
            # Разрыв — начинаем новую группу
            groups.append(current_group)
            current_group = [projections[i][2]]
            current_end = end_i
        else:
            current_group.append(projections[i][2])
            current_end = max(current_end, end_i)

    groups.append(current_group)
    return groups
