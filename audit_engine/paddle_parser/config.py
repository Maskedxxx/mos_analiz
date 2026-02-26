#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Конфигурация Paddle OCR pipeline: Docling Heron-101 + PaddleOCR-VL-1.5.

Константы layout-детекции, промпты по классам, markdown-обёртки.
VLM-параметры (base_url, model, api_key) задаются через AuditConfig.
"""

# ── Layout-модель (Docling Heron-101, RT-DETRv2) ──────────────
LAYOUT_MODEL_REPO = "docling-project/docling-layout-heron-101"
LAYOUT_CONF_THRESHOLD = 0.3      # RT-DETRv2 рекомендует выше, чем YOLO
LAYOUT_POST_CONF_MIN = 0.1       # пост-фильтрация: убрать слабые
LAYOUT_DEDUP_IOU = 0.9           # порог дедупликации перекрывающихся bbox
LAYOUT_CONTAINMENT_THR = 0.7     # порог вложенности: intersection/min_area

# ── Crop настройки ───────────────────────────────────────────
CROP_PADDING_PX = 15             # отступ при вырезке (белый фон)

# ── VLM дефолты (используются если AuditConfig не задаёт) ─────
VLM_API_KEY = "none"             # vLLM не требует ключ
VLM_MAX_TOKENS = 2000
VLM_TEMPERATURE = 0.0

# ── Промпты PaddleOCR-VL-1.5 по классам Heron-101 ────────────
# Docling Heron-101 классы (17 штук):
#   0: Caption, 1: Footnote, 2: Formula, 3: List-item,
#   4: Page-footer, 5: Page-header, 6: Picture, 7: Section-header,
#   8: Table, 9: Text, 10: Title, 11: Document Index,
#   12: Code, 13: Checkbox-Selected, 14: Checkbox-Unselected,
#   15: Form, 16: Key-Value Region
#
# PaddleOCR-VL-1.5 промпты:
#   "OCR:"                 — распознавание текста
#   "Table Recognition:"   — таблицы (структура + содержимое)
#   "Formula Recognition:" — математические формулы (LaTeX)

CLASS_PROMPTS = {
    "Text":                 "OCR:",
    "Title":                "OCR:",
    "Section-header":       "OCR:",
    "Caption":              "OCR:",
    "Footnote":             "OCR:",
    "List-item":            "OCR:",
    "Page-header":          "OCR:",
    "Page-footer":          "OCR:",
    "Document Index":       "OCR:",
    "Code":                 "OCR:",
    "Form":                 "OCR:",
    "Key-Value Region":     "OCR:",
    "Checkbox-Selected":    "OCR:",
    "Checkbox-Unselected":  "OCR:",
    "Table":                "Table Recognition:",
    "Formula":              "Formula Recognition:",
    "Picture":              "OCR:",
}

# Фильтрация мелких элементов по площади (доля от площади страницы)
SMALL_ELEMENT_MIN_AREA_RATIO = 0.001  # 0.1% от площади страницы

# Классы, которые фильтруются по размеру bbox (мелкие = мусор)
SMALL_ELEMENT_CLASSES = {"Page-header", "Page-footer"}

# Markdown-обёртки по типу элемента
CLASS_MD_WRAPPER = {
    "Text":                 "{text}",
    "Title":                "## {text}",
    "Section-header":       "### {text}",
    "Caption":              "*{text}*",
    "Footnote":             "*{text}*",
    "List-item":            "- {text}",
    "Page-header":          "{text}",
    "Page-footer":          "{text}",
    "Document Index":       "{text}",
    "Code":                 "```\n{text}\n```",
    "Form":                 "{text}",
    "Key-Value Region":     "{text}",
    "Checkbox-Selected":    "[x] {text}",
    "Checkbox-Unselected":  "[ ] {text}",
    "Table":                "{text}",
    "Formula":              "$$\n{text}\n$$",
    "Picture":              "{text}",
}

# Формат таблиц: "html" сохраняет colspan/rowspan (лучше для LLM-аудита)
TABLE_OUTPUT_FORMAT = "html"
