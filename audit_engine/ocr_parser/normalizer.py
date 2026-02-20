#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Нормализация OCR-текста.

Конвертация HTML-таблиц в Markdown и базовый cleanup текста
после распознавания OCR-моделью (HunyuanOCR и др.).
"""

import re


def html_table_to_markdown(html: str) -> str:
    """
    Конвертация HTML-таблицы в Markdown-формат.

    Args:
        html: строка с <table>...</table>

    Returns:
        str: таблица в формате | col1 | col2 |
    """
    rows = re.findall(r'<tr>(.*?)</tr>', html, re.DOTALL)
    md_rows = []
    for row in rows:
        # Извлекаем ячейки (td и th)
        cells = re.findall(r'<(?:td|th)[^>]*>(.*?)</(?:td|th)>', row, re.DOTALL)
        # Чистим HTML-теги внутри ячеек
        clean = [re.sub(r'<[^>]+>', ' ', c).strip() for c in cells]
        if clean:
            md_rows.append("| " + " | ".join(clean) + " |")

    if not md_rows:
        return html

    # Разделитель после первой строки (заголовок)
    result = [md_rows[0]]
    ncols = md_rows[0].count("|") - 1
    result.append("|" + "|".join(["---"] * ncols) + "|")
    result.extend(md_rows[1:])
    return "\n".join(result)


def strip_arrow_annotations(text: str) -> str:
    """
    Удаляет аннотации Word-форм вида 'описание поля -> значение'.

    Такие аннотации появляются в документах с формами, где поля имеют
    подписи типа 'Форма и наименование предприятия -> ООО "Ромашка"'.
    HunyuanOCR захватывает их как текст, что вызывает ложные срабатывания
    при сверке с шаблоном.

    Строка с ' -> ' заменяется на значение после стрелки.
    Строки без ' -> ' не затрагиваются.

    Args:
        text: OCR-текст страницы

    Returns:
        str: текст без аннотаций (значения сохраняются)
    """
    if not text:
        return ""

    lines = text.split('\n')
    result = []
    for line in lines:
        if ' -> ' in line:
            # Оставляем только значение после ->
            value = line.split(' -> ', 1)[1].strip()
            if value:
                result.append(value)
            # Пустое значение — пропускаем строку целиком
        else:
            result.append(line)
    return '\n'.join(result)


def unwrap_latex_artifacts(text: str) -> str:
    """
    Убирает LaTeX-обёртки из OCR-текста HunyuanOCR.

    HunyuanOCR оборачивает подчёркнутый/выделенный текст в LaTeX-формат:
    $ \\underline{\\mathrm{TEXT}} $ → TEXT

    Особые случаи:
    - \\mathrm{No} → № (знак номера)

    Безопасно для реальных документов: LaTeX-паттерны не встречаются
    в обычном тексте, только в артефактах OCR.

    Args:
        text: OCR-текст с возможными LaTeX-артефактами

    Returns:
        str: текст с развёрнутыми LaTeX-обёртками
    """
    if not text or '$' not in text:
        return text

    # $ \underline{\mathrm{TEXT}} $ → TEXT (с опциональными пробелами)
    def replace_latex(match):
        inner = match.group(1)
        # No → № (знак номера в русском)
        if inner == 'No':
            return '№'
        return inner

    text = re.sub(
        r'\$\s*\\underline\{\\mathrm\{([^}]*)\}\}\s*\$',
        replace_latex,
        text
    )

    return text


def normalize_ocr_text(text: str) -> str:
    """
    Нормализация OCR-текста: LaTeX-артефакты, HTML-таблицы -> Markdown, cleanup.

    Операции:
    1. Разворачивает LaTeX-обёртки HunyuanOCR (underline/mathrm)
    2. Заменяет <table>...</table> блоки на Markdown-таблицы
    3. Убирает лишние пустые строки
    4. Убирает пробелы в конце строк

    Args:
        text: OCR-текст (может содержать LaTeX и <table> от HunyuanOCR)

    Returns:
        str: нормализованный текст
    """
    if not text:
        return ""

    # LaTeX-артефакты -> чистый текст
    text = unwrap_latex_artifacts(text)

    # HTML-таблицы -> Markdown
    def replace_table(match):
        return html_table_to_markdown(match.group(0))

    text = re.sub(r'<table>.*?</table>', replace_table, text, flags=re.DOTALL)

    # Множественные пустые строки -> одна
    text = re.sub(r'\n{3,}', '\n\n', text)

    # Пробелы в конце строк
    text = re.sub(r' +\n', '\n', text)

    return text.strip()
