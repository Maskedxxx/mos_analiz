#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Smoke-тесты парсеров.

Проверяют что docx/pdf/pptx-парсеры не крашатся на реальных файлах.
Не проверяют качество распарсенного текста — только что функция возвращает dict.
"""
import os
from pathlib import Path

import pytest

from conftest import LOCAL_SMOKE_FIXTURES, PROJECT_ROOT


FIXTURES = {doc_type: str(path.relative_to(PROJECT_ROOT)) for doc_type, path in sorted(LOCAL_SMOKE_FIXTURES.items())}


PDF_DOC_TYPES = ("prikaz_pa", "otchet_rezultatov")


def _layout_api_available() -> bool:
    import urllib.request
    try:
        with urllib.request.urlopen("http://172.16.10.35:11439/health", timeout=2) as r:
            return r.status == 200
    except Exception:
        return False


def _resolve(rel_path: str) -> Path:
    return PROJECT_ROOT / rel_path


@pytest.mark.parametrize("doc_type,rel_path", sorted(FIXTURES.items()))
def test_parser_no_crash(doc_type, rel_path):
    """Парсер возвращает dict с обязательными полями без исключений."""
    file_path = _resolve(rel_path)
    if not file_path.exists():
        pytest.skip(f"Фикстура не найдена: {rel_path}")

    ext = file_path.suffix.lower()
    os.environ["NO_PROXY"] = "172.16.10.35,localhost,127.0.0.1"

    if ext == ".docx":
        from main import parse_docx
        # Без vlm_base_url — OCR картинок не делаем, только проверка что не крашнется
        result = parse_docx(str(file_path))
    elif ext == ".pptx":
        from main import parse_pptx
        result = parse_pptx(str(file_path))
    else:
        pytest.skip(f"Неподдерживаемое расширение: {ext}")

    assert isinstance(result, dict), "Парсер должен вернуть dict"
    assert "filename" in result, "Нет имени файла в результате"
    # Должен быть хотя бы один scope с текстом
    text_keys = [k for k, v in result.items() if isinstance(v, str) and k not in ("filename", "path") and v.strip()]
    assert len(text_keys) > 0, f"Ни одного непустого scope: {list(result.keys())}"


@pytest.mark.parametrize("doc_type", PDF_DOC_TYPES)
def test_pdf_parser_no_crash(doc_type, generated_pdf_fixtures):
    """PDF-парсер (paddle) — требует запущенный Layout API на Spark."""
    if not _layout_api_available():
        pytest.skip("Layout API (172.16.10.35:11439) недоступен")

    file_path = generated_pdf_fixtures[doc_type]

    os.environ["NO_PROXY"] = "172.16.10.35,localhost,127.0.0.1"

    # Только проверяем что PDF рендерится без крашей (pymupdf fix)
    # Полный OCR через VLM — в интеграционном тесте
    import fitz
    doc = fitz.open(str(file_path))
    assert len(doc) > 0, "PDF без страниц"
    # Рендерим первую страницу с ограничением размера
    page = doc[0]
    scale = min(1.0, 4000 / max(page.rect.width, page.rect.height))
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat)
    assert pix.width > 0 and pix.height > 0, "Пустой рендер страницы"
    doc.close()
