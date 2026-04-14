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

from conftest import DOC_CONFIGS, PROJECT_ROOT


# Фикстуры: один файл на тип, найденный в logs_result/*/original/
FIXTURES = {
    "prikaz_ppu":           "logs_result/prikaz_ppu/session_20260404_171205/original/3.10._Приказ о ППУ СпецТехРесурс НТ.docx",
    "prikaz_comp_ppu":      "logs_result/prikaz_comp_ppu/session_20260404_172642/original/3.10._Приказ о конкурсах ППУ СпецТехРесурс НТ.docx",
    "prikaz_vyhod":         "logs_result/prikaz_vyhod/session_20260410_072356/original/3.6._Приказ об проведении выхода ООО Гофромир.docx",
    "prikaz_tirazh":        "logs_result/prikaz_tirazh/session_20260329_182646/original/0_5_Приказ_о_переходе_на_этап_Тиражирования_с_приложениями_ООО_АПТОС.docx",
    "prikaz_ic_potoka_el":  "uploads/53680497/1.3._Приказ о создании ИЦ потока в эл.виде с прилож._МТЭР ЦТС.docx",
    "prikaz_ic_el":         "uploads/6dbbddcf/1.5._Приказ о создании ИЦ с прилож. МТЭР ЦТС эл. вид.docx",
    "polozhenie_comp_ppu":  "logs_result/polozhenie_comp_ppu/session_20260407_105903/original/3.10._Положение о конкурсах проектов и ППУ (1).docx",
    "polozhenie_ppu":       "logs_result/polozhenie_ppu/session_20260402_111335/original/Положение о ППУ.docx",
    "polozhenie_po":        "logs_result/polozhenie_po/session_20260403_104140/original/3.5._Положение о ПО _наименование предприятия_.docx",
    "akt_nachala":          "logs_result/akt_nachala/session_20260403_080809/original/0.1_Акт начала мероприятий _наименование предприятия_.docx",
    "cheklist_eu":          "logs_result/cheklist_eu/session_20260403_085256/original/2.5._Чек лист выбора ЭУ.docx",
    "presentation_eu":      "logs_result/presentation_eu/session_20260403_082452/original/2.5._Создание ЭУ наименование предприятия.pptx",
}


# PDF фикстуры требуют Layout API — отдельный тест с SKIP если API недоступен
PDF_FIXTURES = {
    "prikaz_pa":         "logs_result/prikaz_pa/session_20260326_123839/original/Приказ_О ведении производственного анализа_ООО НПП Лосев.pdf",
    "otchet_rezultatov": "logs_result/otchet_rezultatov/session_20260402_063814/original/Отчет о результатах вскрытия резервов ООО ОПТИМА ИМПОРТ утв...pdf",
}


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

    chunks_path = DOC_CONFIGS / doc_type / "chunks_vision.json"
    if not chunks_path.exists():
        pytest.skip(f"Нет chunks_vision.json для {doc_type}")

    ext = file_path.suffix.lower()
    os.environ["NO_PROXY"] = "172.16.10.35,localhost,127.0.0.1"

    if ext == ".docx":
        from audit_engine.docx_parser import parse_docx
        # Без vlm_base_url — OCR картинок не делаем, только проверка что не крашнется
        result = parse_docx(str(file_path), str(chunks_path))
    elif ext == ".pptx":
        from audit_engine.pptx_parser import parse_pptx
        result = parse_pptx(str(file_path), str(chunks_path))
    else:
        pytest.skip(f"Неподдерживаемое расширение: {ext}")

    assert isinstance(result, dict), "Парсер должен вернуть dict"
    assert "имя_файла" in result or "filename" in result, "Нет имени файла в результате"
    # Должен быть хотя бы один scope с текстом
    text_keys = [k for k, v in result.items() if isinstance(v, str) and k not in ("имя_файла", "путь", "filename", "path") and v.strip()]
    assert len(text_keys) > 0, f"Ни одного непустого scope: {list(result.keys())}"


@pytest.mark.parametrize("doc_type,rel_path", sorted(PDF_FIXTURES.items()))
def test_pdf_parser_no_crash(doc_type, rel_path):
    """PDF-парсер (paddle) — требует запущенный Layout API на Spark."""
    if not _layout_api_available():
        pytest.skip("Layout API (172.16.10.35:11439) недоступен")

    file_path = _resolve(rel_path)
    if not file_path.exists():
        pytest.skip(f"Фикстура не найдена: {rel_path}")

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
