# START_MODULE_CONTRACT
# PURPOSE: Низкоуровневый парсер DOCX. Читает файл и отдаёт полный текст документа целиком плюс метаданные. Не знает ни про страницы, ни про правила, ни про LLM, ни про типы документов.
# INPUTS: Путь к DOCX, опциональные `VLMClient` и `DocxHeaderOcrConfig` — если заданы, парсер распознаёт картинки в шапке документа.
# OUTPUTS: Dict `{filename, path, raw_text}`.
# KEYWORDS: docx, python-docx, raw-text, full-document, header-ocr.
# LINKS: src/format_parsers/__init__.py, src/format_parsers/pdf/_clients.py::VLMClient, config/parsers.json::docx.header_ocr, main.py::_parse_document.
# RATIONALE: Весь DOCX-специфичный код живёт в одном файле. Сервис парсит формат, отдаёт сырой текст — нарезка на смысловые части происходит в верхних слоях. VLM-вызов для шапки делегируется общему `VLMClient` (один клиент для docx и pdf), конфигурация — в `config/parsers.json`.
# END_MODULE_CONTRACT

from __future__ import annotations

# START_IMPORTS
import io
import logging
from pathlib import Path
from typing import Any, List, Optional

from PIL import Image

from config.parsers import DocxHeaderOcrConfig
from src.format_parsers._types import ParsedDocument
from src.format_parsers.pdf._clients import VLMClient
# END_IMPORTS


# START_LOGGER
# PURPOSE: Отдельный логгер DOCX-парсера. Имя совпадает с именем модуля для прозрачной трассировки.
docx_parser_logger = logging.getLogger(__name__)
# END_LOGGER


# START_DOCX_PARSER
# PURPOSE: Прочитать DOCX и отдать полный текст документа вместе с метаданными.
# INPUTS: Путь к DOCX, опциональные `VLMClient` и `DocxHeaderOcrConfig` для header OCR.
# OUTPUTS: Dict `{filename, path, raw_text}`.
# KEYWORDS: docx, raw-text, full-document, public-api.
# RATIONALE: Сервис парсит формат — смысловая нарезка документа делается выше по конвейеру (doc_type / LLM).
def parse_docx(
    file_path: str,
    vlm: Optional[VLMClient] = None,
    header_ocr_cfg: Optional[DocxHeaderOcrConfig] = None,
) -> ParsedDocument:
    """
    Назначение:
        Читает DOCX-файл целиком и возвращает полный текст документа без деления на
        страницы или чанки.

    Вход:
        file_path: Путь к DOCX-файлу.
        vlm: Опциональный `VLMClient`. Если задан И переданы `header_ocr_cfg` — парсер
            распознаёт картинки (эмблему, штампы) в шапке документа. Без клиента OCR
            шапки пропускается.
        header_ocr_cfg: Параметры header OCR (промпт, фильтры, max_tokens). Обязателен
            в паре с `vlm`.

    Выход:
        dict: Словарь `{filename, path, raw_text}`.

    Логика:
        1. Открывает DOCX через python-docx.
        2. Если есть VLM+конфиг — OCR-ит картинки в шапке и ставит текст в самое начало.
        3. Извлекает все текстовые элементы и таблицы в правильном порядке.
        4. Склеивает всё в единый `raw_text` без разбиения по страницам.
    """
    try:
        from docx import Document
    except ImportError as exc:
        raise ModuleNotFoundError("python-docx is required for DOCX parsing") from exc

    file_path_obj = Path(file_path)
    doc = Document(str(file_path_obj))
    docx_parser_logger.info(
        "DOCX: %s, paragraphs=%s, tables=%s",
        file_path_obj.name,
        len(doc.paragraphs),
        len(doc.tables),
    )

    # Дополняем текст шапки OCR-результатом, если переданы клиент и конфиг.
    header_ocr_text = ""
    if vlm is not None and header_ocr_cfg is not None:
        header_ocr_text = _ocr_header_images(doc, vlm, header_ocr_cfg)
        if header_ocr_text:
            docx_parser_logger.info("  Header OCR chars: %s", len(header_ocr_text))

    # Сохраняем фактический порядок элементов документа: абзацы, текстбоксы, таблицы.
    elements = _extract_elements_in_order(doc)
    if header_ocr_text:
        elements.insert(0, header_ocr_text)

    # Склеиваем весь документ в один raw_text. Парсер не знает про страницы и правила —
    # он отдаёт всё содержимое как есть, а решение «что важно» принимают верхние слои.
    raw_text = "\n\n".join(elements)
    docx_parser_logger.info(
        "  Elements: %s, raw_text chars: %s", len(elements), len(raw_text)
    )

    return {
        "filename": file_path_obj.name,
        "path": str(file_path_obj.resolve()),
        "raw_text": raw_text,
    }
# END_DOCX_PARSER


# START_DOCX_HELPERS
# PURPOSE: Приватные хелперы парсинга DOCX, нужны только этому модулю.
# INPUTS: Объект Document python-docx, опциональный `VLMClient` + `DocxHeaderOcrConfig`.
# OUTPUTS: Распознанный текст шапки, упорядоченный список элементов и HTML таблицы.
# KEYWORDS: docx-internals, header-ocr, ordered-extraction, table-html.
def _ocr_header_images(doc: Any, vlm: VLMClient, cfg: DocxHeaderOcrConfig) -> str:
    """
    Назначение:
        Извлекает и OCR-ит изображения из начальных абзацев DOCX через общий VLMClient.

    Вход:
        doc: Загруженный объект DOCX.
        vlm: Инициализированный `VLMClient` (тот же, что используется PDF-парсером).
        cfg: Параметры OCR (промпт, фильтры размеров, окно абзацев, max_tokens).

    Выход:
        str: Объединённый текст, распознанный из картинок в шапке.

    Логика:
        1. Итерирует первые `cfg.max_paragraphs_scanned` абзацев.
        2. Находит embedded-картинки через DrawingML/Relationships namespaces.
        3. Отбрасывает мелкие картинки (мельче `cfg.min_image_width × cfg.min_image_height`).
        4. Отправляет каждую в VLM через `vlm.recognize_sync(image, cfg.prompt, cfg.max_tokens)`.
        5. Склеивает полученные тексты через \n.
    """
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ocr_parts: List[str] = []

    for paragraph in doc.paragraphs[: cfg.max_paragraphs_scanned]:
        # Смотрим только начало документа, потому что шапка обычно живёт именно там.
        blips = paragraph._element.findall(f".//{{{ns_a}}}blip")
        for blip in blips:
            rel_id = blip.get(f"{{{ns_r}}}embed")
            if not rel_id or rel_id not in doc.part.rels:
                continue
            try:
                blob = doc.part.rels[rel_id].target_part.blob
                image = Image.open(io.BytesIO(blob))
                # Отбрасываем мелкие декоративные картинки, чтобы не шуметь в OCR.
                if image.width < cfg.min_image_width or image.height < cfg.min_image_height:
                    continue
                text = vlm.recognize_sync(image, cfg.prompt, max_tokens=cfg.max_tokens)
                if text:
                    ocr_parts.append(text)
                    docx_parser_logger.info(
                        "  OCR image (%sx%s): %s",
                        image.width,
                        image.height,
                        text[:80],
                    )
            except Exception as exc:
                docx_parser_logger.warning("  OCR image failed: %s", exc)

    return "\n".join(ocr_parts)


def _extract_elements_in_order(doc: Any) -> List[str]:
    """
    Назначение:
        Извлекает элементы DOCX в реальном порядке их появления.

    Вход:
        doc: Загруженный объект DOCX.

    Выход:
        list[str]: Список текстовых элементов и HTML-таблиц в правильном порядке.

    Логика:
        1. Обходит XML body документа.
        2. Собирает обычные абзацы и тексты из текстбоксов.
        3. Конвертирует таблицы в HTML.
        4. Возвращает единый линейный список элементов.
    """
    from docx.oxml.ns import qn

    elements: List[str] = []
    body = doc.element.body
    table_elements = {table._element: table for table in doc.tables}

    for child in body:
        tag = child.tag
        if tag == qn("w:p"):
            # Собираем текст по run-узлам, чтобы не потерять дробление на уровне Word XML.
            runs_text: List[str] = []
            for run in child.findall(qn("w:r")):
                for text_node in run.findall(qn("w:t")):
                    if text_node.text:
                        runs_text.append(text_node.text)
            full_text = "".join(runs_text).strip()
            if full_text:
                elements.append(full_text)

            ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            ns_wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            # Отдельно вытаскиваем текст из текстбоксов, который не попадает в обычные paragraphs.
            for txbx_content in child.findall(f".//{{{ns_wps}}}txbx/{{{ns_w}}}txbxContent"):
                txbx_runs: List[str] = []
                for paragraph in txbx_content.findall(f"{{{ns_w}}}p"):
                    for run in paragraph.findall(f"{{{ns_w}}}r"):
                        for text_node in run.findall(f"{{{ns_w}}}t"):
                            if text_node.text:
                                txbx_runs.append(text_node.text)
                txbx_text = "".join(txbx_runs).strip()
                if txbx_text:
                    elements.append(txbx_text)
        elif tag == qn("w:tbl"):
            table = table_elements.get(child)
            if table:
                html = docx_parser__table_to_html(table)
                if html.strip():
                    elements.append(html)
    return elements


def docx_parser__table_to_html(table: Any) -> str:
    """
    Назначение:
        Преобразует таблицу DOCX в простой HTML.

    Вход:
        table: Объект таблицы из python-docx.

    Выход:
        str: HTML-представление таблицы.

    Логика:
        1. Проходит по строкам и ячейкам таблицы.
        2. Нормализует текст ячеек.
        3. Собирает итоговую HTML-таблицу.
    """
    rows_html: List[str] = []
    for row in table.rows:
        cells_html = []
        for cell in row.cells:
            cell_text = cell.text.strip().replace("\n", " ")
            cells_html.append(f"  <td>{cell_text}</td>")
        rows_html.append("<tr>\n" + "\n".join(cells_html) + "\n</tr>")
    if not rows_html:
        return ""
    return "<table>\n" + "\n".join(rows_html) + "\n</table>"
# END_DOCX_HELPERS
