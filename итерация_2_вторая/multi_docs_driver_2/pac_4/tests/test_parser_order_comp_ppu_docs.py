import tempfile
import unittest
from pathlib import Path

from docx import Document

from scripts.parser_order_comp_ppu_docs import parse_order_comp_ppu_version


def _build_doc(path: Path, paragraphs: list[str]) -> None:
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(path)


class ParserChunksTest(unittest.TestCase):
    def test_happy_path(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            doc_path = Path(tmpdir) / "doc.docx"
            _build_doc(
                doc_path,
                [
                    "ООО «Организация»",
                    "ПРИКАЗ № 11БП",
                    "г. Москва 13.11.2025 г.",
                    "О проведении ...",
                    "С целью обеспечения выполнения показателей",
                    "ПРИКАЗЫВАЮ:",
                    "1. Пункт первый.",
                    "Приказа оставляю за собой.",
                    "Генеральный директор",
                    "ООО «Организация» М.В. Ветров",
                ],
            )

            result = parse_order_comp_ppu_version(doc_path)
            self.assertIn("ООО «Организация»", result["чанк_шапка"])
            self.assertTrue(result["чанк_текст"].startswith("С целью обеспечения"))
            self.assertIn("Приказа оставляю за собой.", result["чанк_текст"])
            self.assertIn("Генеральный директор", result["чанк_мета"])

    def test_no_start_marker_entire_doc_in_header(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            doc_path = Path(tmpdir) / "doc.docx"
            _build_doc(doc_path, ["Строка 1", "Строка 2 без маркера"])

            result = parse_order_comp_ppu_version(doc_path)
            self.assertEqual(result["чанк_шапка"], "Строка 1\nСтрока 2 без маркера")
            self.assertEqual(result["чанк_текст"], "")
            self.assertEqual(result["чанк_мета"], "")

    def test_no_end_marker_footer_empty(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            doc_path = Path(tmpdir) / "doc.docx"
            _build_doc(
                doc_path,
                [
                    "Шапка",
                    "С целью обеспечения выполнения показателей",
                    "ПРИКАЗЫВАЮ:",
                    "1. Пункт первый.",
                ],
            )

            result = parse_order_comp_ppu_version(doc_path)
            self.assertEqual(result["чанк_шапка"], "Шапка")
            self.assertIn("ПРИКАЗЫВАЮ", result["чанк_текст"])
            self.assertEqual(result["чанк_мета"], "")

    def test_markers_case_insensitive(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            doc_path = Path(tmpdir) / "doc.docx"
            _build_doc(
                doc_path,
                [
                    "Шапка",
                    "С ЦЕЛЬЮ ОБЕСПЕЧЕНИЯ чего-либо",
                    "ПриказА оставляю за собой",
                    "Подвал",
                ],
            )

            result = parse_order_comp_ppu_version(doc_path)
            self.assertTrue(result["чанк_текст"].lower().startswith("с целью обеспечения"))
            self.assertEqual(result["чанк_мета"], "Подвал")

    def test_empty_document_returns_empty_chunks(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            doc_path = Path(tmpdir) / "doc.docx"
            _build_doc(doc_path, [])

            result = parse_order_comp_ppu_version(doc_path)
            self.assertEqual(result["чанк_шапка"], "")
            self.assertEqual(result["чанк_текст"], "")
            self.assertEqual(result["чанк_мета"], "")

    def test_alt_markers_are_detected(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            doc_path = Path(tmpdir) / "doc.docx"
            _build_doc(
                doc_path,
                [
                    "Шапка",
                    "В целях обеспечения стабильности процессов",  # вариант стартового маркера
                    "Основной текст",
                    "Контроль исполнения настоящего приказа оставляю за собой",  # вариант конечного маркера
                    "Подвал",
                ],
            )

            result = parse_order_comp_ppu_version(doc_path)
            self.assertEqual(result["чанк_шапка"], "Шапка")
            self.assertIn("Основной текст", result["чанк_текст"])
            self.assertIn("Подвал", result["чанк_мета"])


if __name__ == "__main__":
    unittest.main()
