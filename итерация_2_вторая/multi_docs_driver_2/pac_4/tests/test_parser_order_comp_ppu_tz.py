import tempfile
import unittest
from pathlib import Path

from docx import Document

from scripts.parser_order_comp_ppu_tz import parse_tz_document


def _build_doc(path: Path, paragraphs: list[str]) -> None:
    doc = Document()
    for text in paragraphs:
        doc.add_paragraph(text)
    doc.save(path)


class TzParserTest(unittest.TestCase):
    def test_empty_numbered_rules_returns_empty_list(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            doc_path = Path(tmpdir) / "tz.docx"
            _build_doc(doc_path, ["Заголовок без правил", "Еще строка"])

            data = parse_tz_document(doc_path)
            self.assertEqual(data["правила"], [])


if __name__ == "__main__":
    unittest.main()
