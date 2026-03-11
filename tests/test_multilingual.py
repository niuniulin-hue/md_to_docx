from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document

from convert import convert_file, read_markdown_text


MULTILINGUAL_TEXT = """# 多語言測試 / Multilingual Test

繁體中文：歡迎使用 Markdown 轉 DOCX。
日本語：こんにちは、世界。
Français : élève, cœur, façade.
Español: niño, acción, información.
Deutsch: Grüße für München, Straße, äußerst.
"""

TRADITIONAL_CHINESE_TEXT = "# 繁體中文\n\n歡迎使用 Markdown 轉 DOCX，保留格式。\n"
JAPANESE_TEXT = "# 日本語\n\nこんにちは、世界。マークダウンから DOCX へ変換します。\n"
WESTERN_TEXT = "# Langues européennes\n\nFrançais: élève, façade.\nEspañol: niño, acción.\nDeutsch: Grüße, Straße.\n"


class MultilingualEncodingTests(unittest.TestCase):
    def test_utf8_multilingual_round_trip(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "multilingual_utf8.md"
            out_path = Path(tmp_dir) / "multilingual_utf8.docx"
            md_path.write_text(MULTILINGUAL_TEXT, encoding="utf-8")

            decoded = convert_file(md_path, out_path)
            self.assertEqual(decoded.encoding.lower(), "utf-8")

            doc = Document(str(out_path))
            combined_text = "\n".join(p.text for p in doc.paragraphs)
            self.assertIn("繁體中文", combined_text)
            self.assertIn("こんにちは、世界", combined_text)
            self.assertIn("Français", combined_text)
            self.assertIn("Español", combined_text)
            self.assertIn("Grüße", combined_text)

    def test_utf8_bom_detection(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "multilingual_utf8_bom.md"
            raw = MULTILINGUAL_TEXT.encode("utf-8-sig")
            md_path.write_bytes(raw)

            decoded = read_markdown_text(md_path)
            self.assertEqual(decoded.encoding.lower(), "utf-8-sig")
            self.assertIn("日本語", decoded.text)

    def test_cp950_traditional_chinese_detection(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "traditional_zh.md"
            md_path.write_bytes(TRADITIONAL_CHINESE_TEXT.encode("cp950"))

            decoded = read_markdown_text(md_path)
            self.assertIn(decoded.encoding.lower(), {"cp950", "big5", "big5hkscs"})
            self.assertIn("繁體中文", decoded.text)

    def test_shift_jis_japanese_detection(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "japanese.md"
            md_path.write_bytes(JAPANESE_TEXT.encode("shift_jis"))

            decoded = read_markdown_text(md_path)
            self.assertIn(decoded.encoding.lower(), {"shift_jis", "cp932", "euc_jp"})
            self.assertIn("こんにちは", decoded.text)

    def test_cp1252_western_detection(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "western.md"
            md_path.write_bytes(WESTERN_TEXT.encode("cp1252"))

            decoded = read_markdown_text(md_path)
            self.assertIn(decoded.encoding.lower(), {"cp1252", "latin-1"})
            self.assertIn("Français", decoded.text)
            self.assertIn("Español", decoded.text)
            self.assertIn("Straße", decoded.text)

    def test_docx_contains_multilingual_font_mappings(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "multilingual_fonts.md"
            out_path = Path(tmp_dir) / "multilingual_fonts.docx"
            md_path.write_text(MULTILINGUAL_TEXT, encoding="utf-8")

            convert_file(md_path, out_path)

            with zipfile.ZipFile(out_path) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8")

            self.assertIn('w:eastAsia="Microsoft JhengHei"', document_xml)
            self.assertIn('w:eastAsia="Yu Gothic"', document_xml)
            self.assertIn('w:ascii="Calibri"', document_xml)


if __name__ == "__main__":
    unittest.main()
