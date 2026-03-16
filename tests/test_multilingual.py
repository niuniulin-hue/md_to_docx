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
ICON_MARKDOWN = """# Shipping ✅

- [x] Published
- [ ] Draft

Visit [Docs 🔗](https://example.com)

⚠️ Warning section

ℹ️ Info panel

👉 Start here

| Item | Notes |
|---|---|
| Tip | 💡 Keep it simple |

`Code ✅ keeps emoji`
"""
GENERIC_ICON_MARKDOWN = """# Routine 💪

🧘‍♀️ Reset

🏃‍♀️ Sprint

🚫 Avoid this

📦 Deliverables

🔄 Weekly reset

📊 Metrics

📅 Calendar

🩺 Health review

🍽️ Meal prep

☀️ Morning light

💖 Encouragement

`Code 🧘‍♀️ keeps emoji`
"""
AUTO_FALLBACK_ICON_MARKDOWN = """# Future glyphs 🦉

🧶 Yarn

🎓 Study

🇺🇸 Flag

1️⃣ First step

`Code 🦉 keeps emoji`
"""


class MultilingualEncodingTests(unittest.TestCase):
    def _document_xml(self, docx_path: Path) -> str:
        with zipfile.ZipFile(docx_path) as archive:
            return archive.read("word/document.xml").decode("utf-8")

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

            document_xml = self._document_xml(out_path)

            self.assertIn('w:eastAsia="Microsoft JhengHei"', document_xml)
            self.assertIn('w:eastAsia="Yu Gothic"', document_xml)
            self.assertIn('w:ascii="Calibri"', document_xml)

    def test_icons_are_preserved_when_kdp_mode_is_disabled(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "icons.md"
            out_path = Path(tmp_dir) / "icons.docx"
            md_path.write_text(ICON_MARKDOWN, encoding="utf-8")

            convert_file(md_path, out_path)

            document_xml = self._document_xml(out_path)
            self.assertIn("Shipping ✅", document_xml)
            self.assertIn("Docs 🔗", document_xml)
            self.assertIn("⚠️ Warning section", document_xml)
            self.assertIn("ℹ️ Info panel", document_xml)
            self.assertIn("👉 Start here", document_xml)
            self.assertIn("💡 Keep it simple", document_xml)
            self.assertIn("☑ ", document_xml)
            self.assertIn("☐ ", document_xml)

    def test_icon_map_is_ignored_when_kdp_mode_is_disabled(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "icons_no_kdp_map.md"
            out_path = Path(tmp_dir) / "icons_no_kdp_map.docx"
            md_path.write_text("Release ✅ and docs 🔗 with 🦉", encoding="utf-8")

            convert_file(
                md_path,
                out_path,
                kdp_safe_icons=False,
                icon_map={"✅": "[Approved]", "🔗": "(URL)", "🦉": "[Owl]"},
            )

            document_xml = self._document_xml(out_path)
            self.assertIn("Release ✅ and docs 🔗 with 🦉", document_xml)
            self.assertNotIn("[Approved]", document_xml)
            self.assertNotIn("(URL)", document_xml)
            self.assertNotIn("[Owl]", document_xml)

    def test_kdp_safe_icons_replace_common_symbols_but_keep_code(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "icons_kdp.md"
            out_path = Path(tmp_dir) / "icons_kdp.docx"
            md_path.write_text(ICON_MARKDOWN, encoding="utf-8")

            convert_file(md_path, out_path, kdp_safe_icons=True)

            document_xml = self._document_xml(out_path)
            self.assertIn("Shipping [OK]", document_xml)
            self.assertIn("Docs [Link]", document_xml)
            self.assertIn("[Warning] Warning section", document_xml)
            self.assertIn("[Info] Info panel", document_xml)
            self.assertIn("-&gt; Start here", document_xml)
            self.assertIn("[Tip] Keep it simple", document_xml)
            self.assertIn("[x] ", document_xml)
            self.assertIn("[ ] ", document_xml)
            self.assertIn("Code [OK] keeps emoji", document_xml)

    def test_kdp_safe_icons_cover_generic_semantics_and_zwj_variants(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "icons_generic_kdp.md"
            out_path = Path(tmp_dir) / "icons_generic_kdp.docx"
            md_path.write_text(GENERIC_ICON_MARKDOWN, encoding="utf-8")

            convert_file(md_path, out_path, kdp_safe_icons=True)

            document_xml = self._document_xml(out_path)
            self.assertIn("Routine [Strength]", document_xml)
            self.assertIn("[Calm] Reset", document_xml)
            self.assertIn("[Run] Sprint", document_xml)
            self.assertIn("[No] Avoid this", document_xml)
            self.assertIn("[Package] Deliverables", document_xml)
            self.assertIn("[Refresh] Weekly reset", document_xml)
            self.assertIn("[Chart] Metrics", document_xml)
            self.assertIn("[Calendar] Calendar", document_xml)
            self.assertIn("[Health] Health review", document_xml)
            self.assertIn("[Meal] Meal prep", document_xml)
            self.assertIn("[Day] Morning light", document_xml)
            self.assertIn("[Heart] Encouragement", document_xml)
            self.assertIn("Code [Calm] keeps emoji", document_xml)

    def test_custom_icon_map_can_override_defaults(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "icons_custom.md"
            out_path = Path(tmp_dir) / "icons_custom.docx"
            md_path.write_text("Release ✅ and docs 🔗", encoding="utf-8")

            convert_file(
                md_path,
                out_path,
                kdp_safe_icons=True,
                icon_map={"✅": "[Approved]", "🔗": "(URL)"},
            )

            document_xml = self._document_xml(out_path)
            self.assertIn("Release [Approved] and docs (URL)", document_xml)
            self.assertNotIn("Release [OK] and docs [Link]", document_xml)

    def test_custom_icon_map_can_override_normalized_default_icons(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "icons_custom_zwj.md"
            out_path = Path(tmp_dir) / "icons_custom_zwj.docx"
            md_path.write_text("🧘‍♀️ Reset and 🏃‍♀️ sprint", encoding="utf-8")

            convert_file(
                md_path,
                out_path,
                kdp_safe_icons=True,
                icon_map={"🧘‍♀️": "[Yoga]", "🏃‍♀️": "[Cardio]"},
            )

            document_xml = self._document_xml(out_path)
            self.assertIn("[Yoga] Reset and [Cardio] sprint", document_xml)
            self.assertNotIn("[Calm] Reset and [Run] sprint", document_xml)

    def test_kdp_safe_icons_fallback_to_unicode_names_for_unseen_icons(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "icons_auto_fallback.md"
            out_path = Path(tmp_dir) / "icons_auto_fallback.docx"
            md_path.write_text(AUTO_FALLBACK_ICON_MARKDOWN, encoding="utf-8")

            convert_file(md_path, out_path, kdp_safe_icons=True)

            document_xml = self._document_xml(out_path)
            self.assertIn("Future glyphs [Owl]", document_xml)
            self.assertIn("[Yarn] Yarn", document_xml)
            self.assertIn("[Graduation Cap] Study", document_xml)
            self.assertIn("[Flag US] Flag", document_xml)
            self.assertIn("[1] First step", document_xml)
            self.assertIn("Code [Owl] keeps emoji", document_xml)


if __name__ == "__main__":
    unittest.main()
