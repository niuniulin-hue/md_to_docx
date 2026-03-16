from __future__ import annotations

import tempfile
import unittest
import zipfile
import subprocess
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import convert
from docx import Document

from convert import (
    PDF_BACKEND_WEASYPRINT,
    convert_file,
    convert_file_to_pdf,
    export_markdown_to_pdf_via_weasyprint,
    export_docx_to_pdf,
    main,
    read_markdown_text,
    resolve_output_targets,
)


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
WEASY_MARKDOWN = "# Weasy ✅\n\nBody with [link 🔗](https://example.com).\n"


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

    def test_resolve_output_targets_defaults_to_matching_docx_and_pdf(self):
        source = Path("book.md")

        docx_only = resolve_output_targets(source)
        with_pdf = resolve_output_targets(source, pdf=True)

        self.assertEqual(docx_only.docx_path, Path("book.docx"))
        self.assertIsNone(docx_only.pdf_path)
        self.assertEqual(with_pdf.docx_path, Path("book.docx"))
        self.assertEqual(with_pdf.pdf_path, Path("book.pdf"))

    def test_resolve_output_targets_accepts_explicit_pdf_output(self):
        targets = resolve_output_targets(
            Path("book.md"),
            Path("dist/final-kdp.pdf"),
            pdf=True,
        )

        self.assertEqual(targets.docx_path, Path("dist/final-kdp.docx"))
        self.assertEqual(targets.pdf_path, Path("dist/final-kdp.pdf"))

    def test_resolve_output_targets_rejects_unsupported_extension_for_pdf(self):
        with self.assertRaises(ValueError):
            resolve_output_targets(Path("book.md"), Path("out.txt"), pdf=True)

    def test_convert_file_to_pdf_exports_pdf_after_docx_generation(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "pdf_source.md"
            docx_path = Path(tmp_dir) / "pdf_source.docx"
            pdf_path = Path(tmp_dir) / "pdf_source.pdf"
            md_path.write_text("# PDF export\n\nHello world ✅", encoding="utf-8")

            def fake_export(
                resolved_docx: Path,
                resolved_pdf: Path,
                *,
                timeout_seconds: int,
                pdf_backend: str,
                soffice_path: str | None,
            ) -> None:
                self.assertTrue(resolved_docx.exists())
                self.assertGreater(timeout_seconds, 0)
                self.assertEqual(pdf_backend, "auto")
                self.assertIsNone(soffice_path)
                resolved_pdf.write_bytes(b"%PDF-1.4\n% fake test pdf\n")

            with patch("convert.export_docx_to_pdf", side_effect=fake_export) as mocked_export:
                decoded = convert_file_to_pdf(md_path, docx_path, pdf_path, kdp_safe_icons=True)

            self.assertEqual(decoded.encoding.lower(), "utf-8")
            self.assertTrue(docx_path.exists())
            self.assertTrue(pdf_path.exists())
            mocked_export.assert_called_once_with(
                docx_path,
                pdf_path,
                timeout_seconds=120,
                pdf_backend="auto",
                soffice_path=None,
            )

    def test_convert_file_to_pdf_weasyprint_bypasses_docx_export(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "weasy_source.md"
            docx_path = Path(tmp_dir) / "weasy_source.docx"
            pdf_path = Path(tmp_dir) / "weasy_source.pdf"
            md_path.write_text(WEASY_MARKDOWN, encoding="utf-8")

            with patch("convert.export_markdown_to_pdf_via_weasyprint") as mocked_weasy:
                mocked_weasy.return_value = SimpleNamespace(encoding="utf-8")
                convert_file_to_pdf(
                    md_path,
                    docx_path,
                    pdf_path,
                    pdf_backend=PDF_BACKEND_WEASYPRINT,
                )

            mocked_weasy.assert_called_once_with(
                md_path,
                pdf_path,
                encoding=None,
                kdp_safe_icons=False,
                icon_map=None,
                css_paths=(),
                base_url=None,
                color_emoji_images=False,
                emoji_cdn_base="https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/svg",
            )

    def test_export_markdown_to_pdf_via_weasyprint_applies_kdp_mapping(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "weasy_icons.md"
            pdf_path = Path(tmp_dir) / "weasy_icons.pdf"
            css_path = Path(tmp_dir) / "book.css"
            md_path.write_text("# Ship ✅\n\nDocs 🔗\n", encoding="utf-8")
            css_path.write_text("body { font-size: 12pt; }", encoding="utf-8")

            seen: dict[str, object] = {}

            class FakeCSS:
                def __init__(self, filename: str):
                    seen["css_filename"] = filename

            class FakeHTML:
                def __init__(self, *, string: str, base_url: str):
                    seen["html_string"] = string
                    seen["base_url"] = base_url

                def write_pdf(self, target: str, *, stylesheets: list[object]) -> None:
                    seen["target"] = target
                    seen["stylesheets_count"] = len(stylesheets)
                    Path(target).write_bytes(b"%PDF-1.4\n% weasy fake\n")

            with patch.dict("sys.modules", {"weasyprint": SimpleNamespace(HTML=FakeHTML, CSS=FakeCSS)}):
                decoded = export_markdown_to_pdf_via_weasyprint(
                    md_path,
                    pdf_path,
                    kdp_safe_icons=True,
                    css_paths=(css_path,),
                )

            self.assertEqual(decoded.encoding.lower(), "utf-8")
            self.assertTrue(pdf_path.exists())
            html_string = str(seen["html_string"])
            # Labels are wrapped in coloured badge spans when kdp_safe_icons=True
            self.assertIn("Ship", html_string)
            self.assertIn("[OK]", html_string)
            # The span must carry a green foreground colour
            self.assertIn("color:#1a7f37", html_string)
            self.assertIn("Docs", html_string)
            self.assertIn("[Link]", html_string)
            # The span must carry a blue foreground colour
            self.assertIn("color:#0969da", html_string)
            self.assertEqual(seen["target"], str(pdf_path))
            self.assertEqual(seen["stylesheets_count"], 1)
            self.assertEqual(seen["css_filename"], str(css_path))
            self.assertEqual(seen["base_url"], str(md_path.parent.resolve()))

    def test_export_markdown_to_pdf_via_weasyprint_color_emoji_images(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "weasy_emoji.md"
            pdf_path = Path(tmp_dir) / "weasy_emoji.pdf"
            md_path.write_text("# Ship ✅\n\nDocs 🔗\n\n`Code ✅`\n", encoding="utf-8")

            seen: dict[str, object] = {}

            class FakeCSS:
                def __init__(self, filename: str):
                    self.filename = filename

            class FakeHTML:
                def __init__(self, *, string: str, base_url: str):
                    seen["html_string"] = string
                    seen["base_url"] = base_url

                def write_pdf(self, target: str, *, stylesheets: list[object]) -> None:
                    seen["target"] = target
                    Path(target).write_bytes(b"%PDF-1.4\n% weasy fake\n")

            with patch("convert._resolve_twemoji_svg_image_src", return_value="data:image/svg+xml;base64,PHN2Zy8+"):
                with patch.dict("sys.modules", {"weasyprint": SimpleNamespace(HTML=FakeHTML, CSS=FakeCSS)}):
                    export_markdown_to_pdf_via_weasyprint(
                        md_path,
                        pdf_path,
                    )

            html_string = str(seen["html_string"])
            self.assertIn('class="emoji"', html_string)
            self.assertIn("data:image/svg+xml;base64,PHN2Zy8+", html_string)
            self.assertIn("<code>Code ✅</code>", html_string)
            self.assertNotIn("<code>Code <img", html_string)

    def test_resolve_twemoji_svg_image_src_embeds_svg_data_uri(self):
        fake_response = SimpleNamespace(content=b"<svg xmlns='http://www.w3.org/2000/svg'></svg>")
        fake_response.raise_for_status = lambda: None

        with patch("requests.get", return_value=fake_response) as mocked_get:
            src = convert._resolve_twemoji_svg_image_src(
                "✅",
                asset_base="https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/svg",
                cache={},
            )

        self.assertIsNotNone(src)
        self.assertTrue(str(src).startswith("data:image/svg+xml;base64,"))
        mocked_get.assert_called_once()

    def test_export_docx_to_pdf_uses_worker_subprocess(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            docx_path = Path(tmp_dir) / "worker.docx"
            pdf_path = Path(tmp_dir) / "worker.pdf"
            docx_path.write_bytes(b"fake-docx")

            with patch("convert.find_soffice_executable", return_value=None):
                with patch(
                    "convert.subprocess.run",
                    return_value=SimpleNamespace(returncode=0, stderr="", stdout=""),
                ) as mocked_run:
                    export_docx_to_pdf(docx_path, pdf_path, timeout_seconds=5)

            self.assertEqual(mocked_run.call_count, 1)
            command = mocked_run.call_args.args[0]
            self.assertIn("--pdf-export-worker", command)
            self.assertIn(str(docx_path), command)
            self.assertIn(str(pdf_path), command)

    def test_export_docx_to_pdf_auto_falls_back_from_libreoffice_to_word(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            docx_path = Path(tmp_dir) / "worker_fallback.docx"
            pdf_path = Path(tmp_dir) / "worker_fallback.pdf"
            docx_path.write_bytes(b"fake-docx")

            with patch("convert.find_soffice_executable", return_value=Path("C:/LibreOffice/program/soffice.exe")):
                with patch(
                    "convert._export_docx_to_pdf_via_libreoffice",
                    side_effect=RuntimeError("LibreOffice failed"),
                ) as mocked_lo:
                    with patch("convert._export_docx_to_pdf_via_word") as mocked_word:
                        export_docx_to_pdf(docx_path, pdf_path, timeout_seconds=5)

            mocked_lo.assert_called_once_with(
                docx_path,
                pdf_path,
                timeout_seconds=5,
                soffice_path=None,
            )
            mocked_word.assert_called_once_with(
                docx_path,
                pdf_path,
                timeout_seconds=5,
            )

    def test_export_docx_to_pdf_times_out_cleanly(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            docx_path = Path(tmp_dir) / "worker_timeout.docx"
            pdf_path = Path(tmp_dir) / "worker_timeout.pdf"
            docx_path.write_bytes(b"fake-docx")

            with patch(
                "convert.subprocess.run",
                side_effect=subprocess.TimeoutExpired(cmd=["python"], timeout=1),
            ):
                with patch("convert._kill_word_process") as mocked_kill:
                    with self.assertRaises(RuntimeError) as context:
                        export_docx_to_pdf(docx_path, pdf_path, timeout_seconds=1)

            self.assertIn("timed out", str(context.exception).lower())
            mocked_kill.assert_called_once()

    def test_main_supports_pdf_output_path_for_single_file(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "cli_pdf.md"
            pdf_path = Path(tmp_dir) / "exports" / "cli_pdf.pdf"
            docx_path = pdf_path.with_suffix(".docx")
            md_path.write_text("# CLI PDF\n\nGenerated via --pdf", encoding="utf-8")

            def fake_export(
                resolved_docx: Path,
                resolved_pdf: Path,
                *,
                timeout_seconds: int,
                pdf_backend: str,
                soffice_path: str | None,
            ) -> None:
                self.assertEqual(resolved_docx, docx_path)
                self.assertEqual(timeout_seconds, 120)
                self.assertEqual(pdf_backend, "auto")
                self.assertIsNone(soffice_path)
                resolved_pdf.parent.mkdir(parents=True, exist_ok=True)
                resolved_pdf.write_bytes(b"%PDF-1.4\n% cli test pdf\n")

            argv = ["convert.py", str(md_path), "-o", str(pdf_path), "--pdf"]
            with patch("convert.export_docx_to_pdf", side_effect=fake_export):
                with patch("sys.argv", argv):
                    main()

            self.assertTrue(docx_path.exists())
            self.assertTrue(pdf_path.exists())

    def test_main_passes_pdf_timeout_to_export(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "cli_pdf_timeout.md"
            pdf_path = Path(tmp_dir) / "cli_pdf_timeout.pdf"
            docx_path = pdf_path.with_suffix(".docx")
            md_path.write_text("# CLI PDF timeout\n\nGenerated via --pdf-timeout", encoding="utf-8")

            seen: dict[str, object] = {}

            def fake_export(
                resolved_docx: Path,
                resolved_pdf: Path,
                *,
                timeout_seconds: int,
                pdf_backend: str,
                soffice_path: str | None,
            ) -> None:
                seen["docx"] = resolved_docx
                seen["pdf"] = resolved_pdf
                seen["timeout"] = timeout_seconds
                seen["backend"] = pdf_backend
                seen["soffice_path"] = soffice_path
                resolved_pdf.write_bytes(b"%PDF-1.4\n% timeout cli test pdf\n")

            argv = [
                "convert.py",
                str(md_path),
                "-o",
                str(pdf_path),
                "--pdf",
                "--pdf-timeout",
                "7",
            ]
            with patch("convert.export_docx_to_pdf", side_effect=fake_export):
                with patch("sys.argv", argv):
                    main()

            self.assertEqual(seen["docx"], docx_path)
            self.assertEqual(seen["pdf"], pdf_path)
            self.assertEqual(seen["timeout"], 7)
            self.assertEqual(seen["backend"], "auto")
            self.assertIsNone(seen["soffice_path"])
            self.assertTrue(pdf_path.exists())

    def test_main_weasy_backend_forwards_css_and_base_url(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "cli_weasy.md"
            pdf_path = Path(tmp_dir) / "cli_weasy.pdf"
            css_path = Path(tmp_dir) / "book.css"
            md_path.write_text("# Weasy ✅", encoding="utf-8")
            css_path.write_text("body { font-size: 12pt; }", encoding="utf-8")

            seen: dict[str, object] = {}

            def fake_weasy(
                resolved_md: Path,
                resolved_pdf: Path,
                *,
                encoding: str | None,
                kdp_safe_icons: bool,
                icon_map: dict[str, str] | None,
                css_paths: tuple[Path, ...],
                base_url: str | None,
                color_emoji_images: bool,
                emoji_cdn_base: str,
            ) -> SimpleNamespace:
                seen["md"] = resolved_md
                seen["pdf"] = resolved_pdf
                seen["encoding"] = encoding
                seen["kdp_safe_icons"] = kdp_safe_icons
                seen["icon_map"] = icon_map
                seen["css_paths"] = css_paths
                seen["base_url"] = base_url
                seen["color_emoji_images"] = color_emoji_images
                seen["emoji_cdn_base"] = emoji_cdn_base
                resolved_pdf.write_bytes(b"%PDF-1.4\n% cli weasy\n")
                return SimpleNamespace(encoding="utf-8")

            argv = [
                "convert.py",
                str(md_path),
                "-o",
                str(pdf_path),
                "--pdf",
                "--pdf-backend",
                PDF_BACKEND_WEASYPRINT,
                "--pdf-css",
                str(css_path),
                "--pdf-base-url",
                str(md_path.parent),
            ]

            with patch("convert.export_markdown_to_pdf_via_weasyprint", side_effect=fake_weasy):
                with patch("sys.argv", argv):
                    main()

            self.assertEqual(seen["md"], md_path)
            self.assertEqual(seen["pdf"], pdf_path)
            self.assertIsNone(seen["encoding"])
            self.assertFalse(seen["kdp_safe_icons"])
            self.assertIsNone(seen["icon_map"])
            self.assertEqual(seen["css_paths"], (css_path,))
            self.assertEqual(seen["base_url"], str(md_path.parent))
            self.assertFalse(seen["color_emoji_images"])
            self.assertEqual(
                seen["emoji_cdn_base"],
                "https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/svg",
            )
            self.assertTrue(pdf_path.exists())

    def test_main_rejects_pdf_css_without_weasy_backend(self):
        with tempfile.TemporaryDirectory() as tmp_dir:
            md_path = Path(tmp_dir) / "cli_invalid_css.md"
            css_path = Path(tmp_dir) / "book.css"
            md_path.write_text("# Invalid", encoding="utf-8")
            css_path.write_text("body { color: #111; }", encoding="utf-8")

            argv = [
                "convert.py",
                str(md_path),
                "--pdf",
                "--pdf-css",
                str(css_path),
            ]

            with patch("sys.argv", argv):
                with self.assertRaises(SystemExit) as context:
                    main()

            self.assertEqual(context.exception.code, 2)


if __name__ == "__main__":
    unittest.main()
