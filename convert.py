"""
convert.py — CLI entry point for md_to_docx

Usage
-----
    python convert.py input.md                       # -> input.docx
    python convert.py input.md -o output.docx
    python convert.py docs/                          # convert all .md in folder
    python convert.py input.md -o out.docx --template template.docx
    python convert.py input.md --encoding cp950
    python convert.py input.md -o out.docx --kdp-safe-icons
    python convert.py input.md -o out.docx --kdp-safe-icons --icon-map icon_map.json
    python convert.py input.md --pdf
    python convert.py input.md -o out.pdf --pdf
    python convert.py input.md -o out.pdf --pdf --pdf-timeout 30
"""

from __future__ import annotations

import argparse
import contextlib
import json
import os
import platform
import html as html_lib
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

from charset_normalizer import from_bytes


BOM_ENCODINGS = (
    (b"\xef\xbb\xbf", "utf-8-sig"),
    (b"\xff\xfe\x00\x00", "utf-32"),
    (b"\x00\x00\xfe\xff", "utf-32"),
    (b"\xff\xfe", "utf-16"),
    (b"\xfe\xff", "utf-16"),
)

FALLBACK_ENCODINGS = (
    "cp950",
    "big5hkscs",
    "big5",
    "cp932",
    "shift_jis",
    "euc_jp",
    "cp1252",
    "latin-1",
)

PDF_EXPORT_TIMEOUT_SECONDS = 120
PDF_BACKEND_AUTO = "auto"
PDF_BACKEND_LIBREOFFICE = "libreoffice"
PDF_BACKEND_WORD = "word"
PDF_BACKEND_WEASYPRINT = "weasyprint"
PDF_BACKEND_CHOICES = (
    PDF_BACKEND_AUTO,
    PDF_BACKEND_LIBREOFFICE,
    PDF_BACKEND_WORD,
    PDF_BACKEND_WEASYPRINT,
)


@dataclass(frozen=True)
class DecodedMarkdown:
    text: str
    encoding: str
    source: str


@dataclass(frozen=True)
class OutputTargets:
    docx_path: Path
    pdf_path: Path | None = None


def load_icon_map(icon_map_path: Path | None) -> dict[str, str] | None:
    """Load a JSON object that maps source glyphs to DOCX-safe replacements."""
    if icon_map_path is None:
        return None

    data = json.loads(icon_map_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError(f"Icon map must be a JSON object: {icon_map_path}")

    replacements: dict[str, str] = {}
    for source, target in data.items():
        if not isinstance(source, str) or not isinstance(target, str):
            raise ValueError("Icon map keys and values must both be strings")
        if source:
            replacements[source] = target
    return replacements


def _decode_bytes(raw_bytes: bytes, encoding: str, source: str) -> DecodedMarkdown:
    return DecodedMarkdown(
        text=raw_bytes.decode(encoding, errors="strict"),
        encoding=encoding,
        source=source,
    )


def read_markdown_text(md_path: Path, encoding: str | None = None) -> DecodedMarkdown:
    """Read Markdown text with robust multi-encoding support.

    Supports explicit encoding override, BOM-aware Unicode decoding, automatic
    charset detection, and common Traditional Chinese / Japanese / Western
    encodings used by Markdown files.
    """
    raw_bytes = md_path.read_bytes()
    if not raw_bytes:
        return DecodedMarkdown(text="", encoding=encoding or "utf-8", source="empty-file")

    if encoding:
        return _decode_bytes(raw_bytes, encoding, "user-specified")

    for bom_bytes, bom_encoding in BOM_ENCODINGS:
        if raw_bytes.startswith(bom_bytes):
            return _decode_bytes(raw_bytes, bom_encoding, "bom")

    try:
        return _decode_bytes(raw_bytes, "utf-8", "default-utf-8")
    except UnicodeDecodeError:
        pass

    detection = from_bytes(raw_bytes).best()
    if detection and detection.encoding:
        try:
            return _decode_bytes(raw_bytes, detection.encoding, "charset-normalizer")
        except UnicodeDecodeError:
            pass

    attempted = ["utf-8"]
    for fallback_encoding in FALLBACK_ENCODINGS:
        attempted.append(fallback_encoding)
        try:
            return _decode_bytes(raw_bytes, fallback_encoding, "fallback")
        except UnicodeDecodeError:
            continue

    raise UnicodeDecodeError(
        "markdown",
        raw_bytes,
        0,
        1,
        f"unable to decode {md_path} with encodings: {', '.join(attempted)}",
    )


def resolve_output_targets(
    input_path: Path,
    output_path: Path | None = None,
    *,
    pdf: bool = False,
) -> OutputTargets:
    """Resolve DOCX/PDF output locations for a single Markdown input."""
    if output_path is None:
        docx_path = input_path.with_suffix(".docx")
        pdf_path = input_path.with_suffix(".pdf") if pdf else None
        return OutputTargets(docx_path=docx_path, pdf_path=pdf_path)

    if not pdf:
        return OutputTargets(docx_path=output_path)

    suffix = output_path.suffix.lower()
    if suffix == ".pdf":
        return OutputTargets(
            docx_path=output_path.with_suffix(".docx"),
            pdf_path=output_path,
        )
    if suffix == ".docx":
        return OutputTargets(
            docx_path=output_path,
            pdf_path=output_path.with_suffix(".pdf"),
        )
    if suffix:
        raise ValueError(
            f"Unsupported output extension for PDF export: {output_path}. "
            "Use .docx, .pdf, or omit the extension."
        )

    return OutputTargets(
        docx_path=output_path.with_suffix(".docx"),
        pdf_path=output_path.with_suffix(".pdf"),
    )


def _kill_word_process(pid: int | None) -> None:
    if not pid or platform.system() != "Windows":
        return

    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception:
        pass


def find_soffice_executable(explicit_path: str | None = None) -> Path | None:
    """Locate a LibreOffice/OpenOffice soffice executable."""
    candidates: list[Path] = []

    if explicit_path:
        candidates.append(Path(explicit_path))
    else:
        command_names = (
            ("soffice.com", "soffice", "libreoffice")
            if platform.system() == "Windows"
            else ("soffice", "libreoffice")
        )
        for command_name in command_names:
            resolved = shutil.which(command_name)
            if resolved:
                candidates.append(Path(resolved))

        if platform.system() == "Windows":
            program_files = [
                Path(path)
                for path in (
                    os.environ.get("ProgramFiles"),
                    os.environ.get("ProgramFiles(x86)"),
                )
                if path
            ]
            for base_dir in program_files:
                candidates.extend(
                    [
                        base_dir / "LibreOffice" / "program" / "soffice.com",
                        base_dir / "LibreOffice" / "program" / "soffice.exe",
                        base_dir / "OpenOffice 4" / "program" / "soffice.com",
                        base_dir / "OpenOffice 4" / "program" / "soffice.exe",
                    ]
                )

    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def _export_docx_to_pdf_via_libreoffice(
    docx_path: Path,
    pdf_path: Path,
    *,
    timeout_seconds: int,
    soffice_path: str | None = None,
) -> None:
    executable = find_soffice_executable(soffice_path)
    if executable is None:
        raise RuntimeError(
            "LibreOffice/OpenOffice was not found. Install LibreOffice or pass --pdf-backend word."
        )

    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp_dir:
        out_dir = Path(tmp_dir)
        profile_dir = out_dir / "lo-profile"
        profile_dir.mkdir(parents=True, exist_ok=True)
        expected_pdf = out_dir / docx_path.with_suffix(".pdf").name
        command = [
            str(executable),
            "--headless",
            "--nologo",
            "--nodefault",
            "--norestore",
            "--nolockcheck",
            f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(out_dir),
            str(docx_path.resolve()),
        ]

        try:
            result = subprocess.run(
                command,
                check=False,
                capture_output=True,
                text=True,
                timeout=timeout_seconds,
            )
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(
                f"LibreOffice export timed out after {timeout_seconds} seconds."
            ) from exc

        if result.returncode != 0:
            details = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(details or f"LibreOffice export failed with exit code {result.returncode}.")

        if not expected_pdf.exists():
            details = (result.stderr or result.stdout or "").strip()
            raise RuntimeError(
                details
                or f"LibreOffice did not create the expected PDF output: {expected_pdf}"
            )

        pdf_path.write_bytes(expected_pdf.read_bytes())


def _export_docx_to_pdf_via_word(
    docx_path: Path,
    pdf_path: Path,
    *,
    timeout_seconds: int,
) -> None:
    """Export a DOCX file to PDF with a watchdog timeout around the Word worker."""
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pid") as handle:
        pid_file = Path(handle.name)

    command = [
        sys.executable,
        str(Path(__file__).resolve()),
        "--pdf-export-worker",
        str(docx_path),
        str(pdf_path),
        str(pid_file),
    ]

    try:
        result = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        word_pid = None
        if pid_file.exists():
            try:
                raw_pid = pid_file.read_text(encoding="utf-8").strip()
                word_pid = int(raw_pid) if raw_pid else None
            except Exception:
                word_pid = None
        _kill_word_process(word_pid)
        raise RuntimeError(
            f"Word export timed out after {timeout_seconds} seconds. "
            "Microsoft Word may be waiting on a hidden dialog or stalled during export."
        ) from exc
    finally:
        try:
            pid_file.unlink(missing_ok=True)
        except Exception:
            pass

    if result.returncode != 0:
        details = (result.stderr or result.stdout or "").strip()
        raise RuntimeError(details or f"Word export worker failed for DOCX: {docx_path}")


def _export_docx_to_pdf_worker(
    docx_path: Path,
    pdf_path: Path,
    *,
    pid_file: Path | None = None,
) -> None:
    """Worker process that talks to Word COM directly."""
    if platform.system() != "Windows":
        raise RuntimeError("PDF export is currently supported only on Windows with Microsoft Word installed.")

    try:
        import pythoncom
        from win32com.client import DispatchEx
        from win32process import GetWindowThreadProcessId
    except ImportError as exc:
        raise RuntimeError(
            "PDF export on Windows requires the optional 'pywin32' dependency. "
            "Install it with 'pip install -r requirements.txt'."
        ) from exc

    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    resolved_docx = str(docx_path.resolve())
    resolved_pdf = str(pdf_path.resolve())
    word = None
    document = None
    word_pid = None

    try:
        pythoncom.CoInitialize()
        word = DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        try:
            _, word_pid = GetWindowThreadProcessId(word.Hwnd)
        except Exception:
            word_pid = None

        if pid_file is not None and word_pid:
            pid_file.write_text(str(word_pid), encoding="utf-8")

        document = word.Documents.Open(
            resolved_docx,
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=True,
            NoEncodingDialog=True,
        )
        document.ExportAsFixedFormat(
            OutputFileName=resolved_pdf,
            ExportFormat=17,
        )
    except Exception as exc:  # pragma: no cover - backend/runtime specific
        raise RuntimeError(
            "Failed to export PDF. On Windows, Microsoft Word must be installed and available for COM automation. "
            f"DOCX: {docx_path} | PDF: {pdf_path}"
        ) from exc
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def export_docx_to_pdf(
    docx_path: Path,
    pdf_path: Path,
    *,
    timeout_seconds: int = PDF_EXPORT_TIMEOUT_SECONDS,
    pdf_backend: str = PDF_BACKEND_AUTO,
    soffice_path: str | None = None,
) -> None:
    """Export a DOCX file to PDF using the selected backend or automatic fallback."""
    if pdf_backend not in PDF_BACKEND_CHOICES:
        raise RuntimeError(
            f"Unsupported PDF backend '{pdf_backend}'. Choose from: {', '.join(PDF_BACKEND_CHOICES)}"
        )

    backends: list[str]
    if pdf_backend == PDF_BACKEND_AUTO:
        backends = []
        if find_soffice_executable(soffice_path) is not None:
            backends.append(PDF_BACKEND_LIBREOFFICE)
        if platform.system() == "Windows":
            backends.append(PDF_BACKEND_WORD)
        elif not backends:
            raise RuntimeError(
                "No PDF backend is available. Install LibreOffice or run PDF export on Windows with Microsoft Word."
            )
        if not backends:
            raise RuntimeError(
                "No PDF backend is available. Install LibreOffice or use Windows with Microsoft Word installed."
            )
    else:
        backends = [pdf_backend]

    failures: list[str] = []
    for backend_name in backends:
        try:
            if backend_name == PDF_BACKEND_LIBREOFFICE:
                _export_docx_to_pdf_via_libreoffice(
                    docx_path,
                    pdf_path,
                    timeout_seconds=timeout_seconds,
                    soffice_path=soffice_path,
                )
            elif backend_name == PDF_BACKEND_WORD:
                _export_docx_to_pdf_via_word(
                    docx_path,
                    pdf_path,
                    timeout_seconds=timeout_seconds,
                )
            else:
                raise RuntimeError(f"Unsupported PDF backend: {backend_name}")
            print(f"  ✓  {docx_path}  →  {pdf_path}  [pdf via {backend_name}]")
            return
        except RuntimeError as exc:
            failures.append(f"{backend_name}: {exc}")

    raise RuntimeError("PDF export failed. Tried backends: " + " | ".join(failures))


def _markdown_to_html_document(markdown_text: str, *, title: str) -> str:
    """Render Markdown content into a print-friendly HTML document string."""
    import mistune

    renderer = mistune.HTMLRenderer(escape=False)
    markdown = mistune.create_markdown(
        renderer=renderer,
        plugins=["strikethrough", "table", "task_lists", "url"],
    )
    body_html = markdown(markdown_text)
    escaped_title = html_lib.escape(title)

    # Keep the built-in CSS intentionally conservative so users can layer custom CSS.
    return (
        "<!doctype html>\n"
        "<html lang=\"en\">\n"
        "<head>\n"
        "  <meta charset=\"utf-8\">\n"
        f"  <title>{escaped_title}</title>\n"
        "  <style>\n"
        "    @page { size: A4; margin: 20mm 18mm 22mm 18mm; }\n"
        "    body { font-family: 'Noto Serif CJK SC', 'Noto Serif CJK TC', 'Noto Serif CJK JP', 'Noto Serif', 'Segoe UI', serif; line-height: 1.6; font-size: 11pt; color: #111; }\n"
        "    h1, h2, h3, h4, h5, h6 { font-weight: 700; line-height: 1.25; page-break-after: avoid; }\n"
        "    p, li, blockquote, table, pre { orphans: 3; widows: 3; }\n"
        "    a { color: #1a0dab; text-decoration: none; }\n"
        "    code, pre { font-family: 'Consolas', 'Cascadia Mono', 'Courier New', monospace; }\n"
        "    pre { background: #f6f8fa; padding: 10px 12px; border-radius: 4px; overflow-wrap: anywhere; }\n"
        "    blockquote { margin: 0; padding-left: 12px; border-left: 3px solid #d0d7de; color: #444; }\n"
        "    table { border-collapse: collapse; width: 100%; margin: 12px 0; }\n"
        "    th, td { border: 1px solid #d0d7de; padding: 6px 8px; vertical-align: top; }\n"
        "    th { background: #f6f8fa; }\n"
        "    img { max-width: 100%; height: auto; }\n"
        "  </style>\n"
        "</head>\n"
        "<body>\n"
        f"{body_html}\n"
        "</body>\n"
        "</html>\n"
    )


def _normalize_markdown_for_kdp_icons(
    markdown_text: str,
    *,
    kdp_safe_icons: bool,
    icon_map: dict[str, str] | None,
) -> str:
    if not kdp_safe_icons:
        return markdown_text

    from md_to_docx import renderer as renderer_module

    replacement_lookup = dict(
        renderer_module.build_text_replacements(
            kdp_safe_icons=True,
            icon_map=icon_map,
        )
    )

    parts: list[str] = []
    index = 0
    while index < len(markdown_text):
        char = markdown_text[index]
        if not renderer_module._is_icon_like_char(char) and not renderer_module._starts_keycap_cluster(markdown_text, index):
            parts.append(char)
            index += 1
            continue

        cluster, next_index = renderer_module._consume_icon_cluster(markdown_text, index)
        normalized = cluster.translate(renderer_module._TEXT_NORMALIZATION_TABLE)
        replacement = replacement_lookup.get(normalized)
        if replacement is not None:
            parts.append(replacement)
        elif normalized:
            parts.append(renderer_module._fallback_label_for_cluster(cluster) or normalized)
        else:
            parts.append(cluster)
        index = next_index

    return "".join(parts)


_GTK3_WINDOWS_INSTALL_GUIDE = """
WeasyPrint requires the GTK3 runtime libraries, which are not bundled with the pip package.

On Windows, install them in one of these two ways:

Option A — GTK3 standalone installer (simplest):
  1. Download the latest installer from:
       https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases
  2. Run it and choose "Add to PATH for all users" (or current user).
  3. Reopen your terminal so the new PATH takes effect, then retry.

Option B — MSYS2 (if you already have it):
  In an MSYS2 MinGW64 shell:
      pacman -S mingw-w64-x86_64-gtk3 mingw-w64-x86_64-python-weasyprint
  Add C:\\msys64\\mingw64\\bin to your system PATH, then retry from a normal terminal.

For full details see:
  https://doc.courtbouillon.org/weasyprint/stable/first_steps.html#windows
""".strip()


def _raise_weasyprint_gtk_error(exc: OSError) -> None:  # pragma: no cover
    """Re-raise an OSError from a missing GTK3 DLL with a user-friendly message."""
    msg = str(exc)
    if platform.system() == "Windows" and (
        "libgobject" in msg or "libpango" in msg or "libcairo" in msg or "error 0x7e" in msg.lower()
    ):
        raise RuntimeError(
            "WeasyPrint cannot find the GTK3 system libraries.\n\n"
            + _GTK3_WINDOWS_INSTALL_GUIDE
        ) from exc
    raise RuntimeError(
        "WeasyPrint failed to load a required system library. "
        "Ensure GTK3 (libgobject, libpango, libcairo) is installed and on PATH.\n"
        f"Original error: {exc}"
    ) from exc


def _find_gtk3_bin_dirs() -> list[Path]:
    """Return every directory that looks like it contains GTK3 runtime DLLs."""
    sentinel = "libgobject-2.0-0.dll"
    candidates: list[Path] = []

    # 1. Well-known install locations (our script + system-wide installers)
    for base in filter(None, [
        os.environ.get("LOCALAPPDATA"),
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramFiles(x86)"),
        "C:\\",
    ]):
        for sub in (
            "GTK3-Runtime\\bin",
            "GTK3-Runtime Win64\\bin",
            "GTK3-Runtime Win32\\bin",
            "msys64\\mingw64\\bin",
            "msys2\\mingw64\\bin",
        ):
            candidates.append(Path(base) / sub)

    # 2. Anything already on PATH
    for entry in os.environ.get("PATH", "").split(os.pathsep):
        if entry:
            candidates.append(Path(entry))

    return [p for p in candidates if (p / sentinel).is_file()]


def _register_gtk3_dll_dirs() -> list[Path]:
    """Call os.add_dll_directory() for every GTK3 bin dir found (Windows only).

    Also suppresses the harmless GLib-GIO-WARNING spam that GTK3 emits on
    Windows when it scans UWP app registrations.
    """
    if platform.system() != "Windows":
        return []
    # Prevent GTK3/GIO from scanning Windows UWP associations (harmless but noisy)
    os.environ.setdefault("GIO_USE_VFS", "local")
    os.environ.setdefault("GSETTINGS_BACKEND", "memory")
    registered: list[Path] = []
    for gtk_bin in _find_gtk3_bin_dirs():
        try:
            os.add_dll_directory(str(gtk_bin))  # type: ignore[attr-defined]
            registered.append(gtk_bin)
        except (OSError, AttributeError):
            pass
    return registered


@contextlib.contextmanager
def _suppress_c_stderr():
    """Redirect C-level stderr (fd 2) to devnull for the duration of the block.

    GTK3's GLib writes GLib-GIO-WARNING lines about UWP apps directly to fd 2,
    bypassing Python's sys.stderr.  Redirecting at the OS level silences them
    without hiding Python exceptions (which propagate via the exception
    mechanism, not via stderr).
    """
    if platform.system() != "Windows":
        yield
        return

    try:
        devnull_fd = os.open(os.devnull, os.O_WRONLY)
        saved_fd2 = os.dup(2)
        os.dup2(devnull_fd, 2)
        try:
            yield
        finally:
            os.dup2(saved_fd2, 2)
            os.close(saved_fd2)
            os.close(devnull_fd)
    except OSError:
        # If fd manipulation fails for any reason, run without suppression.
        yield


def export_markdown_to_pdf_via_weasyprint(
    md_path: Path,
    pdf_path: Path,
    *,
    encoding: str | None = None,
    kdp_safe_icons: bool = False,
    icon_map: dict[str, str] | None = None,
    css_paths: tuple[Path, ...] = (),
    base_url: str | None = None,
) -> DecodedMarkdown:
    """Export Markdown directly to PDF using WeasyPrint (Markdown -> HTML -> PDF)."""
    # Register GTK3 DLL directories BEFORE the first import of weasyprint so
    # that cffi can resolve libgobject / libpango / libcairo on Windows even
    # when the user opened their terminal before the PATH was updated.
    _register_gtk3_dll_dirs()

    decoded = read_markdown_text(md_path, encoding=encoding)
    normalized_markdown = _normalize_markdown_for_kdp_icons(
        decoded.text,
        kdp_safe_icons=kdp_safe_icons,
        icon_map=icon_map,
    )

    html_document = _markdown_to_html_document(normalized_markdown, title=md_path.stem)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        with _suppress_c_stderr():
            from weasyprint import CSS, HTML
    except ImportError as exc:  # pragma: no cover - optional dependency
        raise RuntimeError(
            "WeasyPrint backend requires the optional 'weasyprint' dependency. "
            "Install it with 'pip install weasyprint' and ensure system libraries are available."
        ) from exc

    try:
        with _suppress_c_stderr():
            stylesheets = [CSS(filename=str(path)) for path in css_paths]
            html = HTML(string=html_document, base_url=base_url or str(md_path.parent.resolve()))
            html.write_pdf(str(pdf_path), stylesheets=stylesheets)
    except Exception as exc:  # pragma: no cover - backend/runtime specific
        raise RuntimeError(
            "WeasyPrint PDF export failed. Verify CSS paths, fonts, and required system dependencies. "
            f"Markdown: {md_path} | PDF: {pdf_path}"
        ) from exc

    print(f"  ✓  {md_path}  →  {pdf_path}  [pdf via {PDF_BACKEND_WEASYPRINT}]")
    return decoded


def _run_pdf_export_worker_cli(args: list[str]) -> int:
    if len(args) != 3:
        print(
            "Usage: convert.py --pdf-export-worker <input.docx> <output.pdf> <pid-file>",
            file=sys.stderr,
        )
        return 2

    docx_path = Path(args[0])
    pdf_path = Path(args[1])
    pid_file = Path(args[2])

    try:
        _export_docx_to_pdf_worker(docx_path, pdf_path, pid_file=pid_file)
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 1
    return 0


def convert_file(
    md_path: Path,
    out_path: Path,
    template: Path | None = None,
    encoding: str | None = None,
    *,
    kdp_safe_icons: bool = False,
    icon_map: dict[str, str] | None = None,
) -> DecodedMarkdown:
    """Convert a single Markdown file to DOCX."""
    import mistune
    from docx import Document
    from md_to_docx.renderer import ast_to_docx

    decoded = read_markdown_text(md_path, encoding=encoding)

    md = mistune.create_markdown(
        renderer=None,
        plugins=["strikethrough", "table", "task_lists", "url"],
    )
    tokens = md(decoded.text)

    if template and template.is_file():
        doc = Document(str(template))
        for element in list(doc.element.body):
            tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag
            if tag not in ("sectPr",):
                doc.element.body.remove(element)
    else:
        doc = Document()

    effective_icon_map = icon_map if kdp_safe_icons else None

    doc = ast_to_docx(
        tokens,
        doc,
        kdp_safe_icons=kdp_safe_icons,
        icon_map=effective_icon_map,
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    print(f"  ✓  {md_path}  →  {out_path}  [{decoded.encoding}, {decoded.source}]")
    return decoded


def convert_file_to_pdf(
    md_path: Path,
    docx_path: Path,
    pdf_path: Path,
    template: Path | None = None,
    encoding: str | None = None,
    *,
    kdp_safe_icons: bool = False,
    icon_map: dict[str, str] | None = None,
    pdf_timeout_seconds: int = PDF_EXPORT_TIMEOUT_SECONDS,
    pdf_backend: str = PDF_BACKEND_AUTO,
    soffice_path: str | None = None,
    pdf_css_paths: tuple[Path, ...] = (),
    pdf_base_url: str | None = None,
) -> DecodedMarkdown:
    """Convert a single Markdown file to DOCX and then export it to PDF."""
    if pdf_backend == PDF_BACKEND_WEASYPRINT:
        return export_markdown_to_pdf_via_weasyprint(
            md_path,
            pdf_path,
            encoding=encoding,
            kdp_safe_icons=kdp_safe_icons,
            icon_map=icon_map,
            css_paths=pdf_css_paths,
            base_url=pdf_base_url,
        )

    decoded = convert_file(
        md_path,
        docx_path,
        template,
        encoding=encoding,
        kdp_safe_icons=kdp_safe_icons,
        icon_map=icon_map,
    )
    export_docx_to_pdf(
        docx_path,
        pdf_path,
        timeout_seconds=pdf_timeout_seconds,
        pdf_backend=pdf_backend,
        soffice_path=soffice_path,
    )
    return decoded


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--pdf-export-worker":
        sys.exit(_run_pdf_export_worker_cli(sys.argv[2:]))

    parser = argparse.ArgumentParser(
        description="Convert GitHub-Flavored Markdown to DOCX and optionally PDF"
    )
    parser.add_argument("input", help="Input .md file or directory")
    parser.add_argument(
        "-o",
        "--output",
        help="Output file in single-file mode (.docx by default; .pdf also supported with --pdf)",
    )
    parser.add_argument(
        "--template",
        help="Optional .docx template file for styles/fonts",
        default=None,
    )
    parser.add_argument(
        "--encoding",
        help="Optional source Markdown encoding override, e.g. utf-8, cp950, shift_jis, cp1252",
        default=None,
    )
    parser.add_argument(
        "--kdp-safe-icons",
        action="store_true",
        help="Replace common Markdown emoji/icons with KDP-safe plain-text equivalents",
    )
    parser.add_argument(
        "--icon-map",
        help="Optional UTF-8 JSON file with custom text replacements; requires --kdp-safe-icons",
        default=None,
    )
    parser.add_argument(
        "--pdf",
        action="store_true",
        help="Also export a PDF using the selected backend (auto tries LibreOffice first, then Word on Windows)",
    )
    parser.add_argument(
        "--pdf-timeout",
        type=int,
        default=PDF_EXPORT_TIMEOUT_SECONDS,
        help="Timeout in seconds for the PDF export worker when --pdf is enabled",
    )
    parser.add_argument(
        "--pdf-backend",
        choices=PDF_BACKEND_CHOICES,
        default=PDF_BACKEND_AUTO,
        help="PDF export backend: auto tries LibreOffice first (if available), then Word on Windows; use weasyprint for Markdown+CSS direct PDF",
    )
    parser.add_argument(
        "--soffice-path",
        default=None,
        help="Optional path to soffice.exe / libreoffice executable for PDF export",
    )
    parser.add_argument(
        "--pdf-css",
        action="append",
        default=[],
        help="Optional CSS file(s) used only with --pdf-backend weasyprint (can be repeated)",
    )
    parser.add_argument(
        "--pdf-base-url",
        default=None,
        help="Base URL/path for resolving relative assets in weasyprint HTML mode",
    )
    args = parser.parse_args()

    if args.icon_map and not args.kdp_safe_icons:
        parser.error("--icon-map requires --kdp-safe-icons")
    if args.pdf_timeout <= 0:
        parser.error("--pdf-timeout must be a positive integer")
    if args.soffice_path and args.pdf_backend not in (PDF_BACKEND_AUTO, PDF_BACKEND_LIBREOFFICE):
        parser.error("--soffice-path is only valid with --pdf-backend auto/libreoffice")

    pdf_css_paths = tuple(Path(css_path) for css_path in args.pdf_css)
    missing_css_paths = [str(path) for path in pdf_css_paths if not path.is_file()]
    if missing_css_paths:
        parser.error("--pdf-css file not found: " + ", ".join(missing_css_paths))
    if pdf_css_paths and args.pdf_backend != PDF_BACKEND_WEASYPRINT:
        parser.error("--pdf-css requires --pdf-backend weasyprint")
    if args.pdf_base_url and args.pdf_backend != PDF_BACKEND_WEASYPRINT:
        parser.error("--pdf-base-url requires --pdf-backend weasyprint")

    input_path = Path(args.input)
    template = Path(args.template) if args.template else None
    icon_map = load_icon_map(Path(args.icon_map)) if args.icon_map else None

    if input_path.is_dir():
        if args.output:
            parser.error("-o/--output is only supported for a single input file")
        md_files = list(input_path.rglob("*.md"))
        if not md_files:
            print(f"No .md files found in {input_path}", file=sys.stderr)
            sys.exit(1)
        for md_file in md_files:
            targets = resolve_output_targets(md_file, pdf=args.pdf)
            if args.pdf and targets.pdf_path is not None:
                convert_file_to_pdf(
                    md_file,
                    targets.docx_path,
                    targets.pdf_path,
                    template,
                    encoding=args.encoding,
                    kdp_safe_icons=args.kdp_safe_icons,
                    icon_map=icon_map,
                    pdf_timeout_seconds=args.pdf_timeout,
                    pdf_backend=args.pdf_backend,
                    soffice_path=args.soffice_path,
                    pdf_css_paths=pdf_css_paths,
                    pdf_base_url=args.pdf_base_url,
                )
            else:
                convert_file(
                    md_file,
                    targets.docx_path,
                    template,
                    encoding=args.encoding,
                    kdp_safe_icons=args.kdp_safe_icons,
                    icon_map=icon_map,
                )
    elif input_path.is_file():
        try:
            targets = resolve_output_targets(
                input_path,
                Path(args.output) if args.output else None,
                pdf=args.pdf,
            )
        except ValueError as exc:
            parser.error(str(exc))

        if args.pdf and targets.pdf_path is not None:
            convert_file_to_pdf(
                input_path,
                targets.docx_path,
                targets.pdf_path,
                template,
                encoding=args.encoding,
                kdp_safe_icons=args.kdp_safe_icons,
                icon_map=icon_map,
                pdf_timeout_seconds=args.pdf_timeout,
                pdf_backend=args.pdf_backend,
                soffice_path=args.soffice_path,
                pdf_css_paths=pdf_css_paths,
                pdf_base_url=args.pdf_base_url,
            )
        else:
            convert_file(
                input_path,
                targets.docx_path,
                template,
                encoding=args.encoding,
                kdp_safe_icons=args.kdp_safe_icons,
                icon_map=icon_map,
            )
    else:
        print(f"Input not found: {input_path}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
