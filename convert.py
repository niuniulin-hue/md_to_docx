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
"""

from __future__ import annotations

import argparse
import json
import sys
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


@dataclass(frozen=True)
class DecodedMarkdown:
    text: str
    encoding: str
    source: str


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
    doc.save(str(out_path))
    print(f"  ✓  {md_path}  →  {out_path}  [{decoded.encoding}, {decoded.source}]")
    return decoded


def main():
    parser = argparse.ArgumentParser(
        description="Convert GitHub-Flavored Markdown to DOCX"
    )
    parser.add_argument("input", help="Input .md file or directory")
    parser.add_argument("-o", "--output", help="Output .docx file (single-file mode)")
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
    args = parser.parse_args()

    if args.icon_map and not args.kdp_safe_icons:
        parser.error("--icon-map requires --kdp-safe-icons")

    input_path = Path(args.input)
    template = Path(args.template) if args.template else None
    icon_map = load_icon_map(Path(args.icon_map)) if args.icon_map else None

    if input_path.is_dir():
        md_files = list(input_path.rglob("*.md"))
        if not md_files:
            print(f"No .md files found in {input_path}", file=sys.stderr)
            sys.exit(1)
        for md_file in md_files:
            out = md_file.with_suffix(".docx")
            convert_file(
                md_file,
                out,
                template,
                encoding=args.encoding,
                kdp_safe_icons=args.kdp_safe_icons,
                icon_map=icon_map,
            )
    elif input_path.is_file():
        if args.output:
            out = Path(args.output)
        else:
            out = input_path.with_suffix(".docx")
        convert_file(
            input_path,
            out,
            template,
            encoding=args.encoding,
            kdp_safe_icons=args.kdp_safe_icons,
            icon_map=icon_map,
        )
    else:
        print(f"Input not found: {input_path}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
