# md_to_docx

Convert **GitHub-Flavored Markdown** (`.md`) files to well-formatted Word documents (`.docx`).

## Features

| GFM Element | Supported |
|---|---|
| Headings H1–H6 | ✅ |
| Bold, Italic, Strikethrough | ✅ |
| Inline code | ✅ (monospace, colored) |
| Fenced & indented code blocks | ✅ (monospace, gray background) |
| Unordered lists (nested) | ✅ |
| Ordered lists (nested) | ✅ |
| Task lists (checkboxes) | ✅ |
| Blockquotes (nested) | ✅ |
| Tables | ✅ (with styled header row) |
| Links (clickable) | ✅ |
| Images (local & URL) | ✅ |
| Horizontal rules | ✅ |
| Traditional Chinese / 繁體中文 | ✅ |
| Japanese / 日本語 | ✅ |
| French / Français | ✅ |
| Spanish / Español | ✅ |
| German / Deutsch | ✅ |
| Encoding auto-detection | ✅ |
| KDP-safe icon replacement | ✅ |

## Encoding Support

The converter now supports **multi-language, multi-encoding Markdown input**.

### Automatically handled encodings

- `utf-8`
- `utf-8-sig` (BOM)
- `cp950` / `big5` / `big5hkscs` for Traditional Chinese files
- `cp932` / `shift_jis` / `euc_jp` for Japanese files
- `cp1252` / `latin-1` for French, Spanish, and German legacy files

### Font strategy in DOCX

To improve Word rendering for multilingual content, generated documents apply:

- Latin text → `Calibri`
- Traditional Chinese text → `Microsoft JhengHei`
- Japanese text → `Yu Gothic`
- Code text → `Consolas` + East Asia fallback

If you want your own corporate fonts, use a Word template via `--template`.

## Amazon KDP / Kindle-Friendly Icon Mapping

Some GitHub-style Unicode icons and emoji render inconsistently in Kindle Create / Amazon KDP workflows.
This project now supports an **optional KDP-safe replacement mode** that converts common symbols to plain-text equivalents before writing the DOCX.
If you do **not** pass `--kdp-safe-icons`, the original icons are preserved as-is.

### Built-in examples

- `✅` / `✔️` → `[OK]`
- `☑` → `[x]`
- `☐` → `[ ]`
- `⚠️` → `[Warning]`
- `ℹ️` → `[Info]`
- `💡` → `[Tip]`
- `🔗` → `[Link]`
- `👉` / `➡️` → `->`

### Built-in generic semantic categories

The default KDP-safe mode is intentionally **generic rather than book-specific**. It maps many common emoji/icons to short reusable labels so the same Markdown can work across different books:

- Status / alerts: `[OK]`, `[X]`, `[Warning]`, `[Info]`, `[Alert]`, `[No]`, `[Stop]`
- Notes / structure: `[Note]`, `[Tip]`, `[Goal]`, `[Book]`, `[List]`, `[Package]`, `[Calendar]`, `[Guide]`, `[Link]`
- Progress / navigation: `[Chart]`, `[Up]`, `[Down]`, `[Review]`, `[Refresh]`, `[Finish]`, `[Steps]`, `->`
- Time / routine: `[Time]`, `[Morning]`, `[Day]`, `[Evening]`, `[Night]`
- Health / activity: `[Strength]`, `[Run]`, `[Walk]`, `[Calm]`, `[Rest]`, `[Sleep]`, `[Health]`, `[Medication]`, `[Water]`
- Work / devices: `[Work]`, `[Laptop]`, `[Phone]`, `[Home]`
- Food / drink: `[Meal]`, `[Food]`, `[Fruit]`, `[Vegetable]`, `[Protein]`, `[Drink]`
- Tone / emphasis: `[Heart]`

### Automatic fallback for previously unseen icons

When `--kdp-safe-icons` is enabled, the converter does **not rely only on a fixed manual table**.
If it encounters an emoji/icon that is not already covered by the built-in semantic mappings, it automatically falls back to an ASCII-safe label derived from the Unicode name.

Examples:

- `🦉` → `[Owl]`
- `🧶` → `[Yarn]`
- `🎓` → `[Graduation Cap]`
- `🇺🇸` → `[Flag US]`
- `1️⃣` → `[1]`

This means newly encountered icons in future books will still be normalized into KDP-friendlier plain text instead of requiring you to keep adding manual mappings one by one.

Code spans and code blocks keep their code styling; when KDP-safe mode is enabled, icon glyphs inside them are also converted to ASCII-safe labels so the final DOCX avoids unsupported symbols more completely.

### Enable KDP-safe replacements

```bash
python convert.py input.md -o output.docx --kdp-safe-icons
```

### Provide your own mapping JSON

Create a UTF-8 JSON file such as `icon_map.json`:

```json
{
  "✅": "[Approved]",
  "🔗": "(URL)",
  "💡": "Tip:"
}
```

Then run:

```bash
python convert.py input.md -o output.docx --kdp-safe-icons --icon-map icon_map.json
```

`--icon-map` requires `--kdp-safe-icons`.
Custom mappings override the built-in generic semantics when both are provided.

## Setup

```bash
# Create and activate virtual environment (already done if .venv exists)
python -m venv .venv
.venv\Scripts\activate      # Windows
source .venv/bin/activate   # macOS/Linux

# Install dependencies
pip install -r requirements.txt
```

## Usage

```bash
# Convert a single file (output: same name with .docx extension)
python convert.py README.md

# Specify output path
python convert.py input.md -o output.docx

# Convert all .md files in a directory
python convert.py docs/

# Use a custom .docx template for styles/fonts/branding
python convert.py input.md -o output.docx --template my_template.docx

# Override source encoding if needed
python convert.py input.md -o output.docx --encoding cp950
python convert.py input.md -o output.docx --encoding shift_jis
python convert.py input.md -o output.docx --encoding cp1252

# Replace common emoji/icons with KDP-safe text
python convert.py input.md -o output.docx --kdp-safe-icons

# Merge the built-in KDP replacements with your own JSON overrides
python convert.py input.md -o output.docx --kdp-safe-icons --icon-map icon_map.json
```

## Project Structure

```
md_to_docx/
├── convert.py          # CLI entry point
├── requirements.txt    # Python dependencies
├── test.md             # Example Markdown for testing
├── md_to_docx/
│   ├── __init__.py
│   ├── renderer.py     # AST → DOCX converter
│   └── styles.py       # Word style definitions + multilingual fonts
└── .venv/              # Virtual environment
```

## How It Works

1. **Parse** — [mistune](https://github.com/lepture/mistune) parses the Markdown into a token AST with full GFM plugin support (tables, strikethrough, task lists, auto-links).
2. **Decode** — `convert.py` detects BOMs, tries UTF-8 first, uses charset detection, and falls back to common Traditional Chinese / Japanese / Western encodings.
3. **Render** — The AST walker in `renderer.py` maps each token type to the appropriate `python-docx` API calls.
4. **Style** — `styles.py` defines custom Word paragraph/character styles and multilingual font mappings for Latin and CJK text.
