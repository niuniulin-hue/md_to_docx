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
| Optional PDF export | ✅ |
| WeasyPrint Markdown+CSS PDF backend | ✅ |

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

## PDF Export

You can now ask the converter to produce a **PDF** in addition to the generated DOCX.

- Default behavior stays the same: without `--pdf`, it only writes a `.docx`
- With `--pdf`, it first generates the styled `.docx`, then exports that document to `.pdf`
- `--pdf-backend auto` now tries **LibreOffice/OpenOffice first** when available, then falls back to **Microsoft Word automation** on Windows
- You can force a backend explicitly with `--pdf-backend libreoffice`, `--pdf-backend word`, or `--pdf-backend weasyprint`
- If LibreOffice is installed outside `PATH`, you can point to it with `--soffice-path`
- `--pdf-backend weasyprint` uses **Markdown -> HTML -> WeasyPrint -> PDF** and supports additional CSS via `--pdf-css`

### Examples

```bash
# Generate README.docx and README.pdf side-by-side
python convert.py README.md --pdf

# Pick the final PDF path explicitly; the intermediate DOCX will use the same stem
python convert.py README.md -o dist\README-kdp.pdf --pdf

# If you pass a .docx output while using --pdf, the PDF will use the same stem
python convert.py README.md -o dist\README-kdp.docx --pdf

# Shorten the watchdog if Word export is hanging in your environment
python convert.py README.md --pdf --pdf-timeout 30

# Force LibreOffice/OpenOffice instead of Word
python convert.py README.md --pdf --pdf-backend libreoffice

# Point to a specific LibreOffice executable
python convert.py README.md --pdf --pdf-backend libreoffice --soffice-path "C:\Program Files\LibreOffice\program\soffice.com"

# High-quality print layout via WeasyPrint + CSS
python convert.py README.md --pdf --pdf-backend weasyprint --pdf-css styles\book.css

# Multiple CSS files are allowed; pass --pdf-css repeatedly
python convert.py README.md --pdf --pdf-backend weasyprint --pdf-css styles\base.css --pdf-css styles\kdp.css

# Resolve relative images/fonts from a custom base path in WeasyPrint mode
python convert.py README.md --pdf --pdf-backend weasyprint --pdf-base-url "C:\Code\md_to_docx"
```

### WeasyPrint on Windows — GTK3 setup (one-time)

`pip install weasyprint` only installs the Python package. WeasyPrint also needs the **GTK3 system libraries** (`libgobject`, `libpango`, `libcairo`) which must be on your `PATH`.

**Option A — standalone installer (simplest)**

1. Download the latest `gtk3-runtime-*-x64.exe` from:  
   <https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases>
2. Run the installer → tick **"Add to PATH for all users"** (or current user).
3. Close and reopen your terminal so the new `PATH` takes effect.
4. Retry:
   ```powershell
   python convert.py book.md --pdf --pdf-backend weasyprint
   ```

**Option B — MSYS2** (if already installed)

```bash
# In an MSYS2 MinGW64 shell:
pacman -S mingw-w64-x86_64-gtk3 mingw-w64-x86_64-python-weasyprint
```

Then add `C:\msys64\mingw64\bin` to your Windows `PATH` and retry from a normal terminal.

> Full WeasyPrint installation guide: <https://doc.courtbouillon.org/weasyprint/stable/first_steps.html#windows>

When `--pdf` is enabled:

- `python convert.py book.md --pdf` writes `book.docx` and `book.pdf`
- `python convert.py book.md -o final.pdf --pdf` writes `final.docx` and `final.pdf`
- `python convert.py books\ --pdf` writes both `.docx` and `.pdf` for each Markdown file in the folder tree
- In `--pdf-backend weasyprint` mode, PDF is generated directly from Markdown+HTML+CSS (DOCX generation is skipped)

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

# Also export PDF via Microsoft Word automation
python convert.py input.md --pdf

# Choose the final PDF filename explicitly
python convert.py input.md -o output.pdf --pdf

# Fail faster if Word gets stuck on a hidden export dialog
python convert.py input.md --pdf --pdf-timeout 30

# Force a specific PDF backend
python convert.py input.md --pdf --pdf-backend libreoffice
python convert.py input.md --pdf --pdf-backend word
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
5. **Export PDF (optional)** — if `--pdf` is enabled, the generated DOCX is exported via the selected backend. In `auto` mode the converter tries LibreOffice/OpenOffice first when found, then Word on Windows.
