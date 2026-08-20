# pdf-to-xlsx

Converts a PDF containing a gridlined table into a `.xlsx` spreadsheet with
the same headers, row labels, and layout — whether the PDF is a blank
template someone sent you to fill in, or one that already has some cells
filled in (including text added via Preview's Markup/Form-filling tools).

Built to avoid retyping table structures by hand, and to run entirely
locally — no PDF content is ever sent to a cloud service or third-party AI.

## Features

- Extracts a gridlined table from the first page of a PDF, preserving
  headers, row labels, and multi-line cells.
- Correctly handles right-to-left text (e.g. Hebrew), including cells that
  mix Hebrew with numbers or other scripts.
- Fixes a PDF-internal "phantom space" artifact that can otherwise split a
  single word into two.
- Reproduces merged cells from the original PDF as merged cells in the
  spreadsheet.
- Picks up text added to the PDF via Preview's Markup/Form-filling tools
  (stored as a separate "annotation," invisible to normal text extraction)
  and places it in the correct cell.

## Requirements

- Python 3
- [`pdfplumber`](https://github.com/jsvine/pdfplumber)
- [`openpyxl`](https://openpyxl.readthedocs.io/)
- [`python-bidi`](https://github.com/MeirKriheli/python-bidi)

## Setup

```bash
git clone <your-repo-url>
cd pdf-to-xlsx
python3 -m venv .venv
source .venv/bin/activate
pip install pdfplumber openpyxl python-bidi
```

## Usage

```bash
python3 pdf_to_xlsx.py input.pdf output.xlsx
```

### As a Finder Quick Action (macOS)

An Automator Quick Action can run this script directly from a PDF's
right-click menu, saving the `.xlsx` alongside the original PDF:

1. Automator → New → Quick Action.
2. Set "Workflow receives current" to **PDF files** in **Finder**.
3. Add a **Run Shell Script** action, shell `/bin/zsh`, "Pass input" set to
   **as arguments**.
4. Paste in a script that calls this tool with your venv's Python and
   derives the output filename from the input:

   ```bash
   for f in "$@"
   do
       dir=$(dirname "$f")
       base=$(basename "$f" .pdf)
       output="$dir/$base.xlsx"

       /path/to/pdf-to-xlsx/.venv/bin/python3 \
           /path/to/pdf-to-xlsx/pdf_to_xlsx.py \
           "$f" "$output"

       if [ $? -ne 0 ]; then
           osascript -e "display notification \"Failed to convert $(basename "$f")\" with title \"PDF to Spreadsheet\""
       fi
   done
   ```

5. Save it (e.g. as "Convert to Spreadsheet"). It'll now appear when you
   right-click any PDF in Finder.

## How it works

The pipeline is four small, independent stages:

1. **`extract_table`** — reads the PDF's first page, finds a table by its
   gridlines (via `pdfplumber`'s `find_tables`), and reads each cell's text
   by measuring the actual physical gaps between words rather than trusting
   pdfplumber's own space-character detection (which can be fooled by a PDF
   producing a misplaced space). Right-to-left text is corrected per line
   using the Unicode Bidirectional Algorithm (`python-bidi`). Separately,
   any text added via PDF annotations (e.g. Preview's Markup tool) is
   matched to whichever cell it physically overlaps most and placed there.
2. **`write_spreadsheet`** — writes the extracted values into a new
   workbook, one cell at a time.
3. **`style_spreadsheet`** — bolds the header row, right-aligns and sets
   correct reading order on any cell containing Hebrew (so Excel doesn't
   guess the wrong direction based on a cell's first character), and
   auto-sizes each column.
4. **`merge_spanned_cells`** — reproduces any merged cells from the original
   PDF table as merged cells in the spreadsheet.

## Known limitations

- Only the first page of the PDF is scanned, and only the first table found
  on it.
- Requires the table to have visible gridlines — tables formed only by
  aligned whitespace (no drawn borders) aren't detected.
- The "phantom space" fix uses a fixed gap threshold (1pt) tuned against a
  real-world example; a PDF from a different generator with unusual letter
  spacing could in principle need this adjusted.

## Credits

Built collaboratively with Claude Sonnet 5(Anthropic) — every design decision (what
to extract, how to handle each bug, where each piece of logic should live)
was mine; Claude wrote the implementing code and helped diagnose each bug
against the real PDF.
