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
- Correctly handles right-to-left text (e.g. Arabic, Hebrew), including cells that
  mix it with numbers or other scripts.
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

The pipeline has four conceptual stages, each solving one problem:

1. **Find the table and read its cells.** Locate a gridlined table on the
   PDF's first page, then figure out each cell's actual text — correcting
   for a couple of real-world PDF quirks along the way (space characters
   that don't correspond to real gaps between words; text that visually
   spans multiple grid cells). Right-to-left text is reordered into correct
   reading order. Text added via PDF annotations (e.g. Preview's Markup
   tool), which lives outside the page's normal text entirely, is matched
   to whichever cell it visually overlaps and placed there.
2. **Write the values into a spreadsheet.** Nothing clever here — just
   placing each extracted value into the matching cell of a new workbook.
3. **Style it.** Bold the header row, auto-size columns, and make sure
   right-to-left cells are actually displayed right-to-left rather than
   guessed at (Excel's own guess can be wrong depending on what character
   a cell happens to start with).
4. **Reproduce merged cells.** Any cell that visually spanned multiple
   columns in the original PDF is merged the same way in the output.

For the actual technique behind each of these, the code itself is the
source of truth — every non-obvious decision has a comment explaining why,
right where the decision was made.

## Known limitations

- Only the first page of the PDF is scanned, and only the first table found
  on it.
- Requires the table to have visible gridlines — tables formed only by
  aligned whitespace (no drawn borders) aren't detected.
- The "phantom space" fix uses a fixed gap threshold (1pt) tuned against a
  real-world example; a PDF from a different generator with unusual letter
  spacing could in principle need this adjusted.

## Credits

Built collaboratively with Claude Sonnet 5 (Anthropic) — I contributed to the design decisions (what
to extract, how to handle each bug, where each piece of logic should live), reviewed, and tested the code. Claude wrote the code and helped diagnose bugs.