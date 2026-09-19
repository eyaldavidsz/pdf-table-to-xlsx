"""
pdf_to_xlsx.py

Takes a PDF containing a gridlined table (a blank form/template someone
sent you to fill in) and produces a .xlsx spreadsheet with the same
headers, row labels, and layout. Handles right-to-left text (e.g. Hebrew)
correctly, including cells that mix Hebrew with numbers or other scripts,
and correctly displays right-to-left cells in the output spreadsheet.

Usage:
    python3 pdf_to_xlsx.py input.pdf output.xlsx
"""

import sys
import pdfplumber
from bidi.algorithm import get_display
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

MIRRORED_BRACKETS = {
    "(": ")", ")": "(",
    "[": "]", "]": "[",
    "{": "}", "}": "{",
    "<": ">", ">": "<",
}


def mirror_brackets(text):
    return "".join(MIRRORED_BRACKETS.get(ch, ch) for ch in text)


def contains_hebrew(text):
    return any("\u0590" <= ch <= "\u05FF" for ch in text)

def assign_words_to_cells(page, table):
    """
    Get every word on the page once, and assign each one to whichever
    table cell it overlaps the most (reusing the same overlap_area idea
    used for annotations below). More robust than requiring a word be
    fully contained in a cell (which can silently drop text that pokes
    out even slightly) or accepting any overlap at all (which can pick
    up stray slivers bleeding in from a neighboring cell).
    """
    cell_words = {}

    for word in page.extract_words():
        word_bbox = (word["x0"], word["top"], word["x1"], word["bottom"])
        best_position = None
        best_area = 0

        for row_index, row in enumerate(table.rows):
            for col_index, cell_bbox in enumerate(row.cells):
                if cell_bbox is None:
                    continue
                area = overlap_area(word_bbox, cell_bbox)
                if area > best_area:
                    best_area = area
                    best_position = (row_index, col_index)

        if best_position is not None:
            cell_words.setdefault(best_position, []).append(word)

    return cell_words

def join_line_words(line_words, gap_threshold=1.0):
    """
    Join the words on one line into a single string, deciding whether to
    insert a space based on the actual physical gap between each pair of
    words -- not on whatever space characters pdfplumber happened to
    report. Real word gaps measured well above 1pt; a PDF-internal
    "phantom space" artifact measured at ~0 (even overlapping).
    """
    parts = []
    previous_word = None
    for word in line_words:
        if previous_word is not None:
            gap = word["x0"] - previous_word["x1"]
            if gap > gap_threshold:
                parts.append(" ")
        parts.append(word["text"])
        previous_word = word
    return "".join(parts)

def words_to_cell_text(words):
    if not words:
        return ""

    lines = {}
    for word in words:
        line_key = round(word["top"])
        lines.setdefault(line_key, []).append(word)

    ordered_lines = []
    for key in sorted(lines.keys()):
        line_words = sorted(lines[key], key=lambda w: w["x0"])
        ordered_lines.append(line_words)

    fixed_lines = [
        get_display(mirror_brackets(join_line_words(line)))
        for line in ordered_lines
    ]
    return "\n".join(fixed_lines)

def overlap_area(rect1, rect2):
    x0 = max(rect1[0], rect2[0])
    x1 = min(rect1[2], rect2[2])
    top = max(rect1[1], rect2[1])
    bottom = min(rect1[3], rect2[3])
    if x1 <= x0 or bottom <= top:
        return 0
    return (x1 - x0) * (bottom - top)

def extract_text_annotations(page):
    """
    Text typed into a PDF via Preview's Markup/Form-filling tools lives in
    a separate "annotations" layer, not in the page's regular content --
    so it's invisible to normal table-text extraction entirely. This pulls
    annotation that actually has typed text in it.
    """
    annotations = []
    for annot in page.annots or []:
        contents = annot.get("contents")
        if contents:
            bbox = (annot["x0"], annot["top"], annot["x1"], annot["bottom"])
            annotations.append((bbox, contents))
    return annotations

def apply_annotations(page, table, rows):
    """
    For each text annotation on the page, find whichever table cell it
    overlaps the most (by area) and replace that cell's value with the
    annotation's text -- reproducing what you actually see when the PDF
    is opened, instead of the blank cell pdfplumber's normal text reading
    would otherwise report.
    """
    for annot_bbox, text in extract_text_annotations(page):
        best_position = None
        best_area = 0

        for row_index, row in enumerate(table.rows):
            for col_index, cell_bbox in enumerate(row.cells):
                if cell_bbox is None:
                    continue
                area = overlap_area(annot_bbox, cell_bbox)
                if area > best_area:
                    best_area = area
                    best_position = (row_index, col_index)

        if best_position is not None:
            row_index, col_index = best_position
            rows[row_index][col_index] = text

    return rows

def extract_table(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]
        tables = first_page.find_tables()
        if not tables:
            raise ValueError("No table with visible gridlines found on the first page.")
        table = tables[0]

        cell_words = assign_words_to_cells(first_page, table)

        rows = []
        for row_index, row in enumerate(table.rows):
            row_values = []
            for col_index, cell_bbox in enumerate(row.cells):
                if cell_bbox is None:
                    row_values.append(None)
                else:
                    words = cell_words.get((row_index, col_index), [])
                    row_values.append(words_to_cell_text(words))
            rows.append(row_values)

        rows = apply_annotations(first_page, table, rows)

    return rows

def to_cell_value(text):
    """
    Convert genuinely numeric extracted text into a real Python number, so
    Excel/Numbers treats it (and aligns it) the same way it would a number
    you typed in yourself, rather than as plain text.

    Whole numbers only convert if doing so provably loses nothing -- the
    round trip back to text has to match the original exactly. This
    protects things like a leading-zero code ("007") that would otherwise
    silently lose its zeros by becoming the number 7. Decimals convert
    more freely, since a difference like "42.50" becoming 42.5 is just
    numeric formatting, not lost meaning.
    """
    try:
        as_int = int(text)
        if str(as_int) == text:
            return as_int
    except (ValueError, TypeError):
        pass

    if "." in text:
        try:
            return float(text)
        except (ValueError, TypeError):
            pass

    return text


def write_spreadsheet(rows):
    wb = Workbook()
    sheet = wb.active
    for row_index, row in enumerate(rows, start=1):
        for col_index, value in enumerate(row, start=1):
            cell_value = to_cell_value(value) if value is not None else ""
            sheet.cell(row=row_index, column=col_index, value=cell_value)
    return wb

def style_spreadsheet(wb):
    sheet = wb.active
    header_font = Font(bold=True)
    rtl_alignment = Alignment(horizontal="right", readingOrder=2, wrap_text=True)
    wrap_alignment = Alignment(wrap_text=True)
    column_widths = {}

    for row_index, row in enumerate(sheet.iter_rows(), start=1):
        for cell in row:
            if row_index == 1:
                cell.font = header_font

            value = cell.value
            if value and contains_hebrew(str(value)):
                cell.alignment = rtl_alignment
            else:
                cell.alignment = wrap_alignment

            length = len(str(value)) if value else 0
            current_max = column_widths.get(cell.column_letter, 0)
            column_widths[cell.column_letter] = max(current_max, length)

    for column_letter, longest in column_widths.items():
        sheet.column_dimensions[column_letter].width = longest + 4

    return wb

def merge_spanned_cells(wb, rows):
    """
    rows still has None exactly where the PDF's table had a merged cell
    (see extract_table, where a cell's bbox being None is what produces
    this). For each row, find every run of one real
    value followed by one-or-more Nones, and merge that same span in the
    spreadsheet -- reproducing the PDF's original merged layout.
    """
    sheet = wb.active

    for row_index, row in enumerate(rows, start=1):
        col = 0
        while col < len(row):
            if row[col] is None:
                col += 1
                continue

            start_col = col
            end_col = col
            while end_col + 1 < len(row) and row[end_col + 1] is None:
                end_col += 1

            if end_col > start_col:
                sheet.merge_cells(
                    start_row=row_index, start_column=start_col + 1,
                    end_row=row_index, end_column=end_col + 1,
                )

            col = end_col + 1

    return wb

def build_spreadsheet(rows, output_path):
    wb = write_spreadsheet(rows)
    style_spreadsheet(wb)
    merge_spanned_cells(wb, rows)
    wb.save(output_path)

def main():
    if len(sys.argv) != 3:
        print("Usage: python3 pdf_to_xlsx.py input.pdf output.xlsx")
        sys.exit(1)

    pdf_path, output_path = sys.argv[1], sys.argv[2]
    rows = extract_table(pdf_path)
    build_spreadsheet(rows, output_path)

    print(f"Extracted {len(rows)} rows x {len(rows[0])} columns.")
    print(f"Saved to {output_path}")

if __name__ == "__main__":
    main()