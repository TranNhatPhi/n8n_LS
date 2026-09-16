"""Text-only order extraction. Word/PDF order files are read for the text they
state; embedded images and scanned pages are not read.

An .xls/.xlsx order keeps going through the workbook path unchanged. A .pdf or
.docx order is converted into a normalized order workbook carrying the standard
column headers first, so every later step of the pipeline is identical. Only
values the document actually states are written: a column the document does not
carry stays blank, which sends the row to REVIEW instead of guessing it. A
purchase order rarely names the Account or Hang/Flat, so this normally produces
a workbook that a human still has to complete - that is the intended outcome.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import Workbook

from hanger_automation import ALIASES, ORDER_FIELDS, AutomationError, clean, key


MAX_ORDER_ROWS = 5_000
MAX_ORDER_PAGES = 200
# Points. Words closer than this vertically are one line; header columns further
# apart than this horizontally are separate columns.
LINE_TOLERANCE = 2.0
COLUMN_GAP = 8.0
# Purchase orders stack a column title over several lines ("Qty" / "Ordered").
# A title line sits within this many lines and points of the main title line.
HEADER_BAND_LINES = 3
HEADER_BAND_POINTS = 20.0

# Column titles a purchase order uses for the order fields. Only titles that name
# the same thing are listed: a PO "Unit Price" or "Amount" column has no order
# field, so it is read as a column to ignore rather than guessed into one.
ORDER_COLUMN_ALIASES = dict(ALIASES) | {
    "qty": "PO Qty", "qty ordered": "PO Qty", "ordered": "PO Qty",
    "order qty": "PO Qty", "quantity": "PO Qty", "quantity ordered": "PO Qty",
    "total qty": "PO Qty",
    "style number": "Style", "style no": "Style", "style#": "Style",
    "style code": "Style",
    "description": "Product Description", "item description": "Product Description",
    "colour": "Color", "color name": "Color",
    "size": "Size Configuration", "sizes": "Size Configuration",
    "size config": "Size Configuration", "size breakdown": "Size Configuration",
    "po no": "PO#", "purchase order": "PO#", "purchase order number": "PO#",
    "customer po": "PO#",
    "account code": "Account", "account no": "Account", "customer account": "Account",
    "ref no": "Ref#", "reference": "Ref#",
    "hang flat": "Hang/Flat", "hang / flat": "Hang/Flat",
    "master box qty": "Master Box Quantity", "carton qty": "Master Box Quantity",
}

# A line item wraps onto a following line inside these columns only. A line that
# fills any other column starts a new item.
CONTINUATION_FIELDS = {"Product Description", "Color"}

# Totals and footers sit under the last line item in the same columns; reading
# them as an item would invent an order row.
FOOTER_PREFIXES = (
    "total", "subtotal", "sub total", "grand total", "comments", "gst",
    "freight", "balance", "amount due", "page ", "continued",
)

# Fields the pipeline needs to check a row. They are reported, never invented.
KEY_FIELDS = (
    "Account", "PO#", "Ref#", "Style", "Product Description", "Division",
    "Label", "Size Configuration", "Hang/Flat",
)

IDENTITY_FIELDS = {"PO#", "Account", "Ref#", "Style"}


@dataclass(frozen=True)
class TextOrder:
    """The normalized workbook plus what the document did and did not state."""

    path: Path
    format: str
    row_count: int
    fields: tuple
    missing_fields: tuple
    document_fields: dict
    notes: tuple


def _canonical_column(text: Any) -> str:
    """Return the order field a column title names, or "" for a column to ignore."""
    normalized = key(text).replace("\n", " ").rstrip(":")
    normalized = re.sub(r"\s+", " ", normalized).strip()
    return ORDER_COLUMN_ALIASES.get(normalized, "")


def _label_field(text: str) -> str:
    """Return the order field named by the last words before a "label:" colon."""
    words = clean(text).split()
    for size in (4, 3, 2, 1):
        if len(words) >= size:
            field = _canonical_column(" ".join(words[-size:]))
            if field:
                return field
    return ""


def _labelled_values(lines) -> tuple:
    """Read order-level "PO number: PO07564" statements from document text.

    A field stated twice with two different values is dropped rather than
    resolved, so a conflicting document cannot silently pick a winner.
    """
    found: dict = {}
    conflicting: set = set()
    for line in lines:
        for match in re.finditer(":", line):
            field = _label_field(line[:match.start()])
            if not field:
                continue
            tail = line[match.end():].strip()
            # Columnar layouts put the next label two or more spaces further on.
            value = clean(re.split(r"\s{2,}|:", tail)[0])
            if not value or len(value) > 80:
                continue
            if field in found and key(found[field]) != key(value):
                conflicting.add(field)
            found.setdefault(field, value)
    notes = []
    for field in sorted(conflicting):
        notes.append(f"{field} is stated more than once with different values; it was left blank")
        found.pop(field, None)
    return found, tuple(notes)


def _cluster_lines(words) -> list:
    """Group (x0, y0, x1, y1, word) tuples into lines of (x0, x1, word)."""
    lines: list = []
    for x0, y0, x1, y1, word in sorted(words, key=lambda item: (item[1], item[0])):
        if lines and abs(y0 - lines[-1][0]) <= LINE_TOLERANCE:
            lines[-1][1].append((x0, x1, word))
        else:
            lines.append((y0, [(x0, x1, word)]))
    return [(top, sorted(items)) for top, items in lines]


def _line_fields(items) -> dict:
    """Return {order field: (x0, x1)} for the column titles on one line."""
    found: dict = {}
    index = 0
    while index < len(items):
        for size in (3, 2, 1):
            if index + size > len(items):
                continue
            chunk = items[index:index + size]
            field = _canonical_column(" ".join(word for _, _, word in chunk))
            if field:
                left, right = chunk[0][0], chunk[-1][1]
                if field in found:
                    left, right = min(left, found[field][0]), max(right, found[field][1])
                found[field] = (left, right)
                index += size
                break
        else:
            index += 1
    return found


def _header_line_index(lines) -> int:
    """Return the index of the line that titles the line-item table."""
    best = -1
    best_score = 1
    for index, (_, items) in enumerate(lines):
        names = set(_line_fields(items))
        if not names & {"Style", "Product Description", "Ref#"}:
            continue
        if len(names) > best_score:
            best, best_score = index, len(names)
    return best


def _header_band(lines, header_index) -> tuple:
    """Return the title lines of the table and the index the items start after."""
    top = lines[header_index][0]
    first = max(0, header_index - HEADER_BAND_LINES)
    last = min(len(lines), header_index + HEADER_BAND_LINES + 1)
    indexes = [
        index for index in range(first, last)
        if _line_fields(lines[index][1]) and abs(lines[index][0] - top) <= HEADER_BAND_POINTS
    ]
    return [lines[index] for index in indexes], max(indexes)


def _columns_from_header(header_lines) -> list:
    """Merge the title words into one box per column, left to right."""
    boxes: list = []
    for _, items in header_lines:
        fields = _line_fields(items)
        for x0, x1, _word in items:
            name = ""
            for field, (left, right) in fields.items():
                if x0 >= left - 0.5 and x1 <= right + 0.5:
                    name = field
                    break
            boxes.append((x0, x1, name))
    groups: list = []
    for x0, x1, name in sorted(boxes):
        if groups and x0 <= groups[-1][1] + COLUMN_GAP:
            left, right, names = groups[-1]
            groups[-1] = (left, max(right, x1), names | ({name} if name else set()))
        else:
            groups.append((x0, x1, {name} if name else set()))
    columns = []
    for left, right, names in groups:
        if len(names) > 1:
            raise AutomationError(
                "Order columns overlap and cannot be told apart: " + ", ".join(sorted(names))
            )
        columns.append((left, right, next(iter(names)) if names else ""))
    return columns


def _split_line(items, columns) -> dict:
    """Assign every word of a line to one column.

    A word that overlaps a column title belongs to it - that is how a right
    aligned figure sits under its title. A word that overlaps nothing belongs to
    the last column that starts at or before it, which is how the wrapped second
    line of a description stays in the description column instead of drifting
    into the next one.
    """
    cells: dict = {}
    for x0, x1, word in items:
        chosen = None
        overlap = 0.0
        for index, (left, right, _field) in enumerate(columns):
            shared = min(x1, right) - max(x0, left)
            if shared > overlap:
                chosen, overlap = index, shared
        if chosen is None:
            starts = [index for index, (left, _r, _f) in enumerate(columns) if left <= x0]
            chosen = starts[-1] if starts else 0
        field = columns[chosen][2]
        if field:
            cells[field] = (cells.get(field, "") + " " + word).strip()
    return cells


def _page_records(lines, header_index, columns, filename: str) -> list:
    records: list = []
    header_bottom = lines[header_index][0]
    titles = {field for _left, _right, field in columns if field}
    for top, items in lines[header_index + 1:]:
        if top <= header_bottom:
            continue
        text = key(" ".join(word for _, _, word in items))
        if text.startswith(FOOTER_PREFIXES):
            break
        cells = _split_line(items, columns)
        if not cells:
            continue
        names = set(_line_fields(items))
        if len(names) >= 2 and names <= titles:
            # This table's own column titles, repeated; not a line item.
            continue
        if records and not set(cells) - CONTINUATION_FIELDS:
            for field, value in cells.items():
                records[-1][field] = (records[-1].get(field, "") + " " + value).strip()
            continue
        records.append(dict(cells))
        if len(records) > MAX_ORDER_ROWS:
            raise AutomationError(f"Order {filename} contains more than {MAX_ORDER_ROWS} line items")
    return records


def _read_pdf(path: Path) -> tuple:
    try:
        import pymupdf

        records: list = []
        text_lines: list = []
        notes: list = []
        with pymupdf.open(path) as document:
            if document.is_encrypted:
                raise AutomationError(f"Order {path.name} is password-protected")
            if len(document) == 0 or len(document) > MAX_ORDER_PAGES:
                raise AutomationError(f"Order {path.name} has an unsupported page count")
            for number, page in enumerate(document, 1):
                words = [(w[0], w[1], w[2], w[3], w[4]) for w in page.get_text("words") if w[4].strip()]
                if not words:
                    # Reject image-only pages, even when other pages contain text:
                    # whole line items could otherwise be silently missed.
                    if page.get_images(full=True):
                        raise AutomationError(
                            f"Order {path.name}, page {number} has no extractable text "
                            "(scan/image); OCR is disabled"
                        )
                    continue
                lines = _cluster_lines(words)
                text_lines.extend(" ".join(word for _, _, word in items) for _, items in lines)
                header_index = _header_line_index(lines)
                if header_index < 0:
                    notes.append(f"Page {number} carries no line-item table and was not read")
                    continue
                header_lines, last_header = _header_band(lines, header_index)
                columns = _columns_from_header(header_lines)
                records.extend(_page_records(lines, last_header, columns, path.name))
    except AutomationError:
        raise
    except Exception as exc:
        raise AutomationError(
            f"Order {path.name} is not a readable text PDF: {type(exc).__name__}"
        ) from exc
    return records, text_lines, notes


def _read_docx(path: Path) -> tuple:
    try:
        from docx import Document
        from docx.table import Table

        document = Document(path)
        records: list = []
        text_lines: list = []
        notes: list = []
        for block in document.iter_inner_content():
            if not isinstance(block, Table):
                text_lines.append(clean(block.text))
                continue
            rows = [[clean(cell.text) for cell in row.cells] for row in block.rows]
            header_index = -1
            columns: list = []
            for index, row in enumerate(rows):
                fields = [_canonical_column(value) for value in row]
                named = {field for field in fields if field}
                if len(named) >= 2 and named & {"Style", "Product Description", "Ref#"}:
                    header_index, columns = index, fields
                    break
            if header_index < 0:
                text_lines.extend(" | ".join(row) for row in rows)
                continue
            for row in rows[header_index + 1:]:
                if key(" ".join(row)).startswith(FOOTER_PREFIXES):
                    break
                cells = {}
                for field, value in zip(columns, row):
                    if field and value and not cells.get(field):
                        cells[field] = value
                if not cells:
                    continue
                if records and not set(cells) - CONTINUATION_FIELDS:
                    for field, value in cells.items():
                        records[-1][field] = (records[-1].get(field, "") + " " + value).strip()
                    continue
                records.append(cells)
                if len(records) > MAX_ORDER_ROWS:
                    raise AutomationError(
                        f"Order {path.name} contains more than {MAX_ORDER_ROWS} line items"
                    )
    except AutomationError:
        raise
    except Exception as exc:
        raise AutomationError(
            f"Order {path.name} is not a readable .docx: {type(exc).__name__}"
        ) from exc
    return records, text_lines, notes


def _quantity(value: str):
    """Return a plain integer for a quantity a PO prints as "1,224"."""
    text = clean(value).replace(",", "")
    if re.fullmatch(r"\d{1,9}", text):
        return int(text)
    return clean(value)


def write_order_workbook(records, target: Path) -> Path:
    """Write the normalized order workbook the rest of the pipeline reads."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ORDER"
    sheet.append(ORDER_FIELDS)
    sheet.freeze_panes = "A2"
    for record in records:
        row = []
        for field in ORDER_FIELDS:
            value = clean(record.get(field, ""))
            row.append(_quantity(value) if field == "PO Qty" else value)
        sheet.append(row)
    workbook.save(target)
    workbook.close()
    return target


def load_text_order(path: Path, temp_dir: Path) -> TextOrder:
    suffix = path.suffix.casefold()
    if suffix == ".pdf":
        records, text_lines, notes = _read_pdf(path)
        source_format = "pdf"
    elif suffix == ".docx":
        records, text_lines, notes = _read_docx(path)
        source_format = "docx"
    else:
        raise AutomationError("Text order must be .pdf or .docx")
    if not records:
        raise AutomationError(
            f"Order {path.name} has no readable line-item table; "
            "upload the order as .xls/.xlsx instead"
        )

    document_fields, conflict_notes = _labelled_values(text_lines)
    notes = list(notes) + list(conflict_notes)
    applied = {}
    for field, value in document_fields.items():
        if field in ORDER_FIELDS and not any(record.get(field) for record in records):
            applied[field] = value
    for record in records:
        for field, value in applied.items():
            record.setdefault(field, value)

    for record in records:
        for field in list(record):
            if field not in ORDER_FIELDS:
                record.pop(field)
    records = [record for record in records
               if any(clean(record.get(field, "")) for field in IDENTITY_FIELDS | {"Product Description"})]
    if not records:
        raise AutomationError(f"Order {path.name} has line items but none carry a style or description")

    filled = tuple(field for field in ORDER_FIELDS
                   if any(clean(record.get(field, "")) for record in records))
    missing = tuple(field for field in KEY_FIELDS if field not in filled)
    target = temp_dir / f"{path.stem}_order.xlsx"
    write_order_workbook(records, target)
    return TextOrder(
        path=target,
        format=source_format,
        row_count=len(records),
        fields=filled,
        missing_fields=missing,
        document_fields=applied,
        notes=tuple(notes),
    )
