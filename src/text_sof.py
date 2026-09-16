"""Text-only SOForm extraction. Embedded images and scanned pages are not read."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from hanger_automation import AutomationError, clean


MAX_TEXT_LINES = 20_000
MAX_TEXT_CHARS = 1_000_000


@dataclass(frozen=True)
class TextSof:
    sections: dict[str, tuple[str, ...]]
    format: str

    def cited_text(self, section: str, lines: str) -> str:
        if section not in self.sections:
            raise AutomationError(f"SOF section does not exist: {section}")
        match = re.fullmatch(r"L(\d{4,5})(?::L(\d{4,5}))?", lines)
        if not match:
            raise AutomationError(f"Invalid SOF line citation: {lines}")
        first = int(match.group(1))
        last = int(match.group(2) or match.group(1))
        if first < 1 or last < first or last > len(self.sections[section]) or last - first > 11:
            raise AutomationError(f"SOF line citation is out of range: {section} {lines}")
        return " ".join(self.sections[section][first - 1:last])

    def evidence(self, max_chars: int) -> str:
        chunks: list[str] = []
        length = 0
        for section, lines in self.sections.items():
            header = f"\n[SECTION: {section}]\n"
            if length + len(header) > max_chars:
                break
            chunks.append(header)
            length += len(header)
            for index, line in enumerate(lines, 1):
                entry = f"L{index:04d} = {line}\n"
                if length + len(entry) > max_chars:
                    chunks.append("[EVIDENCE TRUNCATED]\n")
                    return "".join(chunks)
                chunks.append(entry)
                length += len(entry)
        return "".join(chunks)


def _normalized_lines(values: list[str], filename: str) -> tuple[str, ...]:
    lines = tuple(clean(value) for value in values if clean(value))
    if len(lines) > MAX_TEXT_LINES or sum(len(line) for line in lines) > MAX_TEXT_CHARS:
        raise AutomationError(f"SOF {filename} contains too much text")
    return lines


def _read_pdf(path: Path) -> TextSof:
    try:
        import pymupdf

        sections: dict[str, tuple[str, ...]] = {}
        with pymupdf.open(path) as document:
            if document.is_encrypted:
                raise AutomationError(f"SOF {path.name} is password-protected")
            if len(document) == 0 or len(document) > 200:
                raise AutomationError(f"SOF {path.name} has an unsupported page count")
            for number, page in enumerate(document, 1):
                blocks = page.get_text("blocks", sort=True)
                values = [line for block in blocks if block[6] == 0
                          for line in block[4].splitlines()]
                lines = _normalized_lines(values, path.name)
                # Reject image-only pages, even when other pages contain text:
                # crucial hanger instructions could otherwise be silently missed.
                if not lines and (page.get_images(full=True) or any(block[6] == 1 for block in blocks)):
                    raise AutomationError(
                        f"SOF {path.name}, page {number} has no extractable text "
                        "(scan/image); OCR is disabled"
                    )
                if lines:
                    sections[f"PAGE {number}"] = lines
                if sum(len(line) for values in sections.values() for line in values) > MAX_TEXT_CHARS:
                    raise AutomationError(f"SOF {path.name} contains too much text")
    except AutomationError:
        raise
    except Exception as exc:
        raise AutomationError(f"SOF {path.name} is not a readable text PDF: {type(exc).__name__}") from exc
    if not sections:
        raise AutomationError(f"SOF {path.name} has no extractable text; OCR is disabled")
    return TextSof(sections, "pdf")


def _read_docx(path: Path) -> TextSof:
    try:
        from docx import Document
        from docx.table import Table

        document = Document(path)
        values: list[str] = []
        for block in document.iter_inner_content():
            if isinstance(block, Table):
                for row in block.rows:
                    values.append(" | ".join(clean(cell.text) for cell in row.cells))
            else:
                values.append(block.text)
        lines = _normalized_lines(values, path.name)
    except AutomationError:
        raise
    except Exception as exc:
        raise AutomationError(f"SOF {path.name} is not a readable .docx: {type(exc).__name__}") from exc
    if not lines:
        raise AutomationError(f"SOF {path.name} has no extractable text; OCR is disabled")
    return TextSof({"DOCUMENT": lines}, "docx")


def load_text_sof(path: Path) -> TextSof:
    suffix = path.suffix.casefold()
    if suffix == ".pdf":
        return _read_pdf(path)
    if suffix == ".docx":
        return _read_docx(path)
    raise AutomationError("Text SOF must be .pdf or .docx")
