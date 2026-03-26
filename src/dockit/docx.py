"""Word document processing — text formatting, quote font splitting.

Pure logic module: bytes in, bytes out. No file paths, no console output.

Usage:
    from dockit.docx import format_text

    with open("input.docx", "rb") as f:
        result = format_text(f.read())

    with open("output.docx", "wb") as f:
        f.write(result.data)

    print(result.stats)  # {"quotes": 5, "punctuation": 12, "units": 3}
"""

import copy
from dataclasses import dataclass, field
from io import BytesIO

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from dockit.text import fix_all


@dataclass
class FormatResult:
    """Result of a document formatting operation."""

    data: bytes
    stats: dict[str, int] = field(default_factory=dict)


# -- Internal helpers ----------------------------------------------------------

QUOTE_CHARS = {"\u201c", "\u201d"}


def _set_run_font(run_element, font_name: str):
    """Set font for a run element (ascii + hAnsi + eastAsia)."""
    rPr = run_element.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        run_element.insert(0, rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:hint"), "eastAsia")


def _split_run_at_quotes(run):
    """Split a run at quote characters, returning [(text, is_quote), ...] or None."""
    text = run.text
    if not text or not any(c in QUOTE_CHARS for c in text):
        return None

    segments = []
    buf = []
    for c in text:
        if c in QUOTE_CHARS:
            if buf:
                segments.append(("".join(buf), False))
                buf = []
            segments.append((c, True))
        else:
            buf.append(c)
    if buf:
        segments.append(("".join(buf), False))
    return segments


def _apply_quote_split(run, segments, quote_font: str):
    """Split run into multiple runs: quote chars get the specified font."""
    parent = run._element.getparent()

    # First segment reuses the original run
    first_text, first_is_quote = segments[0]
    run.text = first_text
    if first_is_quote:
        _set_run_font(run._element, quote_font)

    # Subsequent segments: deep-copy original run, insert after
    insert_after = run._element
    for seg_text, is_quote in segments[1:]:
        new_r = copy.deepcopy(run._element)
        # Clear old text elements from the copy
        for t_elem in new_r.findall(qn("w:t")):
            new_r.remove(t_elem)
        t_elem = OxmlElement("w:t")
        t_elem.text = seg_text
        t_elem.set(qn("xml:space"), "preserve")
        new_r.append(t_elem)

        if is_quote:
            _set_run_font(new_r, quote_font)
        else:
            # Non-quote segments: restore original font
            rPr = new_r.find(qn("w:rPr"))
            if rPr is not None:
                rFonts = rPr.find(qn("w:rFonts"))
                orig_rPr = run._element.find(qn("w:rPr"))
                orig_rFonts = orig_rPr.find(qn("w:rFonts")) if orig_rPr is not None else None
                if rFonts is not None and orig_rFonts is not None:
                    rPr.replace(rFonts, copy.deepcopy(orig_rFonts))
                elif rFonts is not None and orig_rFonts is None:
                    rPr.remove(rFonts)

        parent.insert(list(parent).index(insert_after) + 1, new_r)
        insert_after = new_r


def _process_paragraph(paragraph, stats: dict, quote_counter: int, quote_font: str) -> int:
    """Process a single paragraph: fix text and split quotes into separate runs.

    Returns updated quote_counter.
    """
    if not paragraph.runs:
        return quote_counter

    original_runs = list(paragraph.runs)

    for run in original_runs:
        if not run.text:
            continue

        original_text = run.text
        fixed_text, fix_stats, quote_counter = fix_all(original_text, quote_counter)

        stats["quotes"] += fix_stats["quotes"]
        stats["punctuation"] += fix_stats["punctuation"]
        stats["units"] += fix_stats["units"]

        if fixed_text != original_text:
            run.text = fixed_text

        # Split quotes into separate runs with specified font
        segments = _split_run_at_quotes(run)
        if segments:
            _apply_quote_split(run, segments, quote_font)

    return quote_counter


def _process_table(table, stats: dict, quote_counter: int, quote_font: str) -> int:
    """Process all paragraphs in a table. Returns updated quote_counter."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                quote_counter = _process_paragraph(paragraph, stats, quote_counter, quote_font)
    return quote_counter


# -- Public API ----------------------------------------------------------------


def format_text(
    doc_bytes: bytes,
    *,
    fix_quotes: bool = True,
    fix_punctuation: bool = True,
    fix_units: bool = True,
    quote_font: str = "\u5b8b\u4f53",
    process_headers_footers: bool = True,
    strip_headers_footers: bool = False,
) -> FormatResult:
    """Format text in a Word document.

    Fixes quotes (with font splitting), punctuation, and unit symbols
    across all paragraphs, tables, headers, and footers.

    Args:
        doc_bytes: Raw bytes of a .docx file.
        fix_quotes: Whether to fix quote characters.
        fix_punctuation: Whether to convert English punctuation to Chinese.
        fix_units: Whether to convert Chinese unit names to symbols.
        quote_font: Font name for quote characters (default: Song Ti).
        process_headers_footers: Whether to process headers/footers.
        strip_headers_footers: If True, remove all headers/footers entirely.

    Returns:
        FormatResult with processed document bytes and replacement stats.
    """
    doc = Document(BytesIO(doc_bytes))
    stats = {"quotes": 0, "punctuation": 0, "units": 0}
    quote_counter = 0

    # Process body paragraphs
    for paragraph in doc.paragraphs:
        quote_counter = _process_paragraph(paragraph, stats, quote_counter, quote_font)

    # Process tables
    for table in doc.tables:
        quote_counter = _process_table(table, stats, quote_counter, quote_font)

    # Process headers/footers
    if strip_headers_footers:
        for section in doc.sections:
            sectPr = section._sectPr
            for ref in sectPr.findall(qn("w:headerReference")):
                sectPr.remove(ref)
            for ref in sectPr.findall(qn("w:footerReference")):
                sectPr.remove(ref)
    elif process_headers_footers:
        for section in doc.sections:
            if section.header:
                for paragraph in section.header.paragraphs:
                    quote_counter = _process_paragraph(paragraph, stats, quote_counter, quote_font)
                for table in section.header.tables:
                    quote_counter = _process_table(table, stats, quote_counter, quote_font)
            if section.footer:
                for paragraph in section.footer.paragraphs:
                    quote_counter = _process_paragraph(paragraph, stats, quote_counter, quote_font)
                for table in section.footer.tables:
                    quote_counter = _process_table(table, stats, quote_counter, quote_font)

    # Save to bytes
    buf = BytesIO()
    doc.save(buf)
    return FormatResult(data=buf.getvalue(), stats=stats)
