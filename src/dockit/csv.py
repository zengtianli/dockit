"""CSV/TXT processing — conversion, delimiter detection, text normalization.

Pure logic module: string in, string out. No file I/O.

Usage:
    from dockit.csv import txt_to_csv, csv_to_txt, detect_delimiter

    delimiter = detect_delimiter(raw_text)
    csv_string = txt_to_csv(raw_text)
    txt_string = csv_to_txt(csv_string, delimiter="\\t")
"""

import csv as _csv
import re
from io import StringIO


def detect_delimiter(content: str) -> str:
    """Detect the most likely delimiter in text content.

    Checks for tab, comma, semicolon, and pipe characters.

    Args:
        content: Text content to analyze.

    Returns:
        The most frequently occurring delimiter character.
    """
    delimiters = ["\t", ",", ";", "|"]
    counts = {d: content.count(d) for d in delimiters}
    return max(counts, key=counts.get)


def txt_to_csv(content: str) -> str:
    """Convert delimited text to CSV format.

    Auto-detects the delimiter and converts to standard CSV.

    Args:
        content: Raw text content with some delimiter.

    Returns:
        CSV-formatted string.
    """
    delimiter = detect_delimiter(content)
    buf = StringIO()
    writer = _csv.writer(buf)
    for line in content.splitlines():
        if line.strip():
            writer.writerow(line.split(delimiter))
    return buf.getvalue()


def csv_to_txt(content: str, delimiter: str = "\t") -> str:
    """Convert CSV to delimited text.

    Args:
        content: CSV-formatted string.
        delimiter: Output delimiter (default: tab).

    Returns:
        Delimited text string.
    """
    buf = StringIO()
    reader = _csv.reader(StringIO(content))
    for row in reader:
        buf.write(delimiter.join(row) + "\n")
    return buf.getvalue()


def merge_texts(texts: list[str]) -> str:
    """Merge multiple text files into a single CSV.

    Files are merged column-wise (side by side), sorted by any leading
    number in the content identifier.

    Args:
        texts: List of text contents to merge column-wise.

    Returns:
        CSV-formatted string with columns from each text.
    """
    # Parse each text into rows
    all_rows = []
    max_lines = 0
    for text in texts:
        lines = [line.strip() for line in text.splitlines()]
        all_rows.append(lines)
        max_lines = max(max_lines, len(lines))

    if max_lines == 0:
        return ""

    # Merge column-wise
    merged = []
    for i in range(max_lines):
        row = []
        for lines in all_rows:
            row.append(lines[i] if i < len(lines) else "")
        merged.append(row)

    buf = StringIO()
    writer = _csv.writer(buf)
    writer.writerows(merged)
    return buf.getvalue()


def format_circles(content: str) -> str:
    """Replace circled numbers with plain numbered list format.

    Converts characters like \u2460 \u2461 \u2462 to 1. 2. 3. etc.

    Args:
        content: Text containing circled number characters.

    Returns:
        Text with circled numbers replaced.
    """
    circle_map = {
        "\u2460": "1.", "\u2461": "2.", "\u2462": "3.", "\u2463": "4.", "\u2464": "5.",
        "\u2465": "6.", "\u2466": "7.", "\u2467": "8.", "\u2468": "9.", "\u2469": "10.",
        "\u246a": "11.", "\u246b": "12.", "\u246c": "13.", "\u246d": "14.", "\u246e": "15.",
    }
    result = content
    for circle, number in circle_map.items():
        result = result.replace(circle, number)
    return result


def reorder_rows(csv_content: str, order_list: list[str]) -> str:
    """Reorder CSV rows based on an ordered list of first-column values.

    Rows matching the order list come first (in order), followed by
    any unmatched rows.

    Args:
        csv_content: CSV-formatted string.
        order_list: Ordered list of first-column values.

    Returns:
        Reordered CSV string.
    """
    reader = _csv.reader(StringIO(csv_content))
    rows = list(reader)

    data_dict = {row[0].strip(): row for row in rows if row}
    found_keys = set()
    ordered = []

    for name in order_list:
        clean = name.strip()
        if clean in data_dict:
            ordered.append(data_dict[clean])
            found_keys.add(clean)

    # Append unmatched rows
    for key, row in data_dict.items():
        if key not in found_keys:
            ordered.append(row)

    buf = StringIO()
    writer = _csv.writer(buf)
    writer.writerows(ordered)
    return buf.getvalue()
