"""Markdown processing — table name checking and standardization.

Pure logic module: string in, string out. No file I/O.

Usage:
    from dockit.md import check_tables, fix_table_names, reorder_table_names

    result = check_tables(md_text)
    fixed = fix_table_names(md_text, chapter_num=1)
    reordered, count = reorder_table_names(md_text)
"""

import re
from dataclasses import dataclass, field

# -- Constants ----------------------------------------------------------------

TABLE_NAME_RE = re.compile(r"^表\d+")
TABLE_NAME_FULL_RE = re.compile(r"^表\d+-\d+\s")
TABLE_SEP_RE = re.compile(r"^\|[\s\-:]+(\|[\s\-:]+)+\|?\s*$")


# -- Internal helpers ---------------------------------------------------------


def _is_in_code_block(lines: list[str], line_idx: int) -> bool:
    """Check if a line is inside a fenced code block."""
    in_code = False
    for i in range(line_idx):
        if lines[i].strip().startswith("```"):
            in_code = not in_code
    return in_code


def _find_tables(lines: list[str]) -> list[dict]:
    """Find all markdown tables (header row + separator row)."""
    tables = []
    for i in range(len(lines) - 1):
        if _is_in_code_block(lines, i):
            continue
        stripped = lines[i].strip()
        next_stripped = lines[i + 1].strip() if i + 1 < len(lines) else ""
        if stripped.startswith("|") and stripped.endswith("|"):
            if TABLE_SEP_RE.match(next_stripped):
                # Skip if previous line is also a table row (this is a continuation)
                if i > 0:
                    prev = lines[i - 1].strip()
                    if prev.startswith("|") and prev.endswith("|"):
                        continue
                tables.append({"header_line": i})
    return tables


def _check_table_name(lines: list[str], table: dict) -> dict | None:
    """Check if a table has a name (e.g., '表1-1 ...')."""
    header_line = table["header_line"]
    j = header_line - 1
    search_limit = max(0, header_line - 10)
    while j >= search_limit:
        stripped = lines[j].strip()
        if not stripped:
            j -= 1
            continue
        if stripped.startswith("#"):
            break
        if stripped.startswith("|") and stripped.endswith("|"):
            break
        if TABLE_NAME_RE.match(stripped):
            table["name_line"] = j
            table["has_name"] = True
            return None
        j -= 1
    table["has_name"] = False
    return {"type": "missing_name", "line": header_line + 1}


def _check_table_intro(lines: list[str], table: dict, min_chars: int = 80) -> dict | None:
    """Check if a table has a sufficient intro paragraph."""
    header_line = table["header_line"]
    name_line = table.get("name_line")
    best_intro = ""

    # Check between table name and header
    if name_line is not None and header_line - name_line > 2:
        for j in range(name_line + 1, header_line):
            stripped = lines[j].strip()
            if not stripped or stripped.startswith("|") or stripped.startswith("#") or stripped.startswith(">"):
                continue
            clean = re.sub(r"[*#>\[\]`]", "", stripped)
            if len(clean) > len(best_intro):
                best_intro = clean

    # Check above table name / header
    start = (name_line if name_line is not None else header_line) - 1
    j = start
    while j >= 0 and lines[j].strip() == "":
        j -= 1
    if j >= 0:
        stripped = lines[j].strip()
        if stripped.startswith(">"):
            k = j - 1
            while k >= 0 and (lines[k].strip().startswith(">") or lines[k].strip() == ""):
                k -= 1
            stripped = lines[k].strip() if k >= 0 else ""
        if not stripped.startswith("#") and not (stripped.startswith("|") and stripped.endswith("|")):
            clean = re.sub(r"[*#>\[\]`]", "", stripped)
            if len(clean) > len(best_intro):
                best_intro = clean

    if len(best_intro) == 0:
        return {"type": "missing_intro", "line": header_line + 1, "intro_len": 0}
    if len(best_intro) < min_chars:
        return {
            "type": "short_intro",
            "line": header_line + 1,
            "intro_len": len(best_intro),
            "min_chars": min_chars,
        }
    return None


# -- Public API ---------------------------------------------------------------


@dataclass
class TableCheckResult:
    """Result of table checking."""

    tables: int = 0
    issues: list[dict] = field(default_factory=list)


def check_tables(text: str, *, min_intro_chars: int = 80) -> TableCheckResult:
    """Check markdown tables for missing names and intro paragraphs.

    Args:
        text: Markdown text content.
        min_intro_chars: Minimum characters for intro paragraph.

    Returns:
        TableCheckResult with table count and list of issues found.
        Each issue has: type ("missing_name", "missing_intro", "short_intro"), line number.
    """
    lines = text.split("\n")
    tables = _find_tables(lines)
    result = TableCheckResult(tables=len(tables))

    for table in tables:
        name_issue = _check_table_name(lines, table)
        if name_issue:
            result.issues.append(name_issue)
        intro_issue = _check_table_intro(lines, table, min_intro_chars)
        if intro_issue:
            result.issues.append(intro_issue)

    return result


def fix_table_names(text: str, chapter_num: int, *, min_intro_chars: int = 80) -> str:
    """Insert placeholder table names and intro markers where missing.

    Args:
        text: Markdown text content.
        chapter_num: Chapter number for table numbering (e.g., 1 produces 表1-1, 表1-2).
        min_intro_chars: Minimum characters for intro paragraph.

    Returns:
        Modified text with placeholder table names inserted.
    """
    lines = text.split("\n")
    tables = _find_tables(lines)
    insertions = []

    table_counter = 0
    for table in tables:
        table_counter += 1
        name_issue = _check_table_name(lines, table)
        intro_issue = _check_table_intro(lines, table, min_intro_chars)

        if name_issue:
            table_name = f"表{chapter_num}-{table_counter} [待命名]"
            insertions.append({"line": table["header_line"], "content": table_name, "type": "name"})

        if intro_issue:
            insert_before = table.get("name_line", table["header_line"])
            if name_issue:
                insert_before = table["header_line"]
            insertions.append({"line": insert_before, "content": "<!-- TABLE_NEEDS_INTRO -->", "type": "intro"})

    insertions.sort(key=lambda x: (-x["line"], x["type"] == "name"))

    for ins in insertions:
        idx = ins["line"]
        lines.insert(idx, "")
        lines.insert(idx, ins["content"])
        if idx > 0 and lines[idx - 1].strip() != "":
            lines.insert(idx, "")

    return "\n".join(lines)


def reorder_table_names(text: str) -> tuple[str, int]:
    """Move table names to just above their tables (after intro paragraphs).

    When a table name appears before an intro paragraph, reorder so the
    intro comes first, then the table name, then the table.

    Args:
        text: Markdown text content.

    Returns:
        Tuple of (modified text, number of fixes applied).
    """
    lines = text.split("\n")
    fixes = 0
    i = 0

    while i < len(lines):
        stripped = lines[i].strip()
        if not TABLE_NAME_FULL_RE.match(stripped):
            i += 1
            continue

        name_line_idx = i
        name_line = lines[i]
        j = i + 1
        intro_lines = []
        found_table = False

        while j < len(lines):
            jstripped = lines[j].strip()
            if jstripped == "":
                j += 1
                continue
            if jstripped.startswith("|") and jstripped.endswith("|"):
                if j + 1 < len(lines) and TABLE_SEP_RE.match(lines[j + 1].strip()):
                    found_table = True
                    table_header_idx = j
                    break
                else:
                    intro_lines.append(lines[j])
                    j += 1
                    continue
            if jstripped.startswith("#"):
                break
            intro_lines.append(lines[j])
            j += 1

        if not found_table:
            i += 1
            continue

        if not intro_lines:
            lines[name_line_idx : table_header_idx + 1] = [name_line, lines[table_header_idx]]
            i = name_line_idx + 2
            continue

        new_segment = [*intro_lines, "", name_line, lines[table_header_idx]]
        lines[name_line_idx : table_header_idx + 1] = new_segment
        fixes += 1
        i = name_line_idx + len(new_segment)

    return "\n".join(lines), fixes
