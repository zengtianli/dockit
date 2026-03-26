"""Command-line interface for DocKit.

Usage:
    python3 -m dockit format-word <input_file> [-o output_file]
    python3 -m dockit convert <input_file> --to <csv|xlsx|txt> [-o output_file]
    python3 -m dockit standardize-ppt <input_file> [-o output_file]
"""

import argparse
import json
import sys
from pathlib import Path


def _default_output(input_path: Path, suffix: str = "_formatted", new_ext: str | None = None) -> Path:
    """Generate default output path: same directory, with suffix or new extension."""
    ext = new_ext if new_ext else input_path.suffix
    return input_path.with_name(f"{input_path.stem}{suffix}{ext}")


def _convert_output(input_path: Path, target_format: str) -> Path:
    """Generate default output path for convert: same name, new extension."""
    ext_map = {"csv": ".csv", "xlsx": ".xlsx", "txt": ".txt"}
    new_ext = ext_map[target_format]
    # If same extension, add _converted suffix to avoid overwriting
    if input_path.suffix.lower() == new_ext:
        return input_path.with_name(f"{input_path.stem}_converted{new_ext}")
    return input_path.with_suffix(new_ext)


def cmd_format_word(args: argparse.Namespace) -> int:
    """Handle the format-word subcommand."""
    from dockit.docx import format_text

    input_path = Path(args.input_file).resolve()
    if not input_path.exists():
        print(f"Error: file not found: {input_path}", file=sys.stderr)
        return 1

    if input_path.suffix.lower() not in (".docx",):
        print(f"Error: unsupported file type '{input_path.suffix}', expected .docx", file=sys.stderr)
        return 1

    output_path = Path(args.output).resolve() if args.output else _default_output(input_path)

    try:
        doc_bytes = input_path.read_bytes()
        result = format_text(
            doc_bytes,
            fix_quotes=True,
            fix_punctuation=True,
            fix_units=True,
        )
        output_path.write_bytes(result.data)
        print(json.dumps(result.stats, ensure_ascii=False))
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    return 0


def cmd_convert(args: argparse.Namespace) -> int:
    """Handle the convert subcommand."""
    input_path = Path(args.input_file).resolve()
    if not input_path.exists():
        print(f"Error: file not found: {input_path}", file=sys.stderr)
        return 1

    target = args.to
    input_ext = input_path.suffix.lower()
    output_path = Path(args.output).resolve() if args.output else _convert_output(input_path, target)

    try:
        if input_ext == ".xlsx" and target == "csv":
            from dockit.xlsx import xlsx_to_csv

            xlsx_bytes = input_path.read_bytes()
            csv_dict = xlsx_to_csv(xlsx_bytes)
            # If single sheet, write directly; if multiple, write first and report all
            sheet_names = list(csv_dict.keys())
            if len(sheet_names) == 1:
                output_path.write_text(csv_dict[sheet_names[0]], encoding="utf-8")
            else:
                # Write each sheet as separate file
                for name in sheet_names:
                    sheet_path = output_path.with_name(f"{output_path.stem}_{name}.csv")
                    sheet_path.write_text(csv_dict[name], encoding="utf-8")
            stats = {"sheets": len(sheet_names), "sheet_names": sheet_names}
            print(json.dumps(stats, ensure_ascii=False))

        elif input_ext == ".csv" and target == "xlsx":
            from dockit.xlsx import csv_to_xlsx

            csv_content = input_path.read_text(encoding="utf-8")
            xlsx_bytes = csv_to_xlsx(csv_content)
            output_path.write_bytes(xlsx_bytes)
            row_count = csv_content.count("\n")
            print(json.dumps({"rows": row_count}))

        elif input_ext == ".txt" and target == "xlsx":
            from dockit.xlsx import txt_to_xlsx

            txt_content = input_path.read_text(encoding="utf-8")
            xlsx_bytes = txt_to_xlsx(txt_content)
            output_path.write_bytes(xlsx_bytes)
            row_count = len([l for l in txt_content.splitlines() if l.strip()])
            print(json.dumps({"rows": row_count}))

        elif input_ext == ".txt" and target == "csv":
            from dockit.csv import txt_to_csv

            txt_content = input_path.read_text(encoding="utf-8")
            csv_content = txt_to_csv(txt_content)
            output_path.write_text(csv_content, encoding="utf-8")
            row_count = len([l for l in csv_content.splitlines() if l.strip()])
            print(json.dumps({"rows": row_count}))

        elif input_ext == ".csv" and target == "txt":
            from dockit.csv import csv_to_txt

            csv_content = input_path.read_text(encoding="utf-8")
            txt_content = csv_to_txt(csv_content)
            output_path.write_text(txt_content, encoding="utf-8")
            row_count = len([l for l in txt_content.splitlines() if l.strip()])
            print(json.dumps({"rows": row_count}))

        elif input_ext == ".xlsx" and target == "txt":
            from dockit.csv import csv_to_txt
            from dockit.xlsx import xlsx_to_csv

            xlsx_bytes = input_path.read_bytes()
            csv_dict = xlsx_to_csv(xlsx_bytes)
            sheet_names = list(csv_dict.keys())
            if len(sheet_names) == 1:
                txt_content = csv_to_txt(csv_dict[sheet_names[0]])
                output_path.write_text(txt_content, encoding="utf-8")
            else:
                for name in sheet_names:
                    sheet_path = output_path.with_name(f"{output_path.stem}_{name}.txt")
                    txt_content = csv_to_txt(csv_dict[name])
                    sheet_path.write_text(txt_content, encoding="utf-8")
            stats = {"sheets": len(sheet_names), "sheet_names": sheet_names}
            print(json.dumps(stats, ensure_ascii=False))

        elif input_ext == ".xls" and target == "xlsx":
            from dockit.xlsx import xls_to_xlsx

            xls_bytes = input_path.read_bytes()
            xlsx_bytes = xls_to_xlsx(xls_bytes)
            output_path.write_bytes(xlsx_bytes)
            print(json.dumps({"converted": True}))

        else:
            print(
                f"Error: unsupported conversion: {input_ext} -> {target}",
                file=sys.stderr,
            )
            return 1

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    return 0


def cmd_standardize_ppt(args: argparse.Namespace) -> int:
    """Handle the standardize-ppt subcommand."""
    from dockit.pptx import standardize

    input_path = Path(args.input_file).resolve()
    if not input_path.exists():
        print(f"Error: file not found: {input_path}", file=sys.stderr)
        return 1

    if input_path.suffix.lower() not in (".pptx",):
        print(f"Error: unsupported file type '{input_path.suffix}', expected .pptx", file=sys.stderr)
        return 1

    output_path = Path(args.output).resolve() if args.output else _default_output(input_path)

    try:
        pptx_bytes = input_path.read_bytes()
        result = standardize(pptx_bytes)
        output_path.write_bytes(result.data)
        print(json.dumps(result.stats, ensure_ascii=False))
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    return 0


def main():
    """DocKit CLI entry point."""
    parser = argparse.ArgumentParser(
        prog="dockit",
        description="DocKit - Document processing toolkit",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {_get_version()}",
    )
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # format-word
    p_fw = subparsers.add_parser(
        "format-word",
        help="Format a Word document (fix quotes, punctuation, units)",
    )
    p_fw.add_argument("input_file", help="Input .docx file path")
    p_fw.add_argument("-o", "--output", help="Output file path (default: <input>_formatted.docx)")
    p_fw.set_defaults(func=cmd_format_word)

    # convert
    p_cv = subparsers.add_parser(
        "convert",
        help="Convert between file formats (xlsx/csv/txt)",
    )
    p_cv.add_argument("input_file", help="Input file path")
    p_cv.add_argument("--to", required=True, choices=["csv", "xlsx", "txt"], help="Target format")
    p_cv.add_argument("-o", "--output", help="Output file path (default: same name with new extension)")
    p_cv.set_defaults(func=cmd_convert)

    # standardize-ppt
    p_sp = subparsers.add_parser(
        "standardize-ppt",
        help="Standardize a PowerPoint file (format text, unify fonts, style tables)",
    )
    p_sp.add_argument("input_file", help="Input .pptx file path")
    p_sp.add_argument("-o", "--output", help="Output file path (default: <input>_formatted.pptx)")
    p_sp.set_defaults(func=cmd_standardize_ppt)

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    exit_code = args.func(args)
    sys.exit(exit_code)


def _get_version() -> str:
    """Get the package version."""
    try:
        from dockit import __version__
        return __version__
    except ImportError:
        return "unknown"
