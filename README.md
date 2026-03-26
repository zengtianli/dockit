# DocKit

Document processing toolkit for Word, PowerPoint, Excel, and CSV files.

**Bytes in, bytes out.** Pure processing logic with no file I/O assumptions — use it from CLI scripts, web apps, or any Python program.

## Install

```bash
pip install dockit
```

## Quick Start

```python
from dockit.docx import format_text

# Read file
with open("input.docx", "rb") as f:
    doc_bytes = f.read()

# Process
result = format_text(doc_bytes, fix_quotes=True, fix_punctuation=True, fix_units=True)

# Write result
with open("output.docx", "wb") as f:
    f.write(result.data)

print(result.stats)  # {"quotes": 5, "punctuation": 12, "units": 3}
```

## Features

### Text Formatting (`dockit.text`)
- Fix quote pairing (smart Chinese quotes)
- Convert English punctuation to Chinese equivalents
- Convert Chinese unit names to standard symbols (e.g. 平方米 → m²)

### Word Processing (`dockit.docx`)
- Format text in Word documents (quotes, punctuation, units)
- Quote font splitting (set specific font for quote characters)
- Process paragraphs, tables, headers, and footers

### PowerPoint Processing (`dockit.pptx`)
- Unify fonts across all slides and masters
- Fix text formatting (quotes, punctuation, units)
- Set table style options (header row, banded rows, first column)
- One-click standardization (all of the above)

### Excel Processing (`dockit.xlsx`)
- Convert between XLSX, CSV, and TXT formats
- Split workbook into per-sheet files
- Lowercase column headers
- Convert legacy .xls to .xlsx

### CSV Processing (`dockit.csv`)
- Auto-detect delimiters
- Convert between CSV and delimited text
- Replace circled numbers with plain format
- Reorder rows by a reference list

## Web App

DocKit includes a Streamlit web interface for non-technical users:

```bash
pip install dockit[web]
streamlit run app/app.py
```

## License

MIT
