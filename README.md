# MetadataReview

A lightweight browser interface for reviewing metadata exported to Excel.

## Features

- Upload `.xlsx` or `.xls` files directly in the browser.
- Choose workbook sheet and automatically treat row 1 as header columns.
- Recreate your spreadsheet as an interactive table.
- Expand/collapse Excel-style parent-child groupings inferred from indentation in the first column.
- See quick depth and visibility stats while exploring.

## Run locally

Because this is a static app, you can run it with any local static server:

```bash
python3 -m http.server 8000
```

Then open: `http://localhost:8000`

## Hierarchy parsing

The hierarchy level for each row is inferred from indentation in the first column using:

1. Leading tab characters (`\t`).
2. Leading spaces (2 spaces ≈ 1 level).
3. Repeated bullets/punctuation prefixes (`.`, `-`, `*`, `•`) as a fallback.

Row 1 is always interpreted as headers, and the remaining rows become groupable data rows.
