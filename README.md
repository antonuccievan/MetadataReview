# MetadataReview

A lightweight browser interface for reviewing metadata exported to Excel.

## Features

- Upload `.xlsx` or `.xls` files directly in the browser.
- Choose workbook sheet and hierarchy column.
- Recreate your spreadsheet as an interactive table.
- Expand/collapse parent-child rows inferred from indentation in the hierarchy column.
- See quick depth and visibility stats while exploring.

## Run locally

Because this is a static app, you can run it with any local static server:

```bash
python3 -m http.server 8000
```

Then open: `http://localhost:8000`

## Hierarchy parsing

The hierarchy level for each row is inferred from the selected hierarchy column using:

1. Leading tab characters (`\t`).
2. Leading spaces (2 spaces ≈ 1 level).
3. Repeated bullets/punctuation prefixes (`.`, `-`, `*`, `•`) as a fallback.

You can switch hierarchy columns at any time to re-parse the tree.
