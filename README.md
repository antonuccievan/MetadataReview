# MetadataReview

A lightweight browser interface for reviewing metadata exported to Excel.

## Features

- Upload `.xlsx` or `.xls` files directly in the browser.
- Choose workbook sheet and detect the header row from the first row where column B contains data.
- Recreate your spreadsheet as an interactive table with a generated **Hierarchy** first column.
- Expand/collapse Excel-style parent-child groupings inferred from parent/child key columns.
- Toggle light/dark theme from the top-right button, with saved preference.
- See quick depth and visibility stats while exploring.

## Run locally

Because this is a static app, you can run it with any local static server:

```bash
python3 -m http.server 8000
```

Then open: `http://localhost:8000`
