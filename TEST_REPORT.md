# Test Report — v1.0.0 GA Clean Repository

## Smoke generation

Executed in the cleaned repository:

```bash
python3 tests/smoke_generation.py
```

Result:

```text
SMOKE_GENERATION_OK= True
```

Generated successfully during smoke test:

- Quotation: DOCX + PDF
- Tax Invoice: DOCX + PDF
- Proforma Invoice: DOCX + PDF
- Delivery Challan: DOCX + PDF
- Salary Slip: DOCX + PDF
- Purchase Order: DOCX + PDF

## Packaging fixes included

- macOS `python-docx` fix: copies `docx/templates` and `docx/parts` into `Contents/Frameworks/docx/` after PyInstaller build.
- macOS uninstaller closes running app before removal.
- Uninstaller removes duplicate copies from `/Applications`, `~/Applications`, and Desktop.
- Fixed dashboard document selection state so each card generates its own document type and does not reuse the previous template.
