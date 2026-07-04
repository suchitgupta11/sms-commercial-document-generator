# Optional bundled LibreOffice runtime

For production builds that require DOCX-to-PDF conversion without asking the customer to install LibreOffice, place the portable LibreOffice runtime here before running PyInstaller:

- Windows: `vendor/libreoffice/program/soffice.exe`
- macOS: `vendor/libreoffice/LibreOffice.app/Contents/MacOS/soffice`

The application automatically checks this bundled path first, then falls back to system LibreOffice/Pages when running from source.

Invoice, Proforma Invoice, Delivery Challan, and Salary Slip already include native ReportLab PDF generation. Quotation PDF currently requires DOCX-to-PDF conversion, so bundled LibreOffice is recommended for customer installers.
