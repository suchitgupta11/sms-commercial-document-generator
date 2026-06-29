# Changelog

## v1.0.0 GA

### Added
- Quotation generation.
- Tax Invoice generation.
- Proforma Invoice generation.
- Delivery Challan generation.
- Salary Slip generation.
- Purchase Order generation using Invoice-style premium header/layout.
- DOCX and PDF generation.
- Desktop dashboard UI.
- macOS app/DMG build script.
- Windows installer build scripts.
- macOS uninstall script.

### Fixed
- macOS Tk blank dashboard issue by requiring Python 3.13 runtime during development/build.
- macOS PyInstaller `python-docx` missing XML templates issue.
- Duplicate app cleanup during uninstall.
- App close before uninstall.
- Fixed dashboard document selection state so each card generates its own document type and does not reuse the previous template.
