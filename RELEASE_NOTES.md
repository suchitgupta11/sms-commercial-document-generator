# Release Notes — v1.0.0 GA

This is the General Availability release of **SMS Commercial Document Generator**.

## Highlights

- Generates six commercial document types, including Purchase Order.
- Uses finalized document layouts and business logic.
- Supports DOCX and PDF output.
- Purchase Order uses the same premium header/layout standard as Tax Invoice.
- Includes macOS and Windows build automation.
- Includes macOS uninstall support.

## macOS Packaging Note

The build includes the confirmed `python-docx` package-data fix. The app bundle is patched after PyInstaller build so that `default-footer.xml` and related templates are available under `Contents/Frameworks/docx/templates`.

## Recommended Runtime

- Python 3.13 for development and build.
- End users do not need Python when using the packaged app/installer.
- Fixed dashboard document selection state so each card generates its own document type and does not reuse the previous template.
