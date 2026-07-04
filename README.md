# SMS Commercial Document Generator

**Version:** 1.0.0 GA  
**Company:** SMS Controls & Automation

A desktop application for generating professional commercial documents from Excel templates. Purchase Order support was added using the same premium Invoice-style header and layout approach.

## Supported Documents

- Quotation
- Tax Invoice
- Proforma Invoice
- Delivery Challan
- Salary Slip
- Purchase Order

Each document supports DOCX and PDF generation using the finalized business logic and layouts.

## Repository Name

Recommended Git repository name:

```text
sms-commercial-document-generator
```

## Development Run

Use Python 3.13 from python.org on macOS to avoid the deprecated Apple Tk runtime.

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python commercial_document_generator_app.py
```

## Smoke Test

```bash
source .venv/bin/activate
python tests/smoke_generation.py
```

## Build macOS App + DMG

```bash
chmod +x packaging/macos/build_macos.sh
./packaging/macos/build_macos.sh
```

Output:

```text
release/macos/SMS Commercial Document Generator.dmg
```

The macOS build script includes the confirmed `python-docx` fix by copying `docx/templates` and `docx/parts` into:

```text
dist/SMS Commercial Document Generator.app/Contents/Frameworks/docx/
```

## Uninstall macOS App

```bash
chmod +x Uninstall.command
./Uninstall.command
```

The script closes the running app, removes duplicate app copies, removes user data/cache/logs, and refreshes Launchpad.

## Build Windows Installer

Use Windows PowerShell:

```powershell
.\packaging\windows\build_windows.ps1
.\packaging\windows\build_installer.ps1
```

Output:

```text
release\windows\SMS Commercial Document Generator Setup.exe
```

## Git First Commit

```bash
git init
git add .
git commit -m "Initial commit: SMS Commercial Document Generator v1.0.0"
git branch -M main
git remote add origin <repo-url>
git push -u origin main
git tag -a v1.0.0 -m "General Availability release"
git push origin v1.0.0
```
