$ErrorActionPreference = "Stop"
$Root = Resolve-Path "$PSScriptRoot\..\.."
Set-Location $Root

Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip install --upgrade pyinstaller python-docx
.\.venv\Scripts\pyinstaller.exe --clean --noconfirm SMS_Commercial_Document_Generator.spec

Write-Host "Build complete: dist\SMS Commercial Document Generator"
Write-Host "Next: run packaging\windows\build_installer.ps1 to create Setup.exe"
