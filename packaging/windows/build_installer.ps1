$ErrorActionPreference = "Stop"
$Root = Resolve-Path "$PSScriptRoot\..\.."
$Inno = "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe"
if (!(Test-Path $Inno)) { $Inno = "${env:ProgramFiles}\Inno Setup 6\ISCC.exe" }
if (!(Test-Path $Inno)) { throw "Inno Setup 6 not found. Install it once on the build machine." }
New-Item -ItemType Directory -Force -Path "$Root\release\windows" | Out-Null
& $Inno "$Root\packaging\windows\installer.iss"
Write-Host "Installer created under release\windows"
