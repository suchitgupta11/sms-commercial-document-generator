#!/usr/bin/env bash
set -euo pipefail

APP_NAME="SMS Commercial Document Generator"
SPEC_FILE="SMS_Commercial_Document_Generator.spec"
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
cd "$ROOT"

echo "==============================================="
echo " Building ${APP_NAME} v1.0.0 for macOS"
echo "==============================================="

if ! command -v python3 >/dev/null 2>&1; then
  echo "ERROR: python3 not found. Install Python 3.13 from python.org."
  exit 1
fi

PY_VERSION="$(python3 - <<'PY'
import sys
print(f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
PY
)"
echo "Using Python: ${PY_VERSION}"

rm -rf build dist release/macos
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install --upgrade pyinstaller python-docx

# Create .icns from PNG when iconutil is available.
if command -v iconutil >/dev/null 2>&1 && [ -f assets/app_icon.png ] && [ ! -f assets/app_icon.icns ]; then
  rm -rf assets/app_icon.iconset
  mkdir -p assets/app_icon.iconset
  for s in 16 32 128 256 512; do
    sips -z "$s" "$s" assets/app_icon.png --out "assets/app_icon.iconset/icon_${s}x${s}.png" >/dev/null
    double=$((s*2))
    sips -z "$double" "$double" assets/app_icon.png --out "assets/app_icon.iconset/icon_${s}x${s}@2x.png" >/dev/null
  done
  iconutil -c icns assets/app_icon.iconset -o assets/app_icon.icns
  rm -rf assets/app_icon.iconset
fi

pyinstaller --clean --noconfirm "$SPEC_FILE"

APP="dist/${APP_NAME}.app"
if [ ! -d "$APP" ]; then
  echo "ERROR: Built app not found: $APP"
  exit 1
fi

# Critical macOS PyInstaller/python-docx fix:
# python-docx resolves XML templates relative to its installed package directory under
# Contents/Frameworks/docx. PyInstaller can place them under Contents/Resources/docx,
# so we explicitly copy templates and parts beside the bundled docx package.
DOCX_SRC="$(python - <<'PY'
import docx
from pathlib import Path
print(Path(docx.__file__).parent)
PY
)"
echo "Bundling python-docx package data from: $DOCX_SRC"
mkdir -p "$APP/Contents/Frameworks/docx"
ditto "$DOCX_SRC/templates" "$APP/Contents/Frameworks/docx/templates"
ditto "$DOCX_SRC/parts" "$APP/Contents/Frameworks/docx/parts"

if [ ! -f "$APP/Contents/Frameworks/docx/templates/default-footer.xml" ]; then
  echo "ERROR: python-docx template fix failed: default-footer.xml missing"
  exit 1
fi

echo "python-docx template fix verified."

mkdir -p release/macos
if command -v create-dmg >/dev/null 2>&1; then
  rm -f "release/macos/${APP_NAME}.dmg"
  create-dmg \
    --volname "$APP_NAME" \
    --window-pos 200 120 \
    --window-size 640 420 \
    --icon-size 128 \
    --app-drop-link 480 210 \
    "release/macos/${APP_NAME}.dmg" \
    "$APP"
else
  hdiutil create -volname "$APP_NAME" -srcfolder "$APP" -ov -format UDZO "release/macos/${APP_NAME}.dmg"
fi

echo "Build complete: release/macos/${APP_NAME}.dmg"
echo "Test app directly with: open \"$APP\""
