#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
chmod +x packaging/macos/build_macos.sh
./packaging/macos/build_macos.sh
