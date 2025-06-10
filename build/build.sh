#!/usr/bin/env bash
set -euo pipefail

# Build executable with PyInstaller
pyinstaller build/transcribe_slides.spec --distpath dist --workpath build/tmp

# Build NSIS installer
if command -v makensis >/dev/null 2>&1; then
    makensis build/installer.nsi
else
    echo "makensis not found. Skipping NSIS installer." >&2
fi
