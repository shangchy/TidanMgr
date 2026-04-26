#!/usr/bin/env bash
# Run on macOS: produces dist/TidanMgr.app, helper scripts, and a portable zip (default
# TidanMgr-macos-portable.zip; CI may set TIDANMGR_PORTABLE_ZIP for arm64/intel-specific names).
# Usage: cd app && bash build_macos.sh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 not found. Install Python 3.10+."
  exit 1
fi

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

# shellcheck disable=SC1091
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt

APP_NAME="TidanMgr"
ADD_DATA_ARGS=()
if [ -f "template.xlsx" ]; then
  ADD_DATA_ARGS+=(--add-data "template.xlsx:.")
  echo "Bundled: template.xlsx"
else
  echo "Warning: template.xlsx missing; place it in app/ before packaging for export."
fi
if [ -f "sum-template.xlsx" ]; then
  ADD_DATA_ARGS+=(--add-data "sum-template.xlsx:.")
  echo "Bundled: sum-template.xlsx"
else
  echo "Warning: sum-template.xlsx missing; merge stats sheet may lose template style."
fi

echo "Running PyInstaller..."
pyinstaller \
  --noconfirm \
  --windowed \
  --clean \
  --name "$APP_NAME" \
  --hidden-import bill_theme \
  "${ADD_DATA_ARGS[@]}" \
  bill_app.py

# Portable mode: data in dist/TidanMgrData (do not use "open" for GUI)
PORTABLE_LAUNCH="$SCRIPT_DIR/dist/PortableStart.command"
cat >"$PORTABLE_LAUNCH" <<EOS
#!/bin/bash
cd "\$(dirname "\$0")" || exit 1
export TIDANMGR_PORTABLE=1
exec "\$(cd "\$(dirname "\$0")" && pwd)/${APP_NAME}.app/Contents/MacOS/${APP_NAME}"
EOS
chmod +x "$PORTABLE_LAUNCH"

RUN_SH="$SCRIPT_DIR/dist/run_TidanMgr.sh"
cat >"$RUN_SH" <<EOS
#!/bin/bash
cd "\$(dirname "\$0")" || exit 1
export TIDANMGR_PORTABLE=1
exec "\$(cd "\$(dirname "\$0")" && pwd)/${APP_NAME}.app/Contents/MacOS/${APP_NAME}" "\$@"
EOS
chmod +x "$RUN_SH"

# Default zip name keeps historical CI/docs; override with TIDANMGR_PORTABLE_ZIP for arch-specific builds.
ZIP_BASENAME="${TIDANMGR_PORTABLE_ZIP:-TidanMgr-macos-portable.zip}"
ZIP_PATH="$SCRIPT_DIR/dist/$ZIP_BASENAME"
ZIP_TMP="$SCRIPT_DIR/${ZIP_BASENAME}.tmp"
rm -f "$ZIP_PATH" "$ZIP_TMP"
(
  cd "$SCRIPT_DIR/dist"
  zip -ry "$ZIP_TMP" "${APP_NAME}.app" PortableStart.command run_TidanMgr.sh
)
mv -f "$ZIP_TMP" "$ZIP_PATH"

MAIN_BIN="$SCRIPT_DIR/dist/${APP_NAME}.app/Contents/MacOS/${APP_NAME}"
if [ -f "$MAIN_BIN" ]; then
  echo ""
  echo "Main executable (file):"
  file "$MAIN_BIN" || true
fi

echo ""
echo "Done."
echo "  Standard: drag ${APP_NAME}.app to Applications; data in ~/Library/Application Support/TidanMgr/"
echo "  Portable: double-click dist/PortableStart.command (data in dist/TidanMgrData/)"
echo "  Debug:    bash dist/run_TidanMgr.sh"
echo "  Zip:      $ZIP_PATH"
