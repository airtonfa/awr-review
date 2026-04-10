#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT_DIR"

APP_NAME="AWR Review"
PKG_NAME="AWR-Review-macOS.pkg"
VENV_DIR="${ROOT_DIR}/.venv-pack"
PYTHON_BIN="${PYTHON_BIN:-python3}"

echo "[1/6] Creating packaging virtual environment..."
"$PYTHON_BIN" -m venv "$VENV_DIR"
source "${VENV_DIR}/bin/activate"

echo "[2/6] Installing packaging dependencies..."
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo "[3/6] Cleaning previous build artifacts..."
rm -rf build dist

echo "[4/6] Building macOS .app with PyInstaller..."
pyinstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "$APP_NAME" \
  --add-data "index.html:." \
  --add-data "styles.css:." \
  --add-data "app.js:." \
  --add-data "vendor:vendor" \
  desktop_launcher.py

APP_PATH="dist/${APP_NAME}.app"
PKG_PATH="dist/${PKG_NAME}"

if [[ ! -d "$APP_PATH" ]]; then
  echo "ERROR: App bundle not found at ${APP_PATH}" >&2
  exit 1
fi

if ! command -v pkgbuild >/dev/null 2>&1; then
  echo "ERROR: pkgbuild is not available. Install Xcode Command Line Tools." >&2
  exit 1
fi

echo "[5/6] Building installer package (.pkg)..."
pkgbuild \
  --identifier "com.airtonfa.awr-review" \
  --version "1.0.0" \
  --install-location "/Applications" \
  --component "$APP_PATH" \
  "$PKG_PATH"

echo "[6/6] Done."
echo "App bundle: ${APP_PATH}"
echo "Installer : ${PKG_PATH}"
