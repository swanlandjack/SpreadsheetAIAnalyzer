#!/bin/bash
# ============================================================
#  build-backend-mac.sh
#  Compiles server.py into a standalone Mac binary using PyInstaller.
#  Run this ONCE before `npm run build:mac`.
#
#  Usage:
#    chmod +x build-backend-mac.sh
#    ./build-backend-mac.sh
# ============================================================

set -e  # Exit immediately on any error

echo ""
echo "============================================================"
echo "  Excel AI Analyzer — Backend Compiler (Mac)"
echo "============================================================"
echo ""

# ── Step 1: Find Python ───────────────────────────────────────
PYTHON=""
for candidate in python3 python /opt/homebrew/bin/python3 /usr/local/bin/python3; do
  if command -v "$candidate" &>/dev/null; then
    VERSION=$("$candidate" --version 2>&1)
    if [[ "$VERSION" == Python\ 3* ]]; then
      PYTHON="$candidate"
      echo "✓ Python found: $PYTHON ($VERSION)"
      break
    fi
  fi
done

if [ -z "$PYTHON" ]; then
  echo "✗ Python 3 not found. Please install from https://www.python.org"
  exit 1
fi

# ── Step 2: Install PyInstaller if not present ────────────────
echo ""
echo "→ Checking PyInstaller..."
if ! $PYTHON -m PyInstaller --version &>/dev/null; then
  echo "  Installing PyInstaller..."
  $PYTHON -m pip install pyinstaller
else
  echo "  ✓ PyInstaller already installed ($($PYTHON -m PyInstaller --version))"
fi

# ── Step 3: Install backend dependencies ─────────────────────
echo ""
echo "→ Installing backend Python packages..."
$PYTHON -m pip install flask flask-cors pandas openpyxl xlrd requests numpy

# ── Step 4: Run PyInstaller ───────────────────────────────────
echo ""
echo "→ Compiling server.py → standalone binary..."
echo "  (This takes 1–3 minutes on first run)"
echo ""

cd backend
$PYTHON -m PyInstaller server.spec --clean --noconfirm
cd ..

# ── Step 5: Verify output ─────────────────────────────────────
BINARY="backend/dist/server/server"
if [ -f "$BINARY" ]; then
  SIZE=$(du -sh "$BINARY" | cut -f1)
  echo ""
  echo "============================================================"
  echo "  ✓ Build successful!"
  echo "  Binary: $BINARY"
  echo "  Size:   $SIZE"
  echo ""
  echo "  Next step: npm run build:mac"
  echo "============================================================"
else
  echo ""
  echo "✗ Build failed — binary not found at $BINARY"
  echo "  Check the PyInstaller output above for errors."
  exit 1
fi
