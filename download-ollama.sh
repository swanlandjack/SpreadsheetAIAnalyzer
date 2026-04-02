#!/bin/bash
# ============================================================
#  download-ollama.sh
#  Downloads and extracts the Ollama binary for Mac and places
#  it in the ollama/ folder so electron-builder can bundle it.
#
#  Run this ONCE before building the final installer:
#    chmod +x download-ollama.sh
#    ./download-ollama.sh
# ============================================================

set -e

OLLAMA_DIR="./ollama"
BINARY_NAME="ollama"

echo ""
echo "============================================================"
echo "  Excel AI Analyzer — Ollama Binary Downloader"
echo "============================================================"
echo ""

# ── Get latest version from GitHub API ───────────────────────
echo "→ Checking latest Ollama version..."
LATEST=$(curl -s https://api.github.com/repos/ollama/ollama/releases/latest | grep '"tag_name"' | head -1 | sed 's/.*"v\([^"]*\)".*/\1/')
echo "  Latest version: v$LATEST"

DOWNLOAD_URL="https://github.com/ollama/ollama/releases/download/v${LATEST}/ollama-darwin.tgz"

mkdir -p "$OLLAMA_DIR"

# ── Download the tgz ─────────────────────────────────────────
echo ""
echo "→ Downloading from:"
echo "  $DOWNLOAD_URL"
echo ""

TGZ_PATH="$OLLAMA_DIR/ollama-darwin.tgz"
curl -L --progress-bar "$DOWNLOAD_URL" -o "$TGZ_PATH"

# ── Extract the binary ────────────────────────────────────────
echo ""
echo "→ Extracting binary..."
tar -xzf "$TGZ_PATH" -C "$OLLAMA_DIR"

# The tgz contains the binary named 'ollama' at the root
# Find and move it to the right place
EXTRACTED=$(find "$OLLAMA_DIR" -name "ollama" -type f | head -1)

if [ -z "$EXTRACTED" ]; then
  echo "✗ Could not find ollama binary after extraction."
  echo "  Contents of $OLLAMA_DIR:"
  ls -la "$OLLAMA_DIR"
  exit 1
fi

# Move to final location if not already there
if [ "$EXTRACTED" != "$OLLAMA_DIR/$BINARY_NAME" ]; then
  mv "$EXTRACTED" "$OLLAMA_DIR/$BINARY_NAME"
fi

# Clean up tgz and any subdirectories
rm -f "$TGZ_PATH"
find "$OLLAMA_DIR" -mindepth 1 -not -name "ollama" -delete 2>/dev/null || true

chmod +x "$OLLAMA_DIR/$BINARY_NAME"

# ── Verify ───────────────────────────────────────────────────
DEST="$OLLAMA_DIR/$BINARY_NAME"
if [ -f "$DEST" ]; then
  SIZE=$(du -sh "$DEST" | cut -f1)
  echo ""
  echo "============================================================"
  echo "  ✓ Ollama binary ready!"
  echo "  Location : $DEST"
  echo "  Size     : $SIZE"
  echo "  Version  : v$LATEST"
  echo ""
  echo "  Next step: npx electron-builder --mac"
  echo "============================================================"
else
  echo "✗ Binary not found at expected location: $DEST"
  exit 1
fi
