#!/bin/bash
# ============================================================
#  generate-icns.sh
#  Converts the icon.iconset folder into icon.icns
#  Must be run on a Mac (requires iconutil, built into macOS)
#
#  Usage:
#    chmod +x generate-icns.sh
#    ./generate-icns.sh
# ============================================================

ICONSET="assets/icon.iconset"
OUTPUT="assets/icon.icns"

if [ ! -d "$ICONSET" ]; then
  echo "✗ Iconset folder not found: $ICONSET"
  echo "  Make sure you are running this from the excel-ai-electron directory."
  exit 1
fi

echo "→ Converting iconset to .icns..."
iconutil -c icns "$ICONSET" -o "$OUTPUT"

if [ -f "$OUTPUT" ]; then
  SIZE=$(du -sh "$OUTPUT" | cut -f1)
  echo "✓ Icon created: $OUTPUT ($SIZE)"
  echo ""
  echo "You can now run: npx electron-builder --mac"
else
  echo "✗ Failed to create icon."
  exit 1
fi
