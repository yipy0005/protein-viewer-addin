#!/bin/bash
# Build the lightweight DMG installer for Protein PDB Viewer
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
STAGING="$PROJECT_DIR/installer/staging"
DMG_NAME="ProteinViewerInstaller"
OUTPUT="$PROJECT_DIR/release"

echo "🧬 Building Protein PDB Viewer DMG..."

# Clean
rm -rf "$STAGING" "$OUTPUT/$DMG_NAME.dmg"
mkdir -p "$STAGING" "$OUTPUT"

# Copy manifests (GitHub Pages versions)
cp "$PROJECT_DIR/manifest-ghpages.xml" "$STAGING/ProteinViewer.xml"
cp "$PROJECT_DIR/manifest-content-ghpages.xml" "$STAGING/ProteinViewerContent.xml"

# Copy install/uninstall scripts
cp "$SCRIPT_DIR/install.command" "$STAGING/Install Protein Viewer.command"
cp "$SCRIPT_DIR/uninstall.command" "$STAGING/Uninstall Protein Viewer.command"
chmod +x "$STAGING/Install Protein Viewer.command"
chmod +x "$STAGING/Uninstall Protein Viewer.command"

# Create a README
cat > "$STAGING/README.txt" << 'EOF'
Protein PDB Viewer for PowerPoint
==================================

To install:
  Double-click "Install Protein Viewer.command"

To uninstall:
  Double-click "Uninstall Protein Viewer.command"

After installing, restart PowerPoint. The add-in appears in the Home tab.

Presenter window (for live demos):
  https://yipy0005.github.io/protein-viewer-addin/presenter.html
EOF

# Build DMG
hdiutil create -volname "$DMG_NAME" \
  -srcfolder "$STAGING" \
  -ov -format UDZO \
  "$OUTPUT/$DMG_NAME.dmg"

# Clean staging
rm -rf "$STAGING"

echo ""
echo "✅ DMG created: $OUTPUT/$DMG_NAME.dmg"
echo "   Size: $(du -h "$OUTPUT/$DMG_NAME.dmg" | cut -f1)"
