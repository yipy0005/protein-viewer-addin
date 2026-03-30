#!/bin/bash
# Protein PDB Viewer — macOS Installer
# Double-click this file to install the PowerPoint add-in.

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo ""
echo "🧬 Protein PDB Viewer — Installer"
echo "==================================="
echo ""

# Sideload manifests
WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
mkdir -p "$WEF_DIR"

cp "$SCRIPT_DIR/ProteinViewer.xml" "$WEF_DIR/ProteinViewer.xml"
cp "$SCRIPT_DIR/ProteinViewerContent.xml" "$WEF_DIR/ProteinViewerContent.xml"

echo "✅ Add-in manifests installed."
echo ""
echo "Next steps:"
echo "  1. Quit PowerPoint if it's open (Cmd+Q)"
echo "  2. Reopen PowerPoint"
echo "  3. Look for 'Protein PDB Viewer' in the Home tab"
echo "  4. The presenter window is at:"
echo "     https://yipy0005.github.io/protein-viewer-addin/presenter.html"
echo ""

# Try to open PowerPoint
if [ -d "/Applications/Microsoft PowerPoint.app" ]; then
  read -p "Open PowerPoint now? (y/n) " -n 1 -r
  echo ""
  if [[ $REPLY =~ ^[Yy]$ ]]; then
    open -a "Microsoft PowerPoint"
  fi
fi

echo "Done! You can close this window."
