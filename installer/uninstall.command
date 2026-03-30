#!/bin/bash
# Protein PDB Viewer — Uninstaller

WEF_DIR="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"

echo ""
echo "🧬 Protein PDB Viewer — Uninstaller"
echo "====================================="
echo ""

rm -f "$WEF_DIR/ProteinViewer.xml"
rm -f "$WEF_DIR/ProteinViewerContent.xml"

echo "✅ Add-in removed. Restart PowerPoint."
