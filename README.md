# Protein PDB Viewer — PowerPoint Add-in

Interactive 3D protein structure visualization directly in PowerPoint. Load PDB files, explore binding sites, view molecular interactions, and present live 3D structures.

## Install (macOS)

1. Download `ProteinViewerInstaller.dmg` from [Releases](https://github.com/yipy0005/protein-viewer-addin/releases)
2. Open the DMG and double-click **"Install Protein Viewer.command"**
3. Restart PowerPoint — the add-in appears in the **Home** tab

## Features

- Load structures by PDB ID or local file
- Protein styles: Cartoon, Stick, Sphere, Line
- Color schemes: Spectrum, Chain, Secondary Structure, Residue, Element
- Ligand picker with binding site visualization
- Interactions: H-bonds, salt bridges, π–π stacking, π–cation
- Surfaces: VDW, SAS, SES, ESP (electrostatic)
- Adjustable opacity for protein and surfaces
- Insert static snapshots into slides
- Download .glb for native PowerPoint 3D Models (interactive in Slide Show)
- Standalone presenter window with full controls for live demos

## Presenter Window

Open in any browser for live 3D during presentations — no PowerPoint add-in needed:

**https://yipy0005.github.io/protein-viewer-addin/presenter.html**

Supports multiple structures, structural alignment, show/hide entries, and all visualization options.

## Development

```bash
cd protein-viewer-addin
npm install
pixi install
npm run build
pixi run sideload
pixi run serve
```

## Build DMG

```bash
bash installer/build-dmg.sh
```

Output: `release/ProteinViewerInstaller.dmg`
