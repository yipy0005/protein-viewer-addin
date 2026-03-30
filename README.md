# Protein PDB Viewer — PowerPoint Add-in

Interactive 3D protein structure visualization directly in Microsoft PowerPoint. Load PDB files, explore ligand binding sites, visualize molecular interactions, render electrostatic surfaces, and present live 3D structures — all without leaving PowerPoint.

Built with [3Dmol.js](https://3dmol.csb.pitt.edu/) and the [Office.js](https://learn.microsoft.com/office/dev/add-ins/) Add-in API.

---

## Table of Contents

- [Installation (macOS)](#installation-macos)
- [Uninstallation](#uninstallation)
- [Getting Started](#getting-started)
- [TaskPane Controls](#taskpane-controls)
  - [Loading a Structure](#loading-a-structure)
  - [Protein Style](#protein-style)
  - [Ligand & Binding Site](#ligand--binding-site)
  - [Molecular Interactions](#molecular-interactions)
  - [Surfaces](#surfaces)
  - [Background & Rotation](#background--rotation)
- [Inserting into Slides](#inserting-into-slides)
  - [Interactive 3D in Slide (Content Add-in)](#interactive-3d-in-slide-content-add-in)
  - [Static Snapshot](#static-snapshot)
  - [3D Model for Slide Show (.glb)](#3d-model-for-slide-show-glb)
- [Presenter Window](#presenter-window)
  - [Opening the Presenter Window](#opening-the-presenter-window)
  - [Loading Multiple Structures](#loading-multiple-structures)
  - [Per-Entry Settings](#per-entry-settings)
  - [Structural Alignment](#structural-alignment)
  - [Keyboard Shortcuts](#keyboard-shortcuts)
- [Development Setup](#development-setup)
- [Building the Installer](#building-the-installer)
- [Architecture](#architecture)
- [Known Limitations](#known-limitations)

---

## Installation (macOS)

### Option A: DMG Installer (Recommended)

1. Download **ProteinViewerInstaller.dmg** from the [Releases page](https://github.com/yipy0005/protein-viewer-addin/releases)
2. Open the DMG file
3. Double-click **"Install Protein Viewer.command"**
   - This copies the add-in manifests to PowerPoint's sideload folder
   - You may see a macOS security prompt — right-click the file and select "Open" if needed
4. Quit PowerPoint completely (**Cmd+Q**) if it is running
5. Reopen PowerPoint
6. The **"Protein PDB Viewer"** button appears in the **Home** tab of the ribbon

### Option B: Manual Installation

1. Download `manifest-ghpages.xml` and `manifest-content-ghpages.xml` from the repository
2. Open Finder, press **Cmd+Shift+G**, and paste:
   ```
   ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
   ```
   Create the `wef` folder if it does not exist.
3. Copy both XML files into this folder, renaming them:
   - `manifest-ghpages.xml` → `ProteinViewer.xml`
   - `manifest-content-ghpages.xml` → `ProteinViewerContent.xml`
4. Quit and reopen PowerPoint

### Windows Installation

1. Download `manifest-ghpages.xml` and `manifest-content-ghpages.xml`
2. Copy them to:
   ```
   %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
   ```
   Rename as `ProteinViewer.xml` and `ProteinViewerContent.xml`.
3. Restart PowerPoint

---

## Uninstallation

### Using the Uninstaller (macOS)

1. Open the DMG file
2. Double-click **"Uninstall Protein Viewer.command"**
3. Restart PowerPoint

### Manual Uninstallation

Delete the manifest files from the sideload folder:

**macOS:**
```bash
rm ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/ProteinViewer.xml
rm ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/ProteinViewerContent.xml
```

**Windows:**
```
del %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\ProteinViewer.xml
del %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\ProteinViewerContent.xml
```

Restart PowerPoint after removing the files.

---

## Getting Started

1. Open PowerPoint and create or open a presentation
2. Click **"Protein PDB Viewer"** in the **Home** tab — the task pane opens on the right
3. Enter a 4-character PDB ID (e.g. `1CBS`, `4HHB`, `2FGI`) and click **Load**, or switch to the **File** tab to upload a local `.pdb` file
4. The 3D structure appears in the task pane viewer
5. Interact with the structure: **left-drag** to rotate, **scroll** to zoom, **right-drag** to pan

---

## TaskPane Controls

### Loading a Structure

| Method | How |
|--------|-----|
| PDB ID | Enter a 4-character ID (e.g. `1CBS`) and click **Load** |
| Local file | Switch to the **File** tab, select a `.pdb`, `.ent`, `.cif`, `.sdf`, `.mol2`, or `.xyz` file, and click **Load** |

### Protein Style

| Control | Options |
|---------|---------|
| Style | Cartoon, Stick, Sphere, Line, Cross |
| Color | Spectrum (rainbow), By Chain, Secondary Structure, By Residue, By Element |
| Opacity | 0–100% slider — lower values make the protein semi-transparent, useful when viewing ligands and binding sites |

### Ligand & Binding Site

The **Ligand** section appears automatically when the loaded structure contains ligands (non-water, non-ion HETATM records). If only one ligand is present, it is auto-selected.

| Control | Description |
|---------|-------------|
| Pick | Select a ligand from the dropdown. Ligands are listed as `RESNAME (Chain:ResID)` |
| Style | Ball & Stick (default), Stick, or Sphere for the selected ligand |
| Zoom to ligand | Centers and zooms the view on the selected ligand |
| Show binding site | Displays protein residues within a configurable distance of the ligand as sticks |
| Distance | Slider (2–10 Å) controlling the binding site radius |
| Label residues | Adds residue name labels (e.g. ASP 99, PHE 120) to binding site residues |

### Molecular Interactions

These checkboxes appear within the Ligand section when a ligand is selected and binding site is enabled. Each interaction type is shown as colored dashed lines:

| Interaction | Color | Detection Criteria |
|-------------|-------|--------------------|
| H-bonds | Yellow | N, O, or S atoms within 2.0–3.5 Å between ligand and protein |
| Salt bridges | Magenta | Charged groups (ARG/LYS NH ↔ ASP/GLU carboxylate, or ligand N/O) within 4.0 Å |
| π–π stacking | Cyan | Aromatic ring centroids within 5.5 Å |
| π–cation | Orange | Aromatic ring centroid to cation (LYS Nζ, ARG NH) within 6.0 Å |

Protein aromatic rings are detected for PHE, TYR, TRP (both 5- and 6-membered rings), and HIS. Ligand aromatic rings are detected automatically from C/N atom connectivity (5- and 6-membered rings).

### Surfaces

#### Full Protein Surface

| Control | Options |
|---------|---------|
| Show Surface | Toggle the molecular surface on/off |
| Type | Van der Waals (VDW), Solvent Accessible (SAS), Solvent Excluded (SES), **ESP (Electrostatic)** |
| Color | White, Same as Protein, ESP, Hydrophobicity, Element (only used when Type is not ESP) |
| Opacity | 0–100% slider |

**ESP Surface**: Selecting "ESP (Electrostatic)" as the surface type renders a Solvent Accessible Surface colored by residue-level charge: **red** (negative: ASP, GLU), **white** (neutral), **blue** (positive: ARG, LYS, HIS). This is a residue-level approximation, not a full Poisson-Boltzmann calculation.

#### Binding Site Surface

Available when a ligand is selected with binding site enabled:

| Control | Options |
|---------|---------|
| Binding site surface | Toggle on/off |
| Color | ESP (default), White, Element, Hydrophobicity |
| Opacity | 0–100% slider |

The binding site surface is computed only from protein atoms — the ligand does not influence the surface shape, so you see the true pocket geometry.

### Background & Rotation

| Control | Options |
|---------|---------|
| Auto-rotate | Toggle continuous Y-axis rotation |
| Background | White, Black, Dark Blue, Light Gray |

---

## Inserting into Slides

### Interactive 3D in Slide (Content Add-in)

1. Go to **Insert → Add-ins** (or **Insert → My Add-ins**)
2. Find **"Protein Viewer Slide"** and insert it onto your slide
3. In the task pane, configure your view (load structure, set style, select ligand, etc.)
4. Click **"Push to Slide Viewer"**
5. The in-slide viewer updates within ~1 second with the same view
6. You can rotate, zoom, and pan the molecule directly on the slide in edit mode
7. Resize and reposition the viewer frame as needed

> **Note:** Content add-ins do not render during PowerPoint Slide Show mode. Use the static snapshot or .glb export for presentations.

### Static Snapshot

Click **"Insert Snapshot"** to capture the current 3D view as a PNG image and insert it directly into the active slide. This works in all modes including Slide Show.

### 3D Model for Slide Show (.glb)

1. Click **"Download .glb"** to export the molecular structure as a GLB 3D model file
2. In PowerPoint, go to **Insert → 3D Models → This Device** and select the downloaded `.glb` file
3. The 3D model is embedded in the slide and **interactive during Slide Show mode** — click and drag to rotate, scroll to zoom
4. The `.glb` file is embedded in the `.pptx`, so recipients can interact with it without any add-in installed

> **Note:** The GLB export includes atoms (as spheres) and bonds (as cylinders) with element-based coloring. Surfaces, labels, and interaction lines are not included in the GLB format.

---

## Presenter Window

The presenter window is a standalone browser-based 3D viewer with full controls. It works independently of PowerPoint and requires no installation.

**URL:** [https://yipy0005.github.io/protein-viewer-addin/presenter.html](https://yipy0005.github.io/protein-viewer-addin/presenter.html)

### Opening the Presenter Window

- Click **"Open Presenter Window"** in the task pane, or
- Open the URL directly in any browser

### Loading Multiple Structures

The presenter supports loading multiple PDB entries simultaneously:

- **By PDB ID:** Enter an ID and click **Add** — each entry is added to the list (not replaced)
- **By file:** Use the file picker or **drag and drop** one or more PDB files onto the viewer
- **From PowerPoint:** Clicking "Push to Slide Viewer" in the task pane also sends the structure to the presenter window (appears as a "PowerPoint" entry)

### Per-Entry Settings

Each loaded entry appears in the **Loaded Entries** list with:

| Control | Description |
|---------|-------------|
| 👁 (eye icon) | Toggle visibility — hide/show individual structures |
| Entry name | Click to select and edit that entry's settings |
| ✕ | Remove the entry |

When an entry is selected, the settings panel below shows all the same controls as the task pane (style, color, opacity, ligand, binding site, interactions, surfaces) — each configured independently per entry.

### Structural Alignment

When 2 or more entries are loaded:

1. The **alignment section** appears below the entry list
2. Select a **reference structure** from the dropdown
3. Click **"Align All"**
4. All other structures are superimposed onto the reference using the **Kabsch algorithm** (RMSD-minimizing rotation and translation on Cα atoms)
5. Alignment matches Cα atoms by sequential index — works well for comparing the same protein in different conformations, mutant vs wild-type, apo vs holo, etc.

### Keyboard Shortcuts

| Key | Action |
|-----|--------|
| **S** | Toggle auto-spin |
| **P** | Toggle side panel |

---

## Development Setup

For local development with hot-reloading:

```bash
cd protein-viewer-addin

# Install Node.js dependencies
npm install

# Install Python dependencies (for local HTTPS server)
pixi install

# Build the frontend
npm run build

# Sideload manifests (local dev versions pointing to localhost:3001)
pixi run sideload

# Start the local HTTPS server on port 3001
pixi run serve
```

### Pixi Tasks

| Task | Command | Description |
|------|---------|-------------|
| `pixi run serve` | `python server/server_main.py` | Start HTTPS server on port 3001 |
| `pixi run build` | `npx webpack --mode production` | Build webpack frontend to `dist/` |
| `pixi run sideload` | *(copies manifests)* | Register local dev manifests with PowerPoint |

---

## Building the Installer

```bash
bash installer/build-dmg.sh
```

Output: `release/ProteinViewerInstaller.dmg` (~20 KB)

The DMG contains:
- `ProteinViewer.xml` — TaskPane add-in manifest (points to GitHub Pages)
- `ProteinViewerContent.xml` — Content add-in manifest (points to GitHub Pages)
- `Install Protein Viewer.command` — macOS install script
- `Uninstall Protein Viewer.command` — macOS uninstall script
- `README.txt` — Quick start guide

---

## Architecture

```
protein-viewer-addin/
├── src/
│   ├── taskpane/           # PowerPoint side panel (Office.js + 3Dmol.js)
│   │   ├── taskpane.html
│   │   ├── taskpane.js
│   │   └── taskpane.css
│   ├── content/            # In-slide 3D viewer (Content Add-in)
│   │   ├── content.html
│   │   ├── content.js
│   │   └── content.css
│   ├── presenter/          # Standalone browser viewer
│   │   ├── presenter.html
│   │   ├── presenter.js
│   │   └── presenter.css
│   ├── commands/           # Ribbon button handlers
│   │   ├── commands.html
│   │   └── commands.js
│   └── viewer/
│       └── glbexport.js    # Three.js GLB exporter
├── server/                 # Python FastAPI dev server (local HTTPS)
│   ├── server.py
│   └── server_main.py
├── assets/                 # Add-in icons (16, 32, 80px)
├── installer/              # DMG build scripts
│   ├── build-dmg.sh
│   ├── install.command
│   └── uninstall.command
├── manifest.xml            # TaskPane manifest (localhost, for dev)
├── manifest-content.xml    # Content manifest (localhost, for dev)
├── manifest-ghpages.xml    # TaskPane manifest (GitHub Pages, for distribution)
├── manifest-content-ghpages.xml  # Content manifest (GitHub Pages)
├── webpack.config.js
├── package.json
├── pixi.toml
└── .github/workflows/deploy.yml  # GitHub Pages auto-deployment
```

### How It Works

- The **TaskPane add-in** runs in PowerPoint's side panel, providing all controls for loading structures and configuring visualization
- The **Content add-in** embeds a 3Dmol.js viewer directly inside the slide canvas. It polls `localStorage` every 500ms for updates pushed from the TaskPane
- The **Presenter window** is a standalone HTML page with its own 3Dmol.js viewer and full controls, designed for live demos during presentations
- Communication between TaskPane and Content/Presenter uses `localStorage` (same-origin)
- The **GLB exporter** uses Three.js to convert 3Dmol.js molecular geometry into a binary glTF file that PowerPoint can natively display as an interactive 3D model

### Technology Stack

| Component | Technology |
|-----------|-----------|
| 3D Rendering | [3Dmol.js](https://3dmol.csb.pitt.edu/) (WebGL) |
| Office Integration | [Office.js](https://learn.microsoft.com/office/dev/add-ins/) |
| GLB Export | [Three.js](https://threejs.org/) r128 + GLTFExporter |
| Frontend Build | Webpack 5 + Babel |
| Local Dev Server | Python FastAPI + uvicorn (via pixi) |
| Hosting | GitHub Pages |
| CI/CD | GitHub Actions |

---

## Known Limitations

- **Slide Show mode:** Content add-ins (in-slide 3Dmol.js viewer) do not render during PowerPoint Slide Show mode. Use static snapshots or .glb 3D models for presentations.
- **ESP surface:** Uses residue-level charge assignments (ARG/LYS positive, ASP/GLU negative), not a full Poisson-Boltzmann electrostatic calculation.
- **GLB export:** Includes atoms and bonds only. Surfaces, labels, interaction lines, and cartoon representations are not exported.
- **Structural alignment:** Matches Cα atoms by sequential index. For distantly related proteins with different lengths, a sequence alignment step would improve matching.
- **localStorage sync:** The TaskPane → Content/Presenter sync via localStorage only works when both are served from the same origin. The GitHub Pages version syncs between TaskPane and Content add-in (both on `github.io`), but the Presenter window also needs to be on the same origin.
- **Large structures:** Very large PDB files (>10 MB) may be slow to render or transfer via localStorage.

---

## License

MIT
