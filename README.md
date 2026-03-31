# Protein PDB Viewer — PowerPoint Add-in

Interactive 3D protein structure visualization directly in Microsoft PowerPoint. Load PDB files, explore ligand binding sites, visualize molecular interactions, render electrostatic surfaces, and present live 3D structures — all without leaving PowerPoint.

Built with [3Dmol.js](https://3dmol.csb.pitt.edu/) and the [Office.js](https://learn.microsoft.com/office/dev/add-ins/) Add-in API.

---

## Table of Contents

- [Installation (macOS)](#installation-macos)
- [Installation (Windows)](#installation-windows)
- [Uninstallation](#uninstallation)
- [Clearing the Add-in Cache](#clearing-the-add-in-cache)
- [Getting Started](#getting-started)
- [TaskPane Controls](#taskpane-controls)
  - [Loading a Structure](#loading-a-structure)
  - [Electron Density Maps](#electron-density-maps)
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
  - [Electron Density Maps (Presenter)](#electron-density-maps-presenter)
  - [Push to Slide](#push-to-slide)
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

---

## Installation (Windows)

### Option A: Sideload via Shared Folder

1. Download `manifest-ghpages.xml` and `manifest-content-ghpages.xml` from the repository
2. Create a folder to store the manifests, for example:
   ```
   C:\OfficeAddins\
   ```
3. Copy both XML files into this folder, renaming them:
   - `manifest-ghpages.xml` → `ProteinViewer.xml`
   - `manifest-content-ghpages.xml` → `ProteinViewerContent.xml`
4. Share the folder on the network:
   - Right-click the folder → **Properties** → **Sharing** tab → **Share...**
   - Add your user and click **Share**
   - Note the network path shown (e.g. `\\YOURPC\OfficeAddins`)
5. Register the shared folder as a Trusted Catalog in PowerPoint:
   - Open PowerPoint
   - Go to **File → Options → Trust Center → Trust Center Settings...**
   - Click **Trusted Add-in Catalogs** in the left sidebar
   - In the **Catalog Url** field, enter the network path (e.g. `\\YOURPC\OfficeAddins`)
   - Click **Add catalog**
   - Check the **Show in Menu** checkbox next to the catalog you just added
   - Click **OK** to close Trust Center Settings, then **OK** again to close Options
6. Restart PowerPoint
7. Go to **Insert → My Add-ins → Shared Folder** tab
8. Select **Protein PDB Viewer** and click **Add**
9. Repeat for **Protein Viewer Slide** (the content add-in)

### Option B: Upload Directly (No Shared Folder)

1. Download `manifest-ghpages.xml` and `manifest-content-ghpages.xml` from the repository
2. Open PowerPoint
3. Go to **Insert → My Add-ins** (or **Get Add-ins** depending on your version)
4. Click **Upload My Add-in** (at the bottom of the dialog)
5. Click **Browse**, select `manifest-ghpages.xml`, and click **Upload**
6. The task pane add-in is now loaded
7. Repeat the upload for `manifest-content-ghpages.xml` to enable the in-slide interactive viewer

> **Note:** Add-ins uploaded this way need to be re-uploaded each time you restart PowerPoint. Use Option A for a persistent installation.

### Option C: Sideload to the Wef Folder

1. Download `manifest-ghpages.xml` and `manifest-content-ghpages.xml` from the repository
2. Open File Explorer and navigate to:
   ```
   %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\
   ```
   If the `Wef` folder does not exist, create it.
3. Copy both XML files into this folder, renaming them:
   - `manifest-ghpages.xml` → `ProteinViewer.xml`
   - `manifest-content-ghpages.xml` → `ProteinViewerContent.xml`
4. Restart PowerPoint
5. The **"Protein PDB Viewer"** button should appear in the **Home** tab of the ribbon

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

If you used the Shared Folder method, also remove the catalog from **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**.

Restart PowerPoint after removing the files.

---

## Clearing the Add-in Cache

If the add-in is not reflecting recent updates (e.g. after a new deployment), you may need to clear the Office webview cache. This forces PowerPoint to re-download the add-in files from GitHub Pages.

### macOS

1. Quit PowerPoint completely (**Cmd+Q**)
2. Open Terminal and run:
   ```bash
   rm -rf ~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches
   ```
3. Reopen PowerPoint

### Windows

1. Close PowerPoint completely
2. Open File Explorer and delete the contents of these folders (the folders themselves can remain):
   ```
   %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\webview2\
   ```
   ```
   %LOCALAPPDATA%\Microsoft\Office\16.0\WebCache\
   ```
   If those folders do not exist, try:
   ```
   %APPDATA%\Microsoft\Teams\Service Worker\CacheStorage\
   ```
3. Alternatively, open a Command Prompt and run:
   ```cmd
   rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\webview2"
   rmdir /s /q "%LOCALAPPDATA%\Microsoft\Office\16.0\WebCache"
   ```
4. Reopen PowerPoint

> **Tip:** If clearing the cache does not help, try restarting your computer. Some cache files may be locked while Office processes are running in the background.

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

### Electron Density Maps

After loading a PDB structure, the **Electron Density Map** section appears. You can load MTZ or CCP4 map files to visualize electron density alongside the model.

| Control | Description |
|---------|-------------|
| Load Map | Select a `.mtz`, `.ccp4`, or `.map` file |
| 2Fo-Fc σ | Sigma level for the 2Fo-Fc map (default 1.5σ, displayed as light blue mesh) |
| Show Fo-Fc | Toggle the difference map (green for +σ, red for -σ) |
| Fo-Fc σ | Sigma level for the Fo-Fc map (default 3.0σ) |
| Radius | Extraction radius in Å around the center point (default 8 Å) |
| Remove Map | Clear the density map from the viewer |

MTZ files are processed entirely in the browser using [gemmi](https://gemmi.readthedocs.io/) compiled to WebAssembly — no server-side computation is needed. The FFT from structure factors to electron density is performed client-side.

If a ligand is selected, the density map is centered on the ligand. Otherwise, it centers on the model's center of mass.

When you click **Push to Slide Viewer**, the electron density isosurface geometry is also pushed to the in-slide viewer.

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

### Electron Density Maps (Presenter)

The presenter window also supports loading electron density maps. The controls are in the **Electron Density** section of the side panel and work the same as in the task pane.

### Push to Slide

The presenter window can push all visible entries (with their styles and electron density) to the in-slide content add-in:

1. Configure your structures and density maps in the presenter
2. Click **"Push to Slide"** in the Slide section
3. All visible entries, their visualization settings, camera orientation, and electron density isosurfaces are synced to the slide

The presenter window also restores its state when reopened — previously pushed entries and their settings are preserved.

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
│       ├── glbexport.js    # Three.js GLB exporter
│       └── edmap.js        # Electron density map support (gemmi WASM)
├── server/                 # Python FastAPI dev server (local HTTPS)
│   ├── server.py
│   └── server_main.py
├── assets/                 # Add-in icons + gemmi WASM files
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
- Communication between TaskPane and Content/Presenter uses `localStorage` (same-origin) and `postMessage` for cross-context view state syncing
- The **Electron density module** uses gemmi compiled to WebAssembly to parse MTZ files and compute FFTs entirely in the browser. Isosurface extraction is also done in WASM for performance
- The **GLB exporter** uses Three.js to convert 3Dmol.js molecular geometry into a binary glTF file that PowerPoint can natively display as an interactive 3D model

### Technology Stack

| Component | Technology |
|-----------|-----------|
| 3D Rendering | [3Dmol.js](https://3dmol.csb.pitt.edu/) (WebGL) |
| Electron Density | [gemmi](https://gemmi.readthedocs.io/) (WebAssembly) — MTZ/CCP4 parsing and FFT |
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
- **Electron density on slide:** The isosurface geometry is serialized to localStorage when pushing to the slide. Very large maps or high-radius extractions may exceed the localStorage size limit (~5-10 MB) and fail silently — the PDB will still push but the density won't appear.
- **MTZ column labels:** The gemmi WASM module uses default column labels for 2Fo-Fc and Fo-Fc map calculation. Non-standard MTZ files with unusual column names may not produce maps.

---

## License

MIT
