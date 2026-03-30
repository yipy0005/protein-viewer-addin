/* global Office, $3Dmol */

import "./content.css";

let viewer = null;
let currentModel = null;
let lastHash = "";

const WATER_RESNS = ["HOH", "WAT"];
const ION_RESNS = ["NA", "CL", "MG", "ZN", "CA", "FE", "MN", "K", "SO4", "PO4"];
const RESIDUE_CHARGES = { ARG: 1, LYS: 1, HIS: 0.5, ASP: -1, GLU: -1 };

Office.onReady(() => {
  viewer = $3Dmol.createViewer("viewer", { backgroundColor: "white", antialias: true });
  checkForUpdates();
  setInterval(checkForUpdates, 500);
});

function checkForUpdates() {
  const pdbData = localStorage.getItem("proteinviewer_pdbData");
  const styleConfig = localStorage.getItem("proteinviewer_styleConfig");
  if (!pdbData) return;
  const hash = pdbData.length + "_" + (styleConfig || "");
  if (hash === lastHash) return;
  lastHash = hash;
  renderStructure(pdbData, styleConfig);
}

function mapColorScheme(s) {
  return { spectrum: "spectrum", chain: "chain", ss: "ssPyMol", residue: "amino", element: "default" }[s] || "default";
}

function assignCharges() {
  if (!currentModel) return;
  for (const atom of currentModel.selectedAtoms({})) {
    atom.charge = RESIDUE_CHARGES[atom.resn] || 0;
  }
}

function getSurfaceColorSpec(mode, colorScheme) {
  switch (mode) {
    case "esp": return { prop: "charge", gradient: new $3Dmol.Gradient.RWB(-1, 1) };
    case "protein": return { colorscheme: mapColorScheme(colorScheme) };
    case "hydrophobicity": return { prop: "charge", gradient: new $3Dmol.Gradient.Sinebow(0, 1) };
    case "element": return { colorscheme: "default" };
    default: return { color: "white" };
  }
}

function getSurfaceType(t) {
  return { VDW: $3Dmol.SurfaceType.VDW, SAS: $3Dmol.SurfaceType.SAS, SES: $3Dmol.SurfaceType.SES }[t] || $3Dmol.SurfaceType.SAS;
}

function renderStructure(pdbData, styleConfigJson) {
  if (!viewer) return;
  viewer.removeAllModels();
  viewer.removeAllSurfaces();
  viewer.removeAllShapes();
  viewer.removeAllLabels();

  currentModel = viewer.addModel(pdbData, "pdb");
  assignCharges();

  let c = {};
  try { c = styleConfigJson ? JSON.parse(styleConfigJson) : {}; } catch (e) { /**/ }

  const style = c.style || "cartoon";
  const colorscheme = mapColorScheme(c.colorScheme || "spectrum");
  const bg = c.backgroundColor || "white";
  const proteinOpacity = c.proteinOpacity !== undefined ? c.proteinOpacity : 1.0;

  const styleObj = {};
  styleObj[style] = { colorscheme, opacity: proteinOpacity };
  viewer.setStyle({ not: { hetflag: true } }, styleObj);

  // Full protein surface
  if (c.showSurface) {
    const sOp = c.surfaceOpacity !== undefined ? c.surfaceOpacity : 0.6;
    const sType = getSurfaceType(c.surfaceType || "SAS");
    const sColor = getSurfaceColorSpec(c.surfaceColor || "white", c.colorScheme);
    viewer.addSurface(sType, { opacity: sOp, ...sColor }, { not: { hetflag: true } });
  }

  // Ligand
  const lig = c.selectedLigand;
  if (lig) {
    const resiInt = parseInt(lig.resi, 10);
    const ligSel = { resn: lig.resn, chain: lig.chain, resi: resiInt };
    viewer.addStyle(ligSel, buildLigandStyle(c.ligandStyle || "ball-and-stick"));

    if (c.showBindingSite) {
      viewer.render();
      renderBindingSite(ligSel, c.bindingDistance || 5, c.showBindingLabels, c.showHbonds, c.showBindingSurface, c.bindingSurfaceColor, c.bindingSurfaceOpacity, c.colorScheme);
    }

    viewer.zoomTo(c.zoomToLigand ? ligSel : undefined);
  } else {
    viewer.addStyle(
      { hetflag: true, not: { resn: [...WATER_RESNS, ...ION_RESNS] } },
      { stick: { colorscheme: "default", radius: 0.15 }, sphere: { colorscheme: "default", radius: 0.3 } }
    );
    viewer.zoomTo();
  }

  viewer.setBackgroundColor(bg);
  viewer.render();
}

function buildLigandStyle(s) {
  switch (s) {
    case "ball-and-stick": return { stick: { colorscheme: "greenCarbon", radius: 0.15 }, sphere: { colorscheme: "greenCarbon", radius: 0.35 } };
    case "stick": return { stick: { colorscheme: "greenCarbon", radius: 0.2 } };
    case "sphere": return { sphere: { colorscheme: "greenCarbon" } };
    default: return { stick: { colorscheme: "greenCarbon", radius: 0.15 } };
  }
}

function renderBindingSite(ligSel, dist, showLabels, showHbonds, showBsSurface, bsColorMode, bsOpacity, colorScheme) {
  const ligandAtoms = currentModel.selectedAtoms(ligSel);
  const proteinAtoms = currentModel.selectedAtoms({ not: { hetflag: true } });
  if (!ligandAtoms || !ligandAtoms.length || !proteinAtoms || !proteinAtoms.length) return;

  const nearbyResidues = new Map();
  for (const la of ligandAtoms) {
    for (const pa of proteinAtoms) {
      const dx = la.x - pa.x, dy = la.y - pa.y, dz = la.z - pa.z;
      if (Math.sqrt(dx * dx + dy * dy + dz * dz) <= dist) {
        const key = `${pa.chain}:${pa.resi}`;
        if (!nearbyResidues.has(key)) nearbyResidues.set(key, { chain: pa.chain, resi: pa.resi });
      }
    }
  }
  if (nearbyResidues.size === 0) return;

  const chainResiMap = {};
  const allResi = [];
  for (const [, r] of nearbyResidues) {
    if (!chainResiMap[r.chain]) chainResiMap[r.chain] = [];
    chainResiMap[r.chain].push(r.resi);
    allResi.push(r.resi);
  }

  for (const [chain, residues] of Object.entries(chainResiMap)) {
    viewer.addStyle({ chain, resi: residues, not: { hetflag: true } }, { stick: { colorscheme: "default", radius: 0.12 } });
  }

  if (showLabels) {
    for (const [chain, residues] of Object.entries(chainResiMap)) {
      viewer.addResLabels({ chain, resi: residues, atom: "CA" },
        { font: "Arial", fontSize: 10, showBackground: true, backgroundColor: 0x333333, backgroundOpacity: 0.8, fontColor: "white" });
    }
  }

  if (showHbonds) {
    const donors = new Set(["N", "O", "S"]);
    for (const la of ligandAtoms) {
      if (!donors.has(la.elem)) continue;
      for (const pa of proteinAtoms) {
        if (!donors.has(pa.elem)) continue;
        if (!nearbyResidues.has(`${pa.chain}:${pa.resi}`)) continue;
        const dx = la.x - pa.x, dy = la.y - pa.y, dz = la.z - pa.z;
        const d = Math.sqrt(dx * dx + dy * dy + dz * dz);
        if (d >= 2.0 && d <= 3.5) {
          viewer.addCylinder({
            start: { x: la.x, y: la.y, z: la.z }, end: { x: pa.x, y: pa.y, z: pa.z },
            radius: 0.04, color: "yellow", fromCap: true, toCap: true,
            dashed: true, dashLength: 0.15, gapLength: 0.1,
          });
        }
      }
    }
  }

  if (showBsSurface) {
    const op = bsOpacity !== undefined ? bsOpacity : 0.5;
    const cSpec = getSurfaceColorSpec(bsColorMode || "esp", colorScheme);
    viewer.addSurface($3Dmol.SurfaceType.SAS, { opacity: op, ...cSpec },
      { resi: allResi, not: { hetflag: true } },
      { not: { hetflag: true } });
  }
}
