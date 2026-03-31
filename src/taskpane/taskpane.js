/* global Office, $3Dmol, Gemmi */

import "./taskpane.css";
import { exportToGLB, downloadGLB } from "../viewer/glbexport.js";
import { ensureGemmi, parseMtz, parseCcp4, renderDensityMap, removeDensityMap } from "../viewer/edmap.js";

let viewer = null;
let currentModel = null;
let currentPdbData = null;
let currentStyle = "cartoon";
let currentColorScheme = "spectrum";
let currentMapData = null;
let currentBg = "white";
let detectedLigands = [];

const WATER_RESNS = ["HOH", "WAT"];
const ION_RESNS = ["NA", "CL", "MG", "ZN", "CA", "FE", "MN", "K", "SO4", "PO4"];
const SKIP_RESNS = new Set([...WATER_RESNS, ...ION_RESNS]);

Office.onReady(() => {
  initViewer();
  bindEvents();
  setStatus("Ready — load a PDB file or fetch by ID.", "success");
});

function initViewer() {
  viewer = $3Dmol.createViewer("viewer-container", { backgroundColor: "white", antialias: true });
}

function setStatus(msg, type) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status-text" + (type ? " " + type : "");
}

function showLoading(text) {
  document.getElementById("loading-text").textContent = text || "Loading...";
  document.getElementById("loading-overlay").style.display = "flex";
}
function hideLoading() {
  document.getElementById("loading-overlay").style.display = "none";
}

function mapColorScheme(scheme) {
  return { spectrum: "spectrum", chain: "chain", ss: "ssPyMol", residue: "amino", element: "default" }[scheme] || "default";
}

// --- Residue-level charge assignment for ESP coloring ---
// Assigns a charge value to each atom based on its residue type.
// Positive residues (ARG, LYS, HIS) → +1, Negative (ASP, GLU) → -1, Neutral → 0
const RESIDUE_CHARGES = {
  ARG: 1, LYS: 1, HIS: 0.5,
  ASP: -1, GLU: -1,
  // everything else is 0
};

function assignCharges() {
  if (!currentModel) return;
  for (const atom of currentModel.selectedAtoms({})) {
    atom.charge = RESIDUE_CHARGES[atom.resn] || 0;
  }
}

function getSurfaceColorSpec(colorMode) {
  switch (colorMode) {
    case "esp":
      return { prop: "charge", gradient: new $3Dmol.Gradient.RWB(-1, 1) };
    case "protein":
      return { colorscheme: mapColorScheme(currentColorScheme) };
    case "hydrophobicity":
      return { prop: "charge", gradient: new $3Dmol.Gradient.Sinebow(0, 1) };
    case "element":
      return { colorscheme: "default" };
    case "white":
    default:
      return { color: "white" };
  }
}

function getSurfaceType(typeName) {
  const map = {
    VDW: $3Dmol.SurfaceType.VDW,
    SAS: $3Dmol.SurfaceType.SAS,
    SES: $3Dmol.SurfaceType.SES,
  };
  return map[typeName] || $3Dmol.SurfaceType.SAS;
}

// --- Ligand detection (parses PDB text directly, no model dependency) ---
function detectLigandsFromPdb(pdbText) {
  const ligandMap = {};
  const lines = pdbText.split("\n");
  for (const line of lines) {
    if (!line.startsWith("HETATM")) continue;
    const resn = line.substring(17, 20).trim();
    if (SKIP_RESNS.has(resn)) continue;
    const chain = line.substring(21, 22).trim();
    const resi = line.substring(22, 26).trim();
    const key = `${resn}_${chain}_${resi}`;
    if (!ligandMap[key]) {
      ligandMap[key] = { resn, chain, resi, count: 0 };
    }
    ligandMap[key].count++;
  }
  return Object.values(ligandMap);
}

function populateLigandDropdown() {
  const sel = document.getElementById("ligand-select");
  sel.innerHTML = '<option value="">None</option>';
  detectedLigands = detectLigandsFromPdb(currentPdbData);

  if (detectedLigands.length === 0) {
    document.getElementById("ligand-section").style.display = "none";
    return;
  }

  document.getElementById("ligand-section").style.display = "";
  for (const lig of detectedLigands) {
    const opt = document.createElement("option");
    opt.value = JSON.stringify({ resn: lig.resn, chain: lig.chain, resi: lig.resi });
    opt.textContent = `${lig.resn} (Chain ${lig.chain || "?"}, #${lig.resi}) — ${lig.count} atoms`;
    sel.appendChild(opt);
  }

  // Auto-select first ligand if only one
  if (detectedLigands.length === 1) {
    sel.selectedIndex = 1;
  }
}

// --- Visualization ---
function applyVisualization() {
  if (!viewer || !currentPdbData) return;

  viewer.removeAllModels();
  viewer.removeAllSurfaces();
  viewer.removeAllShapes();
  viewer.removeAllLabels();

  currentModel = viewer.addModel(currentPdbData, "pdb");

  // Assign charges for ESP coloring
  assignCharges();

  // Protein style
  const colorscheme = mapColorScheme(currentColorScheme);
  const proteinOpacity = parseInt(document.getElementById("protein-opacity").value) / 100;
  const proteinStyleObj = {};
  proteinStyleObj[currentStyle] = { colorscheme, opacity: proteinOpacity };

  // Apply protein style to all non-HETATM atoms
  viewer.setStyle({ not: { hetflag: true } }, proteinStyleObj);

  // Surface
  const showSurface = document.getElementById("surface-toggle").checked;
  document.getElementById("surface-options").style.display = showSurface ? "" : "none";
  if (showSurface) {
    const surfaceOpacity = parseInt(document.getElementById("surface-opacity").value) / 100;
    const surfaceTypeName = document.getElementById("surface-type").value;
    const surfaceColorMode = document.getElementById("surface-color").value;
    let surfType, colorSpec;
    if (surfaceTypeName === "ESP") {
      surfType = $3Dmol.SurfaceType.SAS;
      colorSpec = { prop: "charge", gradient: new $3Dmol.Gradient.RWB(-1, 1) };
    } else {
      surfType = getSurfaceType(surfaceTypeName);
      colorSpec = getSurfaceColorSpec(surfaceColorMode);
    }
    viewer.addSurface(surfType, { opacity: surfaceOpacity, ...colorSpec }, { not: { hetflag: true } });
  }

  // Read ligand controls
  const ligandVal = document.getElementById("ligand-select").value;
  const ligandStyleName = document.getElementById("ligand-style").value;
  const zoomToLigand = document.getElementById("chk-zoom-ligand").checked;
  const showBindingSite = document.getElementById("chk-binding-site").checked;
  const showLabels = document.getElementById("chk-binding-labels").checked;
  const showHbonds = document.getElementById("chk-hbonds").checked;
  const bindingDist = parseFloat(document.getElementById("binding-dist").value);

  document.getElementById("binding-dist-row").style.display = showBindingSite ? "" : "none";

  if (ligandVal) {
    const lig = JSON.parse(ligandVal);
    // Use integer resi for 3Dmol selection
    const resiInt = parseInt(lig.resi, 10);
    const ligSel = { resn: lig.resn, chain: lig.chain, resi: resiInt };

    // Style the selected ligand
    viewer.addStyle(ligSel, buildLigandStyle(ligandStyleName));

    // Binding site — must render first so atom coords are available
    if (showBindingSite) {
      // Force a render so the model geometry is computed
      viewer.render();
      renderBindingSite(ligSel, bindingDist, showLabels, showHbonds);
    }

    if (zoomToLigand) {
      viewer.zoomTo(ligSel);
    } else {
      viewer.zoomTo();
    }
  } else {
    // Show all non-water/ion heteroatoms
    viewer.addStyle(
      { hetflag: true, not: { resn: [...WATER_RESNS, ...ION_RESNS] } },
      { stick: { colorscheme: "default", radius: 0.15 }, sphere: { colorscheme: "default", radius: 0.3 } }
    );
    viewer.zoomTo();
  }

  viewer.setBackgroundColor(currentBg);
  viewer.render();
}

function buildLigandStyle(style) {
  switch (style) {
    case "ball-and-stick":
      return { stick: { colorscheme: "greenCarbon", radius: 0.15 }, sphere: { colorscheme: "greenCarbon", radius: 0.35 } };
    case "stick":
      return { stick: { colorscheme: "greenCarbon", radius: 0.2 } };
    case "sphere":
      return { sphere: { colorscheme: "greenCarbon" } };
    default:
      return { stick: { colorscheme: "greenCarbon", radius: 0.15 } };
  }
}

function renderBindingSite(ligSel, dist, showLabels, showHbonds) {
  const ligandAtoms = currentModel.selectedAtoms(ligSel);
  const proteinAtoms = currentModel.selectedAtoms({ not: { hetflag: true } });
  const showSaltBridges = document.getElementById("chk-salt-bridges").checked;
  const showPiStacking = document.getElementById("chk-pi-stacking").checked;
  const showPiCation = document.getElementById("chk-pi-cation").checked;

  if (!ligandAtoms || ligandAtoms.length === 0) return;
  if (!proteinAtoms || proteinAtoms.length === 0) return;

  const nearbyResidues = new Map();
  for (const la of ligandAtoms) {
    for (const pa of proteinAtoms) {
      const dx = la.x - pa.x, dy = la.y - pa.y, dz = la.z - pa.z;
      const d = Math.sqrt(dx * dx + dy * dy + dz * dz);
      if (d <= dist) {
        const key = `${pa.chain}:${pa.resi}`;
        if (!nearbyResidues.has(key)) {
          nearbyResidues.set(key, { chain: pa.chain, resi: pa.resi, resn: pa.resn });
        }
      }
    }
  }

  if (nearbyResidues.size === 0) return;

  const chainResiMap = {};
  const allResiList = [];
  for (const [, res] of nearbyResidues) {
    if (!chainResiMap[res.chain]) chainResiMap[res.chain] = [];
    chainResiMap[res.chain].push(res.resi);
    allResiList.push(res.resi);
  }

  // Style binding site residues as sticks
  for (const [chain, residues] of Object.entries(chainResiMap)) {
    viewer.addStyle(
      { chain, resi: residues, not: { hetflag: true } },
      { stick: { colorscheme: "default", radius: 0.12 } }
    );
  }

  // Labels
  if (showLabels) {
    for (const [chain, residues] of Object.entries(chainResiMap)) {
      viewer.addResLabels(
        { chain, resi: residues, atom: "CA" },
        { font: "Arial", fontSize: 10, showBackground: true, backgroundColor: 0x333333, backgroundOpacity: 0.8, fontColor: "white" }
      );
    }
  }

  // Collect nearby protein atoms for interaction checks
  const nearbyProteinAtoms = proteinAtoms.filter(
    (a) => nearbyResidues.has(`${a.chain}:${a.resi}`)
  );

  // H-bonds: N/O/S within 2.0–3.5 Å (yellow dashed)
  if (showHbonds) {
    const hbDonors = new Set(["N", "O", "S"]);
    for (const la of ligandAtoms) {
      if (!hbDonors.has(la.elem)) continue;
      for (const pa of nearbyProteinAtoms) {
        if (!hbDonors.has(pa.elem)) continue;
        const d = atomDist(la, pa);
        if (d >= 2.0 && d <= 3.5) {
          addDashedLine(la, pa, "yellow");
        }
      }
    }
  }

  // Salt bridges: charged groups within 4.0 Å (magenta dashed)
  if (showSaltBridges) {
    const posResAtoms = new Set(["NZ", "NH1", "NH2", "NE"]);  // LYS NZ, ARG NH1/NH2/NE
    const negResAtoms = new Set(["OD1", "OD2", "OE1", "OE2"]); // ASP OD, GLU OE
    const chargedLigAtoms = ligandAtoms.filter((a) => a.elem === "N" || a.elem === "O");

    for (const la of chargedLigAtoms) {
      for (const pa of nearbyProteinAtoms) {
        const isPosProt = posResAtoms.has(pa.atom) && ["ARG", "LYS"].includes(pa.resn);
        const isNegProt = negResAtoms.has(pa.atom) && ["ASP", "GLU"].includes(pa.resn);
        if (!isPosProt && !isNegProt) continue;
        const d = atomDist(la, pa);
        if (d >= 1.5 && d <= 4.0) {
          addDashedLine(la, pa, "magenta");
        }
      }
    }
  }

  // π–π stacking: aromatic ring centroids within 5.5 Å (cyan dashed)
  if (showPiStacking) {
    const ligRings = findAromaticRings(ligandAtoms);
    const protRings = findProteinAromaticRings(nearbyProteinAtoms);
    for (const lr of ligRings) {
      for (const pr of protRings) {
        const d = centroidDist(lr.centroid, pr.centroid);
        if (d <= 5.5) {
          addDashedLine(lr.centroid, pr.centroid, "cyan");
        }
      }
    }
  }

  // π–cation: aromatic ring centroid to cation within 6.0 Å (orange dashed)
  if (showPiCation) {
    const ligRings = findAromaticRings(ligandAtoms);
    const protRings = findProteinAromaticRings(nearbyProteinAtoms);
    // Protein cations near ligand rings
    const protCations = nearbyProteinAtoms.filter(
      (a) => (a.atom === "NZ" && a.resn === "LYS") ||
             (["NH1", "NH2"].includes(a.atom) && a.resn === "ARG")
    );
    // Ligand cations (N atoms) near protein rings
    const ligCations = ligandAtoms.filter((a) => a.elem === "N");

    for (const ring of ligRings) {
      for (const cat of protCations) {
        if (centroidDist(ring.centroid, cat) <= 6.0) {
          addDashedLine(ring.centroid, cat, "orange");
        }
      }
    }
    for (const ring of protRings) {
      for (const cat of ligCations) {
        if (centroidDist(ring.centroid, cat) <= 6.0) {
          addDashedLine(ring.centroid, cat, "orange");
        }
      }
    }
  }

  // Binding site surface
  const showBsSurface = document.getElementById("chk-binding-surface").checked;
  document.getElementById("binding-surface-options").style.display = showBsSurface ? "" : "none";
  if (showBsSurface) {
    const bsColorMode = document.getElementById("binding-surface-color").value;
    const bsOpacity = parseInt(document.getElementById("binding-surface-opacity").value) / 100;
    const colorSpec = getSurfaceColorSpec(bsColorMode);
    viewer.addSurface($3Dmol.SurfaceType.SAS, {
      opacity: bsOpacity,
      ...colorSpec,
    }, { resi: allResiList, not: { hetflag: true } }, { not: { hetflag: true } });
  }
}

// --- Geometry helpers ---

function atomDist(a, b) {
  const dx = a.x - b.x, dy = a.y - b.y, dz = a.z - b.z;
  return Math.sqrt(dx * dx + dy * dy + dz * dz);
}

function centroidDist(a, b) {
  const dx = a.x - b.x, dy = a.y - b.y, dz = a.z - b.z;
  return Math.sqrt(dx * dx + dy * dy + dz * dz);
}

function addDashedLine(a, b, color) {
  viewer.addCylinder({
    start: { x: a.x, y: a.y, z: a.z },
    end: { x: b.x, y: b.y, z: b.z },
    radius: 0.04, color: color,
    fromCap: true, toCap: true,
    dashed: true, dashLength: 0.15, gapLength: 0.1,
  });
}

// Aromatic ring detection for ligand atoms (6-membered rings with planar C/N)
function findAromaticRings(atoms) {
  // Group atoms by connectivity — use a simple heuristic:
  // Find sets of 5-6 C/N atoms that are all within ~1.7 Å of at least 2 others in the set
  // Then compute centroid. This is approximate but works for most drug-like ligands.
  const aromatic = atoms.filter((a) => ["C", "N"].includes(a.elem));
  const rings = [];

  // Build adjacency
  const adj = new Map();
  for (const a of aromatic) {
    adj.set(a, []);
  }
  for (let i = 0; i < aromatic.length; i++) {
    for (let j = i + 1; j < aromatic.length; j++) {
      if (atomDist(aromatic[i], aromatic[j]) <= 1.7) {
        adj.get(aromatic[i]).push(aromatic[j]);
        adj.get(aromatic[j]).push(aromatic[i]);
      }
    }
  }

  // Find 5- and 6-membered rings via DFS
  const found = new Set();
  for (const start of aromatic) {
    findRingsDFS(start, start, [start], adj, found, rings, 6);
  }

  return rings;
}

function findRingsDFS(start, current, path, adj, found, rings, maxLen) {
  if (path.length > maxLen) return;
  for (const neighbor of adj.get(current) || []) {
    if (neighbor === start && path.length >= 5) {
      const key = path.map((a) => `${a.x.toFixed(1)},${a.y.toFixed(1)}`).sort().join("|");
      if (!found.has(key)) {
        found.add(key);
        const cx = path.reduce((s, a) => s + a.x, 0) / path.length;
        const cy = path.reduce((s, a) => s + a.y, 0) / path.length;
        const cz = path.reduce((s, a) => s + a.z, 0) / path.length;
        rings.push({ centroid: { x: cx, y: cy, z: cz }, atoms: [...path] });
      }
      return;
    }
    if (path.includes(neighbor)) continue;
    path.push(neighbor);
    findRingsDFS(start, neighbor, path, adj, found, rings, maxLen);
    path.pop();
  }
}

function findProteinAromaticRings(atoms) {
  // Known aromatic residues and their ring atoms
  const aromaticRes = {
    PHE: ["CG", "CD1", "CD2", "CE1", "CE2", "CZ"],
    TYR: ["CG", "CD1", "CD2", "CE1", "CE2", "CZ"],
    TRP: ["CG", "CD1", "CD2", "NE1", "CE2", "CE3", "CZ2", "CZ3", "CH2"],
    HIS: ["CG", "ND1", "CD2", "CE1", "NE2"],
  };

  // Group by residue
  const resGroups = {};
  for (const a of atoms) {
    if (!aromaticRes[a.resn]) continue;
    const key = `${a.chain}:${a.resi}:${a.resn}`;
    if (!resGroups[key]) resGroups[key] = { resn: a.resn, atoms: [] };
    resGroups[key].atoms.push(a);
  }

  const rings = [];
  for (const [, group] of Object.entries(resGroups)) {
    const ringAtomNames = aromaticRes[group.resn];
    // For TRP, find two rings: 5-membered (CG,CD1,NE1,CE2,CD2) and 6-membered (CD2,CE2,CE3,CZ3,CH2,CZ2)
    if (group.resn === "TRP") {
      const ring5names = ["CG", "CD1", "NE1", "CE2", "CD2"];
      const ring6names = ["CD2", "CE2", "CE3", "CZ3", "CH2", "CZ2"];
      for (const rnames of [ring5names, ring6names]) {
        const rAtoms = rnames.map((n) => group.atoms.find((a) => a.atom === n)).filter(Boolean);
        if (rAtoms.length >= 4) {
          rings.push({ centroid: centroidOf(rAtoms), atoms: rAtoms });
        }
      }
    } else {
      const rAtoms = ringAtomNames.map((n) => group.atoms.find((a) => a.atom === n)).filter(Boolean);
      if (rAtoms.length >= 4) {
        rings.push({ centroid: centroidOf(rAtoms), atoms: rAtoms });
      }
    }
  }
  return rings;
}

function centroidOf(atoms) {
  const n = atoms.length;
  return {
    x: atoms.reduce((s, a) => s + a.x, 0) / n,
    y: atoms.reduce((s, a) => s + a.y, 0) / n,
    z: atoms.reduce((s, a) => s + a.z, 0) / n,
  };
}

// --- Load & Info ---
function loadPdbData(data) {
  currentPdbData = data;
  populateLigandDropdown();
  applyVisualization();
  extractInfo(data);
  document.getElementById("btn-insert").disabled = false;
  document.getElementById("btn-push-slide").disabled = false;
  document.getElementById("btn-download-glb").disabled = false;
  document.getElementById("map-section").style.display = "";
  setStatus("Structure loaded.", "success");
}

function extractInfo(pdbText) {
  const lines = pdbText.split("\n");
  let title = "", atoms = 0;
  const chains = new Set(), residues = new Set();
  for (const line of lines) {
    if (line.startsWith("TITLE")) title += line.substring(10).trim() + " ";
    if (line.startsWith("ATOM") || line.startsWith("HETATM")) {
      atoms++;
      const chain = line.substring(21, 22).trim();
      if (chain) chains.add(chain);
      residues.add(chain + line.substring(17, 20).trim() + line.substring(22, 26).trim());
    }
  }
  document.getElementById("info-content").innerHTML = `
    ${title ? `<div><strong>Title:</strong> ${title.trim()}</div>` : ""}
    <div><strong>Atoms:</strong> ${atoms.toLocaleString()}</div>
    <div><strong>Chains:</strong> ${chains.size} (${[...chains].join(", ")})</div>
    <div><strong>Residues:</strong> ${residues.size.toLocaleString()}</div>
    <div><strong>Ligands:</strong> ${detectedLigands.length > 0 ? detectedLigands.map(l => l.resn).join(", ") : "None"}</div>
  `;
  document.getElementById("protein-info").style.display = "";
}

function getStyleConfig() {
  const ligandVal = document.getElementById("ligand-select").value;
  return {
    style: currentStyle,
    colorScheme: currentColorScheme,
    backgroundColor: currentBg,
    proteinOpacity: parseInt(document.getElementById("protein-opacity").value) / 100,
    showSurface: document.getElementById("surface-toggle").checked,
    surfaceType: document.getElementById("surface-type").value,
    surfaceColor: document.getElementById("surface-color").value,
    surfaceOpacity: parseInt(document.getElementById("surface-opacity").value) / 100,
    ligandStyle: document.getElementById("ligand-style").value,
    selectedLigand: ligandVal ? JSON.parse(ligandVal) : null,
    zoomToLigand: document.getElementById("chk-zoom-ligand").checked,
    showBindingSite: document.getElementById("chk-binding-site").checked,
    bindingDistance: parseFloat(document.getElementById("binding-dist").value),
    showBindingLabels: document.getElementById("chk-binding-labels").checked,
    showHbonds: document.getElementById("chk-hbonds").checked,
    showSaltBridges: document.getElementById("chk-salt-bridges").checked,
    showPiStacking: document.getElementById("chk-pi-stacking").checked,
    showPiCation: document.getElementById("chk-pi-cation").checked,
    showBindingSurface: document.getElementById("chk-binding-surface").checked,
    bindingSurfaceColor: document.getElementById("binding-surface-color").value,
    bindingSurfaceOpacity: parseInt(document.getElementById("binding-surface-opacity").value) / 100,
  };
}

function pushToSlideViewer() {
  if (!currentPdbData) return;
  localStorage.setItem("proteinviewer_pdbData", currentPdbData);
  localStorage.setItem("proteinviewer_styleConfig", JSON.stringify(getStyleConfig()));
  setStatus("Pushed to slide viewer.", "success");
}

// --- Event Binding ---
function bindEvents() {
  document.querySelectorAll('input[name="source"]').forEach((r) => {
    r.addEventListener("change", () => {
      const src = document.querySelector('input[name="source"]:checked').value;
      document.getElementById("source-pdb").style.display = src === "pdb" ? "" : "none";
      document.getElementById("source-file").style.display = src === "file" ? "" : "none";
    });
  });
  document.getElementById("btn-fetch").addEventListener("click", handleFetchPdb);
  document.getElementById("pdb-input").addEventListener("keydown", (e) => { if (e.key === "Enter") handleFetchPdb(); });
  document.getElementById("btn-load-file").addEventListener("click", handleLoadFile);

  // Viz controls
  document.getElementById("style-select").addEventListener("change", (e) => { currentStyle = e.target.value; applyVisualization(); });
  document.getElementById("color-scheme").addEventListener("change", (e) => { currentColorScheme = e.target.value; applyVisualization(); });
  document.getElementById("surface-toggle").addEventListener("change", () => applyVisualization());
  document.getElementById("surface-type").addEventListener("change", () => applyVisualization());
  document.getElementById("surface-color").addEventListener("change", () => applyVisualization());
  document.getElementById("surface-opacity").addEventListener("input", () => {
    document.getElementById("surface-opacity-val").textContent = document.getElementById("surface-opacity").value + "%";
  });
  document.getElementById("surface-opacity").addEventListener("change", () => applyVisualization());
  document.getElementById("protein-opacity").addEventListener("input", () => {
    document.getElementById("protein-opacity-val").textContent = document.getElementById("protein-opacity").value + "%";
  });
  document.getElementById("protein-opacity").addEventListener("change", () => applyVisualization());
  document.getElementById("spin-toggle").addEventListener("change", (e) => { if (viewer) viewer.spin(e.target.checked ? "y" : false); });
  document.getElementById("bg-select").addEventListener("change", (e) => { currentBg = e.target.value; applyVisualization(); });

  // Ligand controls
  const ligandCtrls = ["ligand-select", "ligand-style", "chk-zoom-ligand", "chk-binding-site", "chk-binding-labels", "chk-hbonds", "chk-salt-bridges", "chk-pi-stacking", "chk-pi-cation", "chk-binding-surface", "binding-surface-color"];
  ligandCtrls.forEach((id) => document.getElementById(id).addEventListener("change", () => applyVisualization()));
  document.getElementById("binding-dist").addEventListener("input", () => {
    document.getElementById("binding-dist-val").textContent = parseFloat(document.getElementById("binding-dist").value).toFixed(1) + " Å";
  });
  document.getElementById("binding-dist").addEventListener("change", () => applyVisualization());
  document.getElementById("binding-surface-opacity").addEventListener("input", () => {
    document.getElementById("binding-surface-opacity-val").textContent = document.getElementById("binding-surface-opacity").value + "%";
  });
  document.getElementById("binding-surface-opacity").addEventListener("change", () => applyVisualization());

  document.getElementById("btn-insert").addEventListener("click", handleInsertSnapshot);
  document.getElementById("btn-push-slide").addEventListener("click", pushToSlideViewer);
  let presenterWin = null;
  document.getElementById("btn-open-presenter").addEventListener("click", () => {
    if (presenterWin && !presenterWin.closed) {
      // Relay the current view state from localStorage to the presenter via postMessage
      const vs = localStorage.getItem("proteinviewer_viewState");
      if (vs) presenterWin.postMessage({ type: "proteinviewer_viewState", viewState: vs }, "*");
      presenterWin.focus();
    } else {
      const v = Date.now();
      presenterWin = window.open(`https://yipy0005.github.io/protein-viewer-addin/presenter.html?v=${v}`, "ProteinPresenter", "width=1200,height=800");
    }
  });
  document.getElementById("btn-download-glb").addEventListener("click", handleDownloadGLB);

  // Map controls
  document.getElementById("btn-load-map").addEventListener("click", handleLoadMap);
  document.getElementById("btn-remove-map").addEventListener("click", handleRemoveMap);
  document.getElementById("map-2fofc-sigma").addEventListener("input", (e) => {
    document.getElementById("map-2fofc-sigma-val").textContent = parseFloat(e.target.value).toFixed(1) + "σ";
  });
  document.getElementById("map-2fofc-sigma").addEventListener("change", reRenderMap);
  document.getElementById("chk-fofc-map").addEventListener("change", () => {
    document.getElementById("fofc-controls").style.display = document.getElementById("chk-fofc-map").checked ? "" : "none";
    reRenderMap();
  });
  document.getElementById("map-fofc-sigma").addEventListener("input", (e) => {
    document.getElementById("map-fofc-sigma-val").textContent = parseFloat(e.target.value).toFixed(1) + "σ";
  });
  document.getElementById("map-fofc-sigma").addEventListener("change", reRenderMap);
  document.getElementById("map-radius").addEventListener("input", (e) => {
    document.getElementById("map-radius-val").textContent = e.target.value + " Å";
  });
  document.getElementById("map-radius").addEventListener("change", reRenderMap);
}

async function handleFetchPdb() {
  const pdbId = document.getElementById("pdb-input").value.trim().toUpperCase();
  if (!pdbId || pdbId.length !== 4) { setStatus("Enter a valid 4-character PDB ID.", "error"); return; }
  showLoading(`Fetching ${pdbId}...`);
  try {
    const resp = await fetch(`https://files.rcsb.org/download/${pdbId}.pdb`);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    loadPdbData(await resp.text());
  } catch (err) { setStatus(`Failed to fetch ${pdbId}: ${err.message}`, "error"); }
  finally { hideLoading(); }
}

function handleLoadFile() {
  const fi = document.getElementById("pdb-file");
  if (!fi.files.length) { setStatus("Select a file.", "error"); return; }
  showLoading("Reading file...");
  const reader = new FileReader();
  reader.onload = (ev) => { loadPdbData(ev.target.result); hideLoading(); };
  reader.onerror = () => { setStatus("Failed to read file.", "error"); hideLoading(); };
  reader.readAsText(fi.files[0]);
}

async function handleInsertSnapshot() {
  if (!viewer || !currentPdbData) { setStatus("No structure loaded.", "error"); return; }
  showLoading("Capturing...");
  try {
    const wasSpinning = document.getElementById("spin-toggle").checked;
    if (wasSpinning) viewer.spin(false);
    const pngUri = viewer.pngURI();
    if (wasSpinning) viewer.spin("y");
    const base64 = pngUri.replace("data:image/png;base64,", "");
    await new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(base64, { coercionType: Office.CoercionType.Image }, (r) => {
        r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error.message));
      });
    });
    setStatus("Snapshot inserted.", "success");
  } catch (err) { setStatus("Insert failed: " + err.message, "error"); }
  finally { hideLoading(); }
}


async function handleDownloadGLB() {
  if (!viewer || !currentPdbData) { setStatus("No structure loaded.", "error"); return; }
  const glbStatus = document.getElementById("glb-status");
  glbStatus.textContent = "";
  showLoading("Building 3D model...");
  try {
    const glbBuffer = await exportToGLB(viewer);
    const pdbId = document.getElementById("pdb-input").value.trim().toUpperCase();
    const filename = pdbId ? `${pdbId}.glb` : "molecule.glb";
    downloadGLB(glbBuffer, filename);
    glbStatus.textContent = `Downloaded ${filename}. Use Insert → 3D Models to add it.`;
    glbStatus.className = "status-text success";
  } catch (err) {
    glbStatus.textContent = "Error: " + err.message;
    glbStatus.className = "status-text error";
  } finally {
    hideLoading();
  }
}

function handleLoadMap() {
  const fileInput = document.getElementById("map-file");
  const mapStatus = document.getElementById("map-status");
  if (!fileInput.files.length) { mapStatus.textContent = "Select a map file."; mapStatus.className = "status-text error"; return; }
  if (!viewer) { mapStatus.textContent = "Load a PDB first."; mapStatus.className = "status-text error"; return; }

  const file = fileInput.files[0];
  const isMtz = /\.mtz$/i.test(file.name);
  const reader = new FileReader();
  mapStatus.textContent = "Loading map...";
  mapStatus.className = "status-text";

  reader.onload = function (e) {
    ensureGemmi(function (gemmi) {
      try {
        if (isMtz) {
          const { map2fofc, mapFofc } = parseMtz(gemmi, e.target.result);
          currentMapData = { map2fofc, mapFofc };
        } else {
          const ccp4Data = parseCcp4(gemmi, e.target.result);
          currentMapData = ccp4Data;
        }
        reRenderMap();
        document.getElementById("map-controls").style.display = "";
        mapStatus.textContent = "Map loaded: " + file.name;
        mapStatus.className = "status-text success";
      } catch (err) {
        mapStatus.textContent = "Error: " + err.message;
        mapStatus.className = "status-text error";
      }
    });
  };
  reader.readAsArrayBuffer(file);
}

function getSelectedLigandCenterTaskpane() {
  const ligVal = document.getElementById("ligand-select").value;
  if (!ligVal || !viewer || !currentModel) return null;
  const lig = JSON.parse(ligVal);
  const resi = parseInt(lig.resi, 10);
  const atoms = currentModel.selectedAtoms({ resn: lig.resn, chain: lig.chain, resi: resi });
  if (!atoms || !atoms.length) return null;
  let cx = 0, cy = 0, cz = 0;
  for (const a of atoms) { cx += a.x; cy += a.y; cz += a.z; }
  const n = atoms.length;
  return { x: cx / n, y: cy / n, z: cz / n };
}

function reRenderMap() {
  if (!viewer || !currentMapData) return;
  const view = viewer.getView();
  removeDensityMap(viewer);
  const sigma2fofc = parseFloat(document.getElementById("map-2fofc-sigma").value);
  const sigmaFofc = parseFloat(document.getElementById("map-fofc-sigma").value);
  const showFofc = document.getElementById("chk-fofc-map").checked;
  const radius = parseFloat(document.getElementById("map-radius").value);
  const center = getSelectedLigandCenterTaskpane();
  renderDensityMap(viewer, currentMapData, { sigma2fofc, sigmaFofc, showFofc, radius, center });
  if (view) viewer.setView(view);
}

function handleRemoveMap() {
  if (viewer) removeDensityMap(viewer);
  currentMapData = null;
  document.getElementById("map-controls").style.display = "none";
  document.getElementById("map-status").textContent = "";
  document.getElementById("map-file").value = "";
}
