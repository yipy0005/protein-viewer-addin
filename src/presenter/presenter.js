/* global $3Dmol, Gemmi */
import "./presenter.css";
import { ensureGemmi, parseMtz, parseCcp4, renderDensityMap, removeDensityMap } from "../viewer/edmap.js";

let viewer = null;
let isSpinning = false;
let currentMapData = null;
let entries = []; // { id, name, pdbData, model, visible, settings }
let selectedEntryId = null;
let nextId = 1;

const WATER = ["HOH","WAT"];
const IONS = ["NA","CL","MG","ZN","CA","FE","MN","K","SO4","PO4"];
const SKIP = new Set([...WATER,...IONS]);
const CHARGES = { ARG:1, LYS:1, HIS:0.5, ASP:-1, GLU:-1 };

function init() {
  viewer = $3Dmol.createViewer("viewer", { backgroundColor:"white", antialias:true });
  bindGlobalEvents();
  setupDragDrop();
  checkForPush();
  applySlideViewState();

  // Re-sync camera when window is focused (e.g. user clicks "Open Presenter Window")
  window.addEventListener("focus", applySlideViewState);
}

function applySlideViewState() {
  try {
    const vs = localStorage.getItem("proteinviewer_viewState");
    if (vs && viewer) viewer.setView(JSON.parse(vs));
  } catch (e) { /**/ }
}

// Listen for view state relayed from the taskpane via postMessage
window.addEventListener("message", (event) => {
  if (event.data && event.data.type === "proteinviewer_viewState" && event.data.viewState) {
    try {
      if (viewer) viewer.setView(JSON.parse(event.data.viewState));
    } catch (e) { /**/ }
  }
});

let lastPushHash = "";
function checkForPush() {
  // Try multi-entry payload first
  const multiJson = localStorage.getItem("proteinviewer_multiEntries");
  const singlePdb = localStorage.getItem("proteinviewer_pdbData");
  const hash = (multiJson || "") + "_" + (singlePdb ? singlePdb.length : "") + "_" + (localStorage.getItem("proteinviewer_styleConfig") || "");
  if (hash === lastPushHash) return;
  lastPushHash = hash;

  if (multiJson) {
    try {
      const payload = JSON.parse(multiJson);
      if (payload.entries && payload.entries.length) {
        // Remove old PowerPoint-pushed entries
        entries = entries.filter((e) => !e._fromPush);
        for (const pe of payload.entries) {
          const entry = { id: nextId++, name: pe.name || "PowerPoint", pdbData: pe.pdbData, visible: true, settings: defaultSettings(), ligands: detectLigands(pe.pdbData), _fromPush: true };
          // Apply saved style config to entry settings
          const c = pe.styleConfig || {};
          const s = entry.settings;
          if (c.style) s.style = c.style;
          if (c.colorScheme) s.colorScheme = c.colorScheme;
          if (c.proteinOpacity !== undefined) s.opacity = c.proteinOpacity;
          if (c.showSurface !== undefined) s.surface = c.showSurface;
          if (c.surfaceType) s.surfType = c.surfaceType;
          if (c.surfaceColor) s.surfColor = c.surfaceColor;
          if (c.surfaceOpacity !== undefined) s.surfOpacity = c.surfaceOpacity;
          if (c.selectedLigand) s.ligand = c.selectedLigand;
          if (c.ligandStyle) s.ligandStyle = c.ligandStyle;
          if (c.zoomToLigand !== undefined) s.zoomLigand = c.zoomToLigand;
          if (c.showBindingSite !== undefined) s.bindingSite = c.showBindingSite;
          if (c.bindingDistance !== undefined) s.bindingDist = c.bindingDistance;
          if (c.showBindingLabels !== undefined) s.bindingLabels = c.showBindingLabels;
          if (c.showHbonds !== undefined) s.hbonds = c.showHbonds;
          if (c.showSaltBridges !== undefined) s.saltBridges = c.showSaltBridges;
          if (c.showPiStacking !== undefined) s.piStacking = c.showPiStacking;
          if (c.showPiCation !== undefined) s.piCation = c.showPiCation;
          if (c.showBindingSurface !== undefined) s.bindingSurface = c.showBindingSurface;
          if (c.bindingSurfaceColor) s.bsSurfColor = c.bindingSurfaceColor;
          if (c.bindingSurfaceOpacity !== undefined) s.bsSurfOpacity = c.bindingSurfaceOpacity;
          entries.push(entry);
        }
        if (payload.backgroundColor) document.getElementById("bg-select").value = payload.backgroundColor;
        selectedEntryId = entries.length ? entries[0].id : null;
        renderAll();
        renderEntryList();
        if (selectedEntryId) selectEntry(selectedEntryId);
        return;
      }
    } catch (e) { /**/ }
  }

  // Fallback: single-entry from taskpane
  if (!singlePdb) return;
  const existing = entries.find((e) => e.name === "PowerPoint");
  if (existing) {
    existing.pdbData = singlePdb;
    existing.ligands = detectLigands(singlePdb);
    renderAll();
    renderEntryList();
    if (selectedEntryId === existing.id) selectEntry(existing.id);
  } else {
    addEntry("PowerPoint", singlePdb);
  }
}

function setStatus(msg, type) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status-text" + (type ? " " + type : "");
}

function mapCS(s) {
  return { spectrum:"spectrum", chain:"chain", ss:"ssPyMol", residue:"amino", element:"default" }[s] || "default";
}

function defaultSettings() {
  return {
    style:"cartoon", colorScheme:"spectrum", opacity:1.0,
    ligand:null, ligandStyle:"ball-and-stick", zoomLigand:false,
    bindingSite:false, bindingDist:5, bindingLabels:false,
    hbonds:false, saltBridges:false, piStacking:false, piCation:false,
    bindingSurface:false, bsSurfColor:"esp", bsSurfOpacity:0.5,
    surface:false, surfType:"SAS", surfColor:"white", surfOpacity:0.6,
  };
}

// --- Entry management ---
function addEntry(name, pdbData) {
  const entry = { id: nextId++, name, pdbData, visible: true, settings: defaultSettings(), ligands: detectLigands(pdbData) };
  entries.push(entry);
  selectedEntryId = entry.id;
  renderAll();
  renderEntryList();
  selectEntry(entry.id);
  setStatus(`Added ${name}.`, "success");
}

function removeEntry(id) {
  entries = entries.filter((e) => e.id !== id);
  if (selectedEntryId === id) selectedEntryId = entries.length ? entries[0].id : null;
  renderAll();
  renderEntryList();
  if (selectedEntryId) selectEntry(selectedEntryId);
  else document.getElementById("entry-settings").style.display = "none";
}

function toggleVisibility(id) {
  const e = entries.find((e) => e.id === id);
  if (e) { e.visible = !e.visible; renderAll(); renderEntryList(); }
}

function selectEntry(id) {
  selectedEntryId = id;
  renderEntryList();
  const entry = entries.find((e) => e.id === id);
  if (!entry) { document.getElementById("entry-settings").style.display = "none"; return; }
  document.getElementById("entry-settings").style.display = "";
  document.getElementById("entry-settings-title").textContent = entry.name;
  loadSettingsToUI(entry);
}

function getSelectedEntry() { return entries.find((e) => e.id === selectedEntryId); }

function detectLigands(pdbText) {
  const m = {};
  for (const line of pdbText.split("\n")) {
    if (!line.startsWith("HETATM")) continue;
    const resn = line.substring(17,20).trim();
    if (SKIP.has(resn)) continue;
    const chain = line.substring(21,22).trim(), resi = line.substring(22,26).trim();
    const k = `${resn}_${chain}_${resi}`;
    if (!m[k]) m[k] = { resn, chain, resi, count:0 };
    m[k].count++;
  }
  return Object.values(m);
}

// --- Entry list UI ---
function renderEntryList() {
  const list = document.getElementById("entry-list");
  list.innerHTML = "";
  for (const e of entries) {
    const div = document.createElement("div");
    div.className = "entry-item" + (e.id === selectedEntryId ? " selected" : "");
    div.innerHTML = `
      <span class="entry-vis ${e.visible ? "visible" : "hidden"}" data-id="${e.id}">${e.visible ? "👁" : "—"}</span>
      <span class="entry-name" data-id="${e.id}">${e.name}</span>
      <span class="entry-remove" data-id="${e.id}">✕</span>`;
    list.appendChild(div);
  }
  list.querySelectorAll(".entry-name").forEach((el) => el.addEventListener("click", () => selectEntry(+el.dataset.id)));
  list.querySelectorAll(".entry-vis").forEach((el) => el.addEventListener("click", (ev) => { ev.stopPropagation(); toggleVisibility(+el.dataset.id); }));
  list.querySelectorAll(".entry-remove").forEach((el) => el.addEventListener("click", (ev) => { ev.stopPropagation(); removeEntry(+el.dataset.id); }));

  // Update align section
  const alignSec = document.getElementById("align-section");
  if (entries.length >= 2) {
    alignSec.style.display = "";
    const refSel = document.getElementById("align-ref");
    const curVal = refSel.value;
    refSel.innerHTML = "";
    for (const e of entries) {
      const opt = document.createElement("option");
      opt.value = e.id;
      opt.textContent = e.name;
      refSel.appendChild(opt);
    }
    if (curVal && entries.find((e) => e.id === +curVal)) refSel.value = curVal;
  } else {
    alignSec.style.display = "none";
  }
}

// --- Settings UI <-> entry sync ---
function loadSettingsToUI(entry) {
  const s = entry.settings;
  document.getElementById("style-select").value = s.style;
  document.getElementById("color-scheme").value = s.colorScheme;
  document.getElementById("protein-opacity").value = Math.round(s.opacity * 100);
  document.getElementById("protein-opacity-val").textContent = Math.round(s.opacity * 100) + "%";
  document.getElementById("ligand-style").value = s.ligandStyle;
  document.getElementById("chk-zoom-ligand").checked = s.zoomLigand;
  document.getElementById("chk-binding-site").checked = s.bindingSite;
  document.getElementById("binding-dist").value = s.bindingDist;
  document.getElementById("binding-dist-val").textContent = s.bindingDist.toFixed(1) + " Å";
  document.getElementById("binding-dist-row").style.display = s.bindingSite ? "" : "none";
  document.getElementById("chk-binding-labels").checked = s.bindingLabels;
  document.getElementById("chk-hbonds").checked = s.hbonds;
  document.getElementById("chk-salt-bridges").checked = s.saltBridges;
  document.getElementById("chk-pi-stacking").checked = s.piStacking;
  document.getElementById("chk-pi-cation").checked = s.piCation;
  document.getElementById("chk-binding-surface").checked = s.bindingSurface;
  document.getElementById("binding-surface-color").value = s.bsSurfColor;
  document.getElementById("binding-surface-opacity").value = Math.round(s.bsSurfOpacity * 100);
  document.getElementById("binding-surface-opacity-val").textContent = Math.round(s.bsSurfOpacity * 100) + "%";
  document.getElementById("binding-surface-options").style.display = s.bindingSurface ? "" : "none";
  document.getElementById("surface-toggle").checked = s.surface;
  document.getElementById("surface-type").value = s.surfType;
  document.getElementById("surface-color").value = s.surfColor;
  document.getElementById("surface-opacity").value = Math.round(s.surfOpacity * 100);
  document.getElementById("surface-opacity-val").textContent = Math.round(s.surfOpacity * 100) + "%";
  document.getElementById("surface-options").style.display = s.surface ? "" : "none";

  // Populate ligand dropdown for this entry
  const sel = document.getElementById("ligand-select");
  sel.innerHTML = '<option value="">None</option>';
  if (entry.ligands.length) {
    document.getElementById("ligand-section").style.display = "";
    for (const lig of entry.ligands) {
      const opt = document.createElement("option");
      opt.value = JSON.stringify({ resn:lig.resn, chain:lig.chain, resi:lig.resi });
      opt.textContent = `${lig.resn} (${lig.chain}:${lig.resi})`;
      sel.appendChild(opt);
    }
    if (s.ligand) sel.value = JSON.stringify(s.ligand);
  } else {
    document.getElementById("ligand-section").style.display = "none";
  }
}

function saveUIToSettings() {
  const entry = getSelectedEntry();
  if (!entry) return;
  const s = entry.settings;
  s.style = document.getElementById("style-select").value;
  s.colorScheme = document.getElementById("color-scheme").value;
  s.opacity = parseInt(document.getElementById("protein-opacity").value) / 100;
  const ligVal = document.getElementById("ligand-select").value;
  s.ligand = ligVal ? JSON.parse(ligVal) : null;
  s.ligandStyle = document.getElementById("ligand-style").value;
  s.zoomLigand = document.getElementById("chk-zoom-ligand").checked;
  s.bindingSite = document.getElementById("chk-binding-site").checked;
  s.bindingDist = parseFloat(document.getElementById("binding-dist").value);
  s.bindingLabels = document.getElementById("chk-binding-labels").checked;
  s.hbonds = document.getElementById("chk-hbonds").checked;
  s.saltBridges = document.getElementById("chk-salt-bridges").checked;
  s.piStacking = document.getElementById("chk-pi-stacking").checked;
  s.piCation = document.getElementById("chk-pi-cation").checked;
  s.bindingSurface = document.getElementById("chk-binding-surface").checked;
  s.bsSurfColor = document.getElementById("binding-surface-color").value;
  s.bsSurfOpacity = parseInt(document.getElementById("binding-surface-opacity").value) / 100;
  s.surface = document.getElementById("surface-toggle").checked;
  s.surfType = document.getElementById("surface-type").value;
  s.surfColor = document.getElementById("surface-color").value;
  s.surfOpacity = parseInt(document.getElementById("surface-opacity").value) / 100;
}

// --- Rendering ---
function renderAll() {
  if (!viewer) return;
  viewer.removeAllModels(); viewer.removeAllSurfaces(); viewer.removeAllShapes(); viewer.removeAllLabels();

  for (let idx = 0; idx < entries.length; idx++) {
    const entry = entries[idx];
    const model = viewer.addModel(entry.pdbData, "pdb");
    entry.model = model;
    entry.modelIdx = idx;
    // Assign charges directly on atom for surface coloring
    for (const a of model.selectedAtoms({})) { a.charge = CHARGES[a.resn] || 0; }

    if (!entry.visible) {
      viewer.setStyle({ model: model }, {});
      continue;
    }

    const s = entry.settings;
    const cs = mapCS(s.colorScheme);
    const so = {}; so[s.style] = { colorscheme: cs, opacity: s.opacity };
    viewer.setStyle({ model: model, not: { hetflag: true } }, so);

    // Surface
    if (s.surface) {
      let sType, sColor;
      if (s.surfType === "ESP") {
        sType = $3Dmol.SurfaceType.SAS;
        sColor = { prop: "charge", gradient: new $3Dmol.Gradient.RWB(-1, 1) };
      } else {
        sType = getSurfType(s.surfType);
        sColor = getSurfColorSpec(s.surfColor, s.colorScheme);
      }
      viewer.addSurface(sType, { opacity: s.surfOpacity, ...sColor },
        { not: { hetflag: true } });
    }

    // Ligand
    if (s.ligand) {
      const resi = parseInt(s.ligand.resi, 10);
      const ligSel = { model: model, resn: s.ligand.resn, chain: s.ligand.chain, resi };
      viewer.addStyle(ligSel, buildLigStyle(s.ligandStyle));

      if (s.bindingSite) {
        viewer.render();
        renderBindingSite(model, ligSel, s);
      }

      if (s.zoomLigand) viewer.zoomTo(ligSel);
    } else {
      viewer.addStyle({ model: model, hetflag: true, not: { resn: [...WATER,...IONS] } },
        { stick: { colorscheme:"default", radius:0.15 }, sphere: { colorscheme:"default", radius:0.3 } });
    }
  }

  viewer.setBackgroundColor(document.getElementById("bg-select").value);
  if (!entries.some((e) => e.settings.zoomLigand && e.settings.ligand && e.visible)) viewer.zoomTo();
  viewer.render();
}

function getSurfType(t) { return { VDW:$3Dmol.SurfaceType.VDW, SAS:$3Dmol.SurfaceType.SAS, SES:$3Dmol.SurfaceType.SES }[t] || $3Dmol.SurfaceType.SAS; }
function getSurfColorSpec(mode, cs) {
  switch (mode) {
    case "esp": return { prop:"charge", gradient: new $3Dmol.Gradient.RWB(-1,1) };
    case "protein": return { colorscheme: mapCS(cs) };
    case "hydrophobicity": return { prop:"charge", gradient: new $3Dmol.Gradient.Sinebow(0,1) };
    case "element": return { colorscheme:"default" };
    default: return { color:"white" };
  }
}
function buildLigStyle(s) {
  switch (s) {
    case "ball-and-stick": return { stick:{colorscheme:"greenCarbon",radius:0.15}, sphere:{colorscheme:"greenCarbon",radius:0.35} };
    case "stick": return { stick:{colorscheme:"greenCarbon",radius:0.2} };
    case "sphere": return { sphere:{colorscheme:"greenCarbon"} };
    default: return { stick:{colorscheme:"greenCarbon",radius:0.15} };
  }
}

function renderBindingSite(model, ligSel, s) {
  const ligAtoms = model.selectedAtoms(ligSel);
  const protAtoms = model.selectedAtoms({ not:{hetflag:true} });
  if (!ligAtoms?.length || !protAtoms?.length) return;

  const nearby = new Map();
  for (const la of ligAtoms) for (const pa of protAtoms)
    if (ad(la,pa) <= s.bindingDist) { const k=`${pa.chain}:${pa.resi}`; if (!nearby.has(k)) nearby.set(k,{chain:pa.chain,resi:pa.resi,resn:pa.resn}); }
  if (!nearby.size) return;

  const cr = {}; const allR = [];
  for (const [,r] of nearby) { if (!cr[r.chain]) cr[r.chain]=[]; cr[r.chain].push(r.resi); allR.push(r.resi); }

  for (const [ch,res] of Object.entries(cr))
    viewer.addStyle({model:model,chain:ch,resi:res,not:{hetflag:true}},{stick:{colorscheme:"default",radius:0.12}});

  if (s.bindingLabels) for (const [ch,res] of Object.entries(cr))
    viewer.addResLabels({model:model,chain:ch,resi:res,atom:"CA"},{font:"Arial",fontSize:10,showBackground:true,backgroundColor:0x333333,backgroundOpacity:0.8,fontColor:"white"});

  const np = protAtoms.filter((a)=>nearby.has(`${a.chain}:${a.resi}`));

  if (s.hbonds) { const hb=new Set(["N","O","S"]);
    for (const la of ligAtoms) { if (!hb.has(la.elem)) continue; for (const pa of np) { if (!hb.has(pa.elem)) continue; const d=ad(la,pa); if (d>=2&&d<=3.5) dsh(la,pa,"yellow"); } } }

  if (s.saltBridges) { const pos=new Set(["NZ","NH1","NH2","NE"]),neg=new Set(["OD1","OD2","OE1","OE2"]);
    for (const la of ligAtoms.filter(a=>a.elem==="N"||a.elem==="O")) for (const pa of np) {
      if (!(pos.has(pa.atom)&&["ARG","LYS"].includes(pa.resn))&&!(neg.has(pa.atom)&&["ASP","GLU"].includes(pa.resn))) continue;
      const d=ad(la,pa); if (d>=1.5&&d<=4) dsh(la,pa,"magenta"); } }

  if (s.piStacking) { const lr=findAR(ligAtoms),pr=findPAR(np);
    for (const a of lr) for (const b of pr) if (ad(a.centroid,b.centroid)<=5.5) dsh(a.centroid,b.centroid,"cyan"); }

  if (s.piCation) { const lr=findAR(ligAtoms),pr=findPAR(np);
    const pc=np.filter(a=>(a.atom==="NZ"&&a.resn==="LYS")||(["NH1","NH2"].includes(a.atom)&&a.resn==="ARG"));
    const lc=ligAtoms.filter(a=>a.elem==="N");
    for (const r of lr) for (const c of pc) if (ad(r.centroid,c)<=6) dsh(r.centroid,c,"orange");
    for (const r of pr) for (const c of lc) if (ad(r.centroid,c)<=6) dsh(r.centroid,c,"orange"); }

  if (s.bindingSurface) {
    const cs=getSurfColorSpec(s.bsSurfColor,s.colorScheme);
    viewer.addSurface($3Dmol.SurfaceType.SAS,{opacity:s.bsSurfOpacity,...cs},{resi:allR,not:{hetflag:true}},{not:{hetflag:true}}); }
}

function ad(a,b){const dx=a.x-b.x,dy=a.y-b.y,dz=a.z-b.z;return Math.sqrt(dx*dx+dy*dy+dz*dz);}
function dsh(a,b,color){viewer.addCylinder({start:{x:a.x,y:a.y,z:a.z},end:{x:b.x,y:b.y,z:b.z},radius:0.04,color,fromCap:true,toCap:true,dashed:true,dashLength:0.15,gapLength:0.1});}

function findAR(atoms){const ar=atoms.filter(a=>["C","N"].includes(a.elem));const adj=new Map();for(const a of ar)adj.set(a,[]);
  for(let i=0;i<ar.length;i++)for(let j=i+1;j<ar.length;j++)if(ad(ar[i],ar[j])<=1.7){adj.get(ar[i]).push(ar[j]);adj.get(ar[j]).push(ar[i]);}
  const f=new Set(),r=[];for(const s of ar)dfs(s,s,[s],adj,f,r,6);return r;}
function dfs(st,cu,pa,adj,f,r,mx){if(pa.length>mx)return;for(const n of adj.get(cu)||[]){
  if(n===st&&pa.length>=5){const k=pa.map(a=>`${a.x.toFixed(1)},${a.y.toFixed(1)}`).sort().join("|");
    if(!f.has(k)){f.add(k);r.push({centroid:cof(pa),atoms:[...pa]});}return;}
  if(pa.includes(n))continue;pa.push(n);dfs(st,n,pa,adj,f,r,mx);pa.pop();}}
function findPAR(atoms){const ar={PHE:["CG","CD1","CD2","CE1","CE2","CZ"],TYR:["CG","CD1","CD2","CE1","CE2","CZ"],TRP:null,HIS:["CG","ND1","CD2","CE1","NE2"]};
  const g={};for(const a of atoms){if(!ar.hasOwnProperty(a.resn))continue;const k=`${a.chain}:${a.resi}:${a.resn}`;if(!g[k])g[k]={resn:a.resn,atoms:[]};g[k].atoms.push(a);}
  const r=[];for(const[,v]of Object.entries(g)){if(v.resn==="TRP"){for(const rn of[["CG","CD1","NE1","CE2","CD2"],["CD2","CE2","CE3","CZ3","CH2","CZ2"]]){
    const ra=rn.map(n=>v.atoms.find(a=>a.atom===n)).filter(Boolean);if(ra.length>=4)r.push({centroid:cof(ra),atoms:ra});}}
  else{const rn=ar[v.resn];if(!rn)continue;const ra=rn.map(n=>v.atoms.find(a=>a.atom===n)).filter(Boolean);if(ra.length>=4)r.push({centroid:cof(ra),atoms:ra});}}return r;}
function cof(a){const n=a.length;return{x:a.reduce((s,v)=>s+v.x,0)/n,y:a.reduce((s,v)=>s+v.y,0)/n,z:a.reduce((s,v)=>s+v.z,0)/n};}

// --- Structural Alignment (Kabsch) ---
function handleAlign() {
  const refId = +document.getElementById("align-ref").value;
  const refEntry = entries.find((e) => e.id === refId);
  if (!refEntry) return;
  const statusEl = document.getElementById("align-status");

  const refCA = extractCA(refEntry.pdbData);
  if (refCA.length === 0) { statusEl.textContent = "No Cα atoms in reference."; statusEl.className = "status-text error"; return; }

  let aligned = 0;
  for (const entry of entries) {
    if (entry.id === refId) continue;
    const mobCA = extractCA(entry.pdbData);
    if (mobCA.length === 0) continue;

    // Match by residue index (use shorter length)
    const n = Math.min(refCA.length, mobCA.length);
    const refPts = refCA.slice(0, n).map((a) => [a.x, a.y, a.z]);
    const mobPts = mobCA.slice(0, n).map((a) => [a.x, a.y, a.z]);

    const { rotation, translation, rmsd } = kabsch(mobPts, refPts);
    entry.pdbData = applyTransform(entry.pdbData, rotation, translation);
    entry.rmsd = rmsd;
    aligned++;
  }

  renderAll();
  statusEl.textContent = `Aligned ${aligned} structure(s) to ${refEntry.name}.`;
  statusEl.className = "status-text success";
}

function extractCA(pdbText) {
  const cas = [];
  for (const line of pdbText.split("\n")) {
    if (!line.startsWith("ATOM")) continue;
    const atomName = line.substring(12, 16).trim();
    if (atomName !== "CA") continue;
    cas.push({
      x: parseFloat(line.substring(30, 38)),
      y: parseFloat(line.substring(38, 46)),
      z: parseFloat(line.substring(46, 54)),
    });
  }
  return cas;
}

function kabsch(mobile, reference) {
  // mobile and reference are arrays of [x,y,z]
  const n = mobile.length;

  // Compute centroids
  const cM = [0, 0, 0], cR = [0, 0, 0];
  for (let i = 0; i < n; i++) for (let j = 0; j < 3; j++) { cM[j] += mobile[i][j]; cR[j] += reference[i][j]; }
  for (let j = 0; j < 3; j++) { cM[j] /= n; cR[j] /= n; }

  // Center the points
  const P = mobile.map((p) => [p[0] - cM[0], p[1] - cM[1], p[2] - cM[2]]);
  const Q = reference.map((p) => [p[0] - cR[0], p[1] - cR[1], p[2] - cR[2]]);

  // Compute cross-covariance matrix H = P^T * Q (3x3)
  const H = [[0,0,0],[0,0,0],[0,0,0]];
  for (let i = 0; i < n; i++)
    for (let j = 0; j < 3; j++)
      for (let k = 0; k < 3; k++)
        H[j][k] += P[i][j] * Q[i][k];

  // SVD of H using Jacobi rotations (simple 3x3 SVD)
  const { U, V } = svd3x3(H);

  // Ensure proper rotation (det > 0)
  const detUV = det3(U) * det3(V);
  const d = detUV < 0 ? -1 : 1;
  const S = [[1,0,0],[0,1,0],[0,0,d]];

  // R = V * S * U^T
  const SUT = mul3(S, transpose3(U));
  const R = mul3(V, SUT);

  // Translation: t = cR - R * cM
  const RcM = [
    R[0][0]*cM[0] + R[0][1]*cM[1] + R[0][2]*cM[2],
    R[1][0]*cM[0] + R[1][1]*cM[1] + R[1][2]*cM[2],
    R[2][0]*cM[0] + R[2][1]*cM[1] + R[2][2]*cM[2],
  ];
  const t = [cR[0] - RcM[0], cR[1] - RcM[1], cR[2] - RcM[2]];

  // Compute RMSD
  let rmsd = 0;
  for (let i = 0; i < n; i++) {
    const mx = R[0][0]*mobile[i][0]+R[0][1]*mobile[i][1]+R[0][2]*mobile[i][2]+t[0];
    const my = R[1][0]*mobile[i][0]+R[1][1]*mobile[i][1]+R[1][2]*mobile[i][2]+t[1];
    const mz = R[2][0]*mobile[i][0]+R[2][1]*mobile[i][1]+R[2][2]*mobile[i][2]+t[2];
    rmsd += (mx-reference[i][0])**2 + (my-reference[i][1])**2 + (mz-reference[i][2])**2;
  }
  rmsd = Math.sqrt(rmsd / n);

  return { rotation: R, translation: t, rmsd };
}

function applyTransform(pdbText, R, t) {
  return pdbText.split("\n").map((line) => {
    if (!line.startsWith("ATOM") && !line.startsWith("HETATM")) return line;
    const x = parseFloat(line.substring(30, 38));
    const y = parseFloat(line.substring(38, 46));
    const z = parseFloat(line.substring(46, 54));
    const nx = R[0][0]*x + R[0][1]*y + R[0][2]*z + t[0];
    const ny = R[1][0]*x + R[1][1]*y + R[1][2]*z + t[1];
    const nz = R[2][0]*x + R[2][1]*y + R[2][2]*z + t[2];
    return line.substring(0, 30) + nx.toFixed(3).padStart(8) + ny.toFixed(3).padStart(8) + nz.toFixed(3).padStart(8) + line.substring(54);
  }).join("\n");
}

// --- 3x3 matrix utilities ---
function transpose3(M) { return [[M[0][0],M[1][0],M[2][0]],[M[0][1],M[1][1],M[2][1]],[M[0][2],M[1][2],M[2][2]]]; }
function mul3(A, B) {
  const C = [[0,0,0],[0,0,0],[0,0,0]];
  for (let i=0;i<3;i++) for (let j=0;j<3;j++) for (let k=0;k<3;k++) C[i][j]+=A[i][k]*B[k][j];
  return C;
}
function det3(M) {
  return M[0][0]*(M[1][1]*M[2][2]-M[1][2]*M[2][1]) - M[0][1]*(M[1][0]*M[2][2]-M[1][2]*M[2][0]) + M[0][2]*(M[1][0]*M[2][1]-M[1][1]*M[2][0]);
}

// Simple 3x3 SVD via Jacobi iteration (sufficient for Kabsch)
function svd3x3(A) {
  // Compute A^T*A
  const AtA = mul3(transpose3(A), A);
  // Eigendecomposition of AtA via Jacobi
  const { eigVecs, eigVals } = jacobi3(AtA);
  // V = eigenvectors, Sigma = sqrt(eigenvalues)
  const V = eigVecs;
  const sigma = eigVals.map((v) => Math.sqrt(Math.max(v, 0)));
  // U = A * V * Sigma^-1
  const AV = mul3(A, V);
  const U = [[0,0,0],[0,0,0],[0,0,0]];
  for (let i=0;i<3;i++) for (let j=0;j<3;j++) U[i][j] = sigma[j] > 1e-10 ? AV[i][j]/sigma[j] : 0;
  return { U, V, sigma };
}

function jacobi3(A) {
  const M = A.map((r) => [...r]);
  const V = [[1,0,0],[0,1,0],[0,0,1]];
  for (let iter = 0; iter < 50; iter++) {
    // Find largest off-diagonal
    let p=0, q=1, maxVal=Math.abs(M[0][1]);
    if (Math.abs(M[0][2])>maxVal) { p=0;q=2;maxVal=Math.abs(M[0][2]); }
    if (Math.abs(M[1][2])>maxVal) { p=1;q=2; }
    if (Math.abs(M[p][q]) < 1e-12) break;
    const theta = 0.5*Math.atan2(2*M[p][q], M[p][p]-M[q][q]);
    const c=Math.cos(theta), s=Math.sin(theta);
    // Rotate M
    const G = [[1,0,0],[0,1,0],[0,0,1]];
    G[p][p]=c; G[q][q]=c; G[p][q]=-s; G[q][p]=s;
    const Gt = transpose3(G);
    const tmp = mul3(mul3(Gt, M), G);
    for (let i=0;i<3;i++) for (let j=0;j<3;j++) M[i][j]=tmp[i][j];
    // Accumulate V
    const Vnew = mul3(V, G);
    for (let i=0;i<3;i++) for (let j=0;j<3;j++) V[i][j]=Vnew[i][j];
  }
  return { eigVecs: V, eigVals: [M[0][0], M[1][1], M[2][2]] };
}

// --- Push to Slide ---
function entryToStyleConfig(entry) {
  const s = entry.settings;
  return {
    style: s.style,
    colorScheme: s.colorScheme,
    proteinOpacity: s.opacity,
    showSurface: s.surface,
    surfaceType: s.surfType,
    surfaceColor: s.surfColor,
    surfaceOpacity: s.surfOpacity,
    ligandStyle: s.ligandStyle,
    selectedLigand: s.ligand || null,
    zoomToLigand: s.zoomLigand,
    showBindingSite: s.bindingSite,
    bindingDistance: s.bindingDist,
    showBindingLabels: s.bindingLabels,
    showHbonds: s.hbonds,
    showSaltBridges: s.saltBridges,
    showPiStacking: s.piStacking,
    showPiCation: s.piCation,
    showBindingSurface: s.bindingSurface,
    bindingSurfaceColor: s.bsSurfColor,
    bindingSurfaceOpacity: s.bsSurfOpacity,
  };
}

function pushToSlide() {
  const statusEl = document.getElementById("push-status");
  const visibleEntries = entries.filter((e) => e.visible);
  if (!visibleEntries.length) { statusEl.textContent = "No visible entries."; statusEl.className = "status-text error"; return; }

  const bg = document.getElementById("bg-select").value;
  const viewState = viewer ? viewer.getView() : null;
  const multiPayload = visibleEntries.map((e) => ({
    name: e.name,
    pdbData: e.pdbData,
    styleConfig: entryToStyleConfig(e),
  }));

  // Write view state as the single source of truth
  if (viewState) localStorage.setItem("proteinviewer_viewState", JSON.stringify(viewState));

  // Multi-entry key (viewState included for content add-in to use on first render)
  localStorage.setItem("proteinviewer_multiEntries", JSON.stringify({ backgroundColor: bg, viewState, entries: multiPayload }));

  // Also set single-entry keys for backward compat with taskpane
  const first = visibleEntries[0];
  localStorage.setItem("proteinviewer_pdbData", first.pdbData);
  localStorage.setItem("proteinviewer_styleConfig", JSON.stringify({ ...entryToStyleConfig(first), backgroundColor: bg }));

  statusEl.textContent = `Pushed ${visibleEntries.length} entry(s) to slide.`;
  statusEl.className = "status-text success";
  window.blur();
}

// --- Events ---
function bindGlobalEvents() {
  document.getElementById("btn-fetch").addEventListener("click", handleFetch);
  document.getElementById("pdb-input").addEventListener("keydown", (e) => { if (e.key === "Enter") handleFetch(); });
  document.getElementById("pdb-file").addEventListener("change", handleFiles);

  // Per-entry settings — save and re-render on change
  const ids = ["style-select","color-scheme","protein-opacity","ligand-select","ligand-style",
    "chk-zoom-ligand","chk-binding-site","chk-binding-labels","chk-hbonds","chk-salt-bridges",
    "chk-pi-stacking","chk-pi-cation","chk-binding-surface","binding-surface-color",
    "binding-surface-opacity","surface-toggle","surface-type","surface-color","surface-opacity"];
  ids.forEach((id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener("change", () => { saveUIToSettings(); renderAll(); updateConditionalUI(); });
  });

  // Slider labels
  const sliders = [
    ["protein-opacity","protein-opacity-val","%"],
    ["surface-opacity","surface-opacity-val","%"],
    ["binding-surface-opacity","binding-surface-opacity-val","%"],
  ];
  sliders.forEach(([id,labelId,suffix]) => {
    document.getElementById(id).addEventListener("input", (e) => {
      document.getElementById(labelId).textContent = e.target.value + suffix;
    });
  });
  document.getElementById("binding-dist").addEventListener("input", (e) => {
    document.getElementById("binding-dist-val").textContent = parseFloat(e.target.value).toFixed(1) + " Å";
  });
  document.getElementById("binding-dist").addEventListener("change", () => { saveUIToSettings(); renderAll(); });

  // Global controls
  document.getElementById("bg-select").addEventListener("change", () => { if (viewer) { viewer.setBackgroundColor(document.getElementById("bg-select").value); viewer.render(); } });
  document.getElementById("spin-toggle").addEventListener("change", (e) => { isSpinning = e.target.checked; if (viewer) viewer.spin(isSpinning ? "y" : false); });
  document.getElementById("panel-toggle").addEventListener("click", togglePanel);
  document.getElementById("btn-align").addEventListener("click", handleAlign);
  document.getElementById("btn-push-slide").addEventListener("click", pushToSlide);
  window.addEventListener("keydown", handleKey, true);
  // Also catch keys on the viewer canvas directly
  document.getElementById("viewer").addEventListener("keydown", handleKey, true);

  // Map controls
  document.getElementById("map-file").addEventListener("change", handleLoadMapPresenter);
  document.getElementById("btn-remove-map").addEventListener("click", handleRemoveMapPresenter);
  document.getElementById("map-2fofc-sigma").addEventListener("input", (e) => {
    document.getElementById("map-2fofc-sigma-val").textContent = parseFloat(e.target.value).toFixed(1) + "σ";
  });
  document.getElementById("map-2fofc-sigma").addEventListener("change", reRenderMapPresenter);
  document.getElementById("chk-fofc-map").addEventListener("change", () => {
    document.getElementById("fofc-controls").style.display = document.getElementById("chk-fofc-map").checked ? "" : "none";
    reRenderMapPresenter();
  });
  document.getElementById("map-fofc-sigma").addEventListener("input", (e) => {
    document.getElementById("map-fofc-sigma-val").textContent = parseFloat(e.target.value).toFixed(1) + "σ";
  });
  document.getElementById("map-fofc-sigma").addEventListener("change", reRenderMapPresenter);
  document.getElementById("map-radius").addEventListener("input", (e) => {
    document.getElementById("map-radius-val").textContent = e.target.value + " Å";
  });
  document.getElementById("map-radius").addEventListener("change", reRenderMapPresenter);
}

function updateConditionalUI() {
  const entry = getSelectedEntry();
  if (!entry) return;
  const s = entry.settings;
  document.getElementById("binding-dist-row").style.display = s.bindingSite ? "" : "none";
  document.getElementById("binding-surface-options").style.display = s.bindingSurface ? "" : "none";
  document.getElementById("surface-options").style.display = s.surface ? "" : "none";
}

async function handleFetch() {
  const id = document.getElementById("pdb-input").value.trim().toUpperCase();
  if (!id || id.length !== 4) { setStatus("Enter 4-char PDB ID.", "error"); return; }
  setStatus("Fetching " + id + "...");
  try {
    const r = await fetch(`https://files.rcsb.org/download/${id}.pdb`);
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    addEntry(id, await r.text());
    document.getElementById("pdb-input").value = "";
  } catch (e) { setStatus(e.message, "error"); }
}

function handleFiles() {
  const files = document.getElementById("pdb-file").files;
  for (const f of files) {
    const reader = new FileReader();
    reader.onload = (e) => addEntry(f.name.replace(/\.[^.]+$/, ""), e.target.result);
    reader.readAsText(f);
  }
}

function togglePanel() {
  const p = document.getElementById("panel");
  p.classList.toggle("collapsed");
  document.getElementById("panel-toggle").textContent = p.classList.contains("collapsed") ? "▶" : "◀";
  setTimeout(() => { if (viewer) viewer.resize(); }, 250);
}

function handleKey(e) {
  const tag = e.target.tagName;
  if (tag === "INPUT" || tag === "SELECT" || tag === "TEXTAREA") return;
  switch (e.key) {
    case "s": case "S":
      e.preventDefault();
      e.stopImmediatePropagation();
      isSpinning = !isSpinning;
      if (viewer) viewer.spin(isSpinning ? "y" : false);
      document.getElementById("spin-toggle").checked = isSpinning;
      break;
    case "0":
      break;
    case "p": case "P":
      e.preventDefault();
      e.stopImmediatePropagation();
      togglePanel();
      break;
    case "Escape":
      if (document.fullscreenElement) document.exitFullscreen();
      break;
  }
}

function setupDragDrop() {
  const el = document.getElementById("viewer");
  el.addEventListener("dragover", (e) => { e.preventDefault(); el.classList.add("dragover"); });
  el.addEventListener("dragleave", () => el.classList.remove("dragover"));
  el.addEventListener("drop", (e) => {
    e.preventDefault(); el.classList.remove("dragover");
    for (const f of e.dataTransfer.files) {
      const reader = new FileReader();
      reader.onload = (ev) => addEntry(f.name.replace(/\.[^.]+$/, ""), ev.target.result);
      reader.readAsText(f);
    }
  });
}

if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
else init();

function handleLoadMapPresenter() {
  const fileInput = document.getElementById("map-file");
  const mapStatus = document.getElementById("map-status");
  if (!fileInput.files.length) return;
  if (!viewer) { mapStatus.textContent = "Load a structure first."; mapStatus.className = "status-text error"; return; }

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
          const ccp4Map = parseCcp4(gemmi, e.target.result);
          currentMapData = { map2fofc: ccp4Map, mapFofc: null };
        }
        reRenderMapPresenter();
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

function reRenderMapPresenter() {
  if (!viewer || !currentMapData) return;
  removeDensityMap(viewer);
  const sigma2fofc = parseFloat(document.getElementById("map-2fofc-sigma").value);
  const sigmaFofc = parseFloat(document.getElementById("map-fofc-sigma").value);
  const showFofc = document.getElementById("chk-fofc-map").checked;
  renderDensityMap(viewer, currentMapData, { sigma2fofc, sigmaFofc, showFofc });
}

function handleRemoveMapPresenter() {
  if (viewer) removeDensityMap(viewer);
  currentMapData = null;
  document.getElementById("map-controls").style.display = "none";
  document.getElementById("map-status").textContent = "";
  document.getElementById("map-file").value = "";
}
