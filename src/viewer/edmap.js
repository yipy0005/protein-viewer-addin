/* global Gemmi, $3Dmol */

/**
 * Electron density map support using gemmi WASM.
 * Uses gemmi's extract_isosurface for correct crystallographic coordinate handling.
 */

let gemmiModule = null;
let gemmiLoading = false;
let gemmiCallbacks = [];

export function ensureGemmi(callback) {
  if (gemmiModule) { callback(gemmiModule); return; }
  gemmiCallbacks.push(callback);
  if (gemmiLoading) return;
  gemmiLoading = true;
  if (typeof Gemmi === "function") {
    Gemmi().then((mod) => {
      gemmiModule = mod;
      gemmiCallbacks.forEach((cb) => cb(mod));
      gemmiCallbacks = [];
    }).catch((err) => {
      console.error("Failed to load gemmi WASM:", err);
      gemmiCallbacks = [];
    });
  } else {
    console.error("Gemmi not found. Make sure gemmi.js is loaded.");
    gemmiCallbacks = [];
  }
}

export function parseMtz(gemmi, arrayBuffer) {
  const mtz = gemmi.readMtz(arrayBuffer);
  const map2fofc = mtz.calculate_wasm_map(false);
  const mapFofc = mtz.calculate_wasm_map(true);
  return { map2fofc, mapFofc, mtz };
}

export function parseCcp4(gemmi, arrayBuffer) {
  const ccp4Map = gemmi.readCcp4Map(arrayBuffer, true);
  return { map2fofc: ccp4Map, mapFofc: null };
}

/**
 * Get the center of mass of the current model in the viewer.
 */
function getModelCenter(viewer) {
  try {
    const atoms = viewer.getModel().selectedAtoms({});
    if (!atoms || !atoms.length) return { x: 0, y: 0, z: 0 };
    let cx = 0, cy = 0, cz = 0;
    for (const a of atoms) { cx += a.x; cy += a.y; cz += a.z; }
    const n = atoms.length;
    return { x: cx / n, y: cy / n, z: cz / n };
  } catch (e) {
    return { x: 0, y: 0, z: 0 };
  }
}

/**
 * Extract isosurface from a gemmi wasm map and render as lines in 3Dmol.js.
 * Uses gemmi's extract_isosurface which handles crystallographic transforms correctly.
 */
function addIsoLines(viewer, wasmMap, sigma, radius, center, color) {
  const isolevel = wasmMap.mean + sigma * wasmMap.rms;
  const ok = wasmMap.extract_isosurface(radius, center.x, center.y, center.z, isolevel, "");
  if (!ok) return [];

  const verts = wasmMap.isosurface_vertices();
  const segs = wasmMap.isosurface_segments();
  if (!verts || !segs || verts.length === 0 || segs.length === 0) return [];

  // segs contains pairs of indices into verts (each vertex is 3 floats)
  const shapes = [];
  for (let i = 0; i < segs.length; i += 2) {
    const i0 = segs[i] * 3;
    const i1 = segs[i + 1] * 3;
    const shape = viewer.addLine({
      start: { x: verts[i0], y: verts[i0 + 1], z: verts[i0 + 2] },
      end: { x: verts[i1], y: verts[i1 + 1], z: verts[i1 + 2] },
      color: color,
    });
    shapes.push(shape);
  }
  return shapes;
}

let currentMapShapes = [];

export function renderDensityMap(viewer, mapData, opts) {
  const sigma2fofc = opts.sigma2fofc || 1.5;
  const sigmaFofc = opts.sigmaFofc || 3.0;
  const showFofc = opts.showFofc || false;
  const radius = opts.radius || 8;
  const center = getModelCenter(viewer);

  // 2Fo-Fc (blue)
  if (mapData.map2fofc) {
    const shapes = addIsoLines(viewer, mapData.map2fofc, sigma2fofc, radius, center, "blue");
    currentMapShapes.push(...shapes);
  }

  // Fo-Fc positive (green) and negative (red)
  if (showFofc && mapData.mapFofc) {
    const posShapes = addIsoLines(viewer, mapData.mapFofc, sigmaFofc, radius, center, "green");
    currentMapShapes.push(...posShapes);
    const negShapes = addIsoLines(viewer, mapData.mapFofc, -sigmaFofc, radius, center, "red");
    currentMapShapes.push(...negShapes);
  }

  viewer.render();
}

export function removeDensityMap(viewer) {
  for (const shape of currentMapShapes) {
    viewer.removeShape(shape);
  }
  currentMapShapes = [];
  viewer.render();
}
