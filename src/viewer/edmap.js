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
 * Extract isosurface from a gemmi wasm map and render as a single batched
 * line shape in 3Dmol.js for performance.
 */
function addIsoLines(viewer, wasmMap, sigma, radius, center, color) {
  const isolevel = wasmMap.mean + sigma * wasmMap.rms;
  const ok = wasmMap.extract_isosurface(radius, center.x, center.y, center.z, isolevel, "");
  if (!ok) return null;

  const verts = wasmMap.isosurface_vertices();
  const segs = wasmMap.isosurface_segments();
  if (!verts || !segs || verts.length === 0 || segs.length === 0) return null;

  // Build a single custom shape with all line segments
  const vertexArr = [];
  const normalArr = [];
  const faceArr = [];
  const colorArr = [];

  // Parse the color to get r,g,b
  const c = $3Dmol.CC.color(color);
  const cr = c.r, cg = c.g, cb = c.b;

  // For each line segment, create a thin triangle pair (degenerate quad)
  // This renders as lines but in a single draw call
  for (let i = 0; i < segs.length; i += 2) {
    const i0 = segs[i] * 3;
    const i1 = segs[i + 1] * 3;
    const x0 = verts[i0], y0 = verts[i0 + 1], z0 = verts[i0 + 2];
    const x1 = verts[i1], y1 = verts[i1 + 1], z1 = verts[i1 + 2];

    // Create a very thin quad (two triangles) to represent the line
    const dx = x1 - x0, dy = y1 - y0, dz = z1 - z0;
    const len = Math.sqrt(dx * dx + dy * dy + dz * dz);
    if (len < 0.001) continue;

    // Perpendicular offset (tiny, ~0.02 Å)
    let px, py, pz;
    if (Math.abs(dy) > Math.abs(dx)) {
      px = 1; py = 0; pz = 0;
    } else {
      px = 0; py = 1; pz = 0;
    }
    // Cross product for perpendicular
    const cx = dy * pz - dz * py;
    const cy = dz * px - dx * pz;
    const cz = dx * py - dy * px;
    const cl = Math.sqrt(cx * cx + cy * cy + cz * cz);
    const t = 0.02; // thickness
    const ox = (cx / cl) * t, oy = (cy / cl) * t, oz = (cz / cl) * t;

    const base = vertexArr.length / 3;
    vertexArr.push(x0 + ox, y0 + oy, z0 + oz);
    vertexArr.push(x0 - ox, y0 - oy, z0 - oz);
    vertexArr.push(x1 + ox, y1 + oy, z1 + oz);
    vertexArr.push(x1 - ox, y1 - oy, z1 - oz);
    normalArr.push(ox, oy, oz, -ox, -oy, -oz, ox, oy, oz, -ox, -oy, -oz);
    colorArr.push(cr, cg, cb, cr, cg, cb, cr, cg, cb, cr, cg, cb);
    faceArr.push(base, base + 1, base + 2, base + 1, base + 3, base + 2);
  }

  if (faceArr.length === 0) return null;

  const shape = viewer.addCustom({
    vertexArr, normalArr, faceArr, color: colorArr,
  });
  return shape;
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
    const shape = addIsoLines(viewer, mapData.map2fofc, sigma2fofc, radius, center, "blue");
    if (shape) currentMapShapes.push(shape);
  }

  // Fo-Fc positive (green) and negative (red)
  if (showFofc && mapData.mapFofc) {
    const posShape = addIsoLines(viewer, mapData.mapFofc, sigmaFofc, radius, center, "green");
    if (posShape) currentMapShapes.push(posShape);
    const negShape = addIsoLines(viewer, mapData.mapFofc, -sigmaFofc, radius, center, "red");
    if (negShape) currentMapShapes.push(negShape);
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
