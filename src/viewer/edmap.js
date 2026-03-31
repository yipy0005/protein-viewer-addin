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
 * Get the center point for map extraction.
 * If a selection is provided, use that; otherwise use model center of mass.
 */
function getMapCenter(viewer, centerAtom) {
  if (centerAtom) return centerAtom;
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
 * custom shape in 3Dmol.js for performance.
 */
function addIsoLines(viewer, wasmMap, sigma, radius, center, color) {
  const isolevel = wasmMap.mean + sigma * wasmMap.rms;
  const ok = wasmMap.extract_isosurface(radius, center.x, center.y, center.z, isolevel, "");
  if (!ok) return null;

  const verts = wasmMap.isosurface_vertices();
  const segs = wasmMap.isosurface_segments();
  if (!verts || !segs || verts.length === 0 || segs.length === 0) return null;

  const vertexArr = [];
  const normalArr = [];
  const faceArr = [];
  const colorArr = [];

  // Standard crystallography colors — lighter and easier on the eyes
  const colorMap = {
    "2fofc": { r: 0.4, g: 0.7, b: 1.0 },   // light sky blue
    "fofc_pos": { r: 0.2, g: 0.9, b: 0.3 }, // bright green
    "fofc_neg": { r: 1.0, g: 0.3, b: 0.3 }, // bright red
  };
  const col = colorMap[color] || { r: 0.4, g: 0.7, b: 1.0 };

  const t = 0.015; // line thickness in Å — thin chicken-wire look

  for (let i = 0; i < segs.length; i += 2) {
    const i0 = segs[i] * 3;
    const i1 = segs[i + 1] * 3;
    const x0 = verts[i0], y0 = verts[i0 + 1], z0 = verts[i0 + 2];
    const x1 = verts[i1], y1 = verts[i1 + 1], z1 = verts[i1 + 2];

    const dx = x1 - x0, dy = y1 - y0, dz = z1 - z0;
    const len = Math.sqrt(dx * dx + dy * dy + dz * dz);
    if (len < 0.001) continue;

    // Perpendicular offset
    let px = 0, py = 1, pz = 0;
    if (Math.abs(dy) > Math.abs(dx) && Math.abs(dy) > Math.abs(dz)) {
      px = 1; py = 0; pz = 0;
    }
    const cx = dy * pz - dz * py;
    const cy = dz * px - dx * pz;
    const cz = dx * py - dy * px;
    const cl = Math.sqrt(cx * cx + cy * cy + cz * cz);
    if (cl < 0.0001) continue;
    const ox = (cx / cl) * t, oy = (cy / cl) * t, oz = (cz / cl) * t;

    const base = vertexArr.length;
    const n = new $3Dmol.Vector3(ox / t, oy / t, oz / t);
    vertexArr.push(
      new $3Dmol.Vector3(x0 + ox, y0 + oy, z0 + oz),
      new $3Dmol.Vector3(x0 - ox, y0 - oy, z0 - oz),
      new $3Dmol.Vector3(x1 + ox, y1 + oy, z1 + oz),
      new $3Dmol.Vector3(x1 - ox, y1 - oy, z1 - oz)
    );
    normalArr.push(n, n, n, n);
    colorArr.push(col, col, col, col);
    // Double-sided: front and back faces
    faceArr.push(base, base + 1, base + 2, base + 1, base + 3, base + 2);
    faceArr.push(base + 2, base + 1, base, base + 2, base + 3, base + 1);
  }

  if (faceArr.length === 0) return null;

  const shape = viewer.addCustom({
    vertexArr: vertexArr,
    normalArr: normalArr,
    faceArr: faceArr,
    color: colorArr,
  });
  return shape;
}

let currentMapShapes = [];

export function renderDensityMap(viewer, mapData, opts) {
  const sigma2fofc = opts.sigma2fofc || 1.5;
  const sigmaFofc = opts.sigmaFofc || 3.0;
  const showFofc = opts.showFofc || false;
  const radius = opts.radius || 8;
  const center = getMapCenter(viewer, opts.center);

  // 2Fo-Fc (light blue)
  if (mapData.map2fofc) {
    const shape = addIsoLines(viewer, mapData.map2fofc, sigma2fofc, radius, center, "2fofc");
    if (shape) currentMapShapes.push(shape);
  }

  // Fo-Fc positive (green) and negative (red)
  if (showFofc && mapData.mapFofc) {
    const posShape = addIsoLines(viewer, mapData.mapFofc, sigmaFofc, radius, center, "fofc_pos");
    if (posShape) currentMapShapes.push(posShape);
    const negShape = addIsoLines(viewer, mapData.mapFofc, -sigmaFofc, radius, center, "fofc_neg");
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

/**
 * Extract isosurface geometry as a serializable object (no 3Dmol types).
 * Returns array of { verts: Float32Array, segs: Uint32Array, color: string } or null.
 */
export function extractIsoGeometry(mapData, opts) {
  const sigma2fofc = opts.sigma2fofc || 1.5;
  const sigmaFofc = opts.sigmaFofc || 3.0;
  const showFofc = opts.showFofc || false;
  const radius = opts.radius || 8;
  const center = opts.center || { x: 0, y: 0, z: 0 };

  const layers = [];

  function extract(wasmMap, sigma, colorKey) {
    const isolevel = wasmMap.mean + sigma * wasmMap.rms;
    const ok = wasmMap.extract_isosurface(radius, center.x, center.y, center.z, isolevel, "");
    if (!ok) return;
    const v = wasmMap.isosurface_vertices();
    const s = wasmMap.isosurface_segments();
    if (!v || !s || v.length === 0 || s.length === 0) return;
    layers.push({ verts: Array.from(v), segs: Array.from(s), color: colorKey });
  }

  if (mapData.map2fofc) extract(mapData.map2fofc, sigma2fofc, "2fofc");
  if (showFofc && mapData.mapFofc) {
    extract(mapData.mapFofc, sigmaFofc, "fofc_pos");
    extract(mapData.mapFofc, -sigmaFofc, "fofc_neg");
  }

  return layers.length ? layers : null;
}

/**
 * Render pre-extracted isosurface geometry in a 3Dmol.js viewer.
 * No gemmi needed — just the serialized layers from extractIsoGeometry.
 */
let contentMapShapes = [];

export function renderIsoGeometry(viewer, layers) {
  if (!layers || !viewer) return;

  const colorMap = {
    "2fofc": { r: 0.4, g: 0.7, b: 1.0 },
    "fofc_pos": { r: 0.2, g: 0.9, b: 0.3 },
    "fofc_neg": { r: 1.0, g: 0.3, b: 0.3 },
  };
  const t = 0.015;

  for (const layer of layers) {
    const verts = layer.verts;
    const segs = layer.segs;
    const col = colorMap[layer.color] || { r: 0.4, g: 0.7, b: 1.0 };

    const vertexArr = [];
    const normalArr = [];
    const faceArr = [];
    const colorArr = [];

    for (let i = 0; i < segs.length; i += 2) {
      const i0 = segs[i] * 3;
      const i1 = segs[i + 1] * 3;
      const x0 = verts[i0], y0 = verts[i0 + 1], z0 = verts[i0 + 2];
      const x1 = verts[i1], y1 = verts[i1 + 1], z1 = verts[i1 + 2];

      const dx = x1 - x0, dy = y1 - y0, dz = z1 - z0;
      const len = Math.sqrt(dx * dx + dy * dy + dz * dz);
      if (len < 0.001) continue;

      let px = 0, py = 1, pz = 0;
      if (Math.abs(dy) > Math.abs(dx) && Math.abs(dy) > Math.abs(dz)) { px = 1; py = 0; }
      const cx = dy * pz - dz * py;
      const cy = dz * px - dx * pz;
      const cz = dx * py - dy * px;
      const cl = Math.sqrt(cx * cx + cy * cy + cz * cz);
      if (cl < 0.0001) continue;
      const ox = (cx / cl) * t, oy = (cy / cl) * t, oz = (cz / cl) * t;

      const base = vertexArr.length;
      const n = new $3Dmol.Vector3(ox / t, oy / t, oz / t);
      vertexArr.push(
        new $3Dmol.Vector3(x0 + ox, y0 + oy, z0 + oz),
        new $3Dmol.Vector3(x0 - ox, y0 - oy, z0 - oz),
        new $3Dmol.Vector3(x1 + ox, y1 + oy, z1 + oz),
        new $3Dmol.Vector3(x1 - ox, y1 - oy, z1 - oz)
      );
      normalArr.push(n, n, n, n);
      colorArr.push(col, col, col, col);
      faceArr.push(base, base + 1, base + 2, base + 1, base + 3, base + 2);
      faceArr.push(base + 2, base + 1, base, base + 2, base + 3, base + 1);
    }

    if (faceArr.length === 0) continue;
    const shape = viewer.addCustom({ vertexArr, normalArr, faceArr, color: colorArr });
    contentMapShapes.push(shape);
  }
  viewer.render();
}

export function removeIsoGeometry(viewer) {
  for (const shape of contentMapShapes) {
    viewer.removeShape(shape);
  }
  contentMapShapes = [];
}
