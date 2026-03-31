/* global Gemmi, $3Dmol */

/**
 * Electron density map support using gemmi WASM.
 * Handles MTZ and CCP4 map files, renders isosurfaces via 3Dmol.js.
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

/**
 * Parse an MTZ file and compute electron density maps.
 * Returns { map2fofc: MtzMap|null, mapFofc: MtzMap|null, mtz: Mtz }
 */
export function parseMtz(gemmi, arrayBuffer) {
  const mtz = gemmi.readMtz(arrayBuffer);
  const map2fofc = mtz.calculate_wasm_map(false);
  const mapFofc = mtz.calculate_wasm_map(true);
  return { map2fofc, mapFofc, mtz };
}

/**
 * Parse a CCP4/MRC map file.
 * Returns a Ccp4Map object.
 */
export function parseCcp4(gemmi, arrayBuffer) {
  return gemmi.readCcp4Map(arrayBuffer, true);
}

/**
 * Build a CCP4 binary buffer from a gemmi MtzMap/Ccp4Map so 3Dmol.js can read it.
 * We extract the raw grid data and construct a minimal CCP4 header.
 */
function buildCcp4Buffer(wasmMap) {
  const nx = wasmMap.nx;
  const ny = wasmMap.ny;
  const nz = wasmMap.nz;
  const cell = wasmMap.cell;
  const rawData = wasmMap.data();
  const floatData = new Float32Array(rawData.length);
  floatData.set(rawData);

  // Compute stats
  let dmin = Infinity, dmax = -Infinity, dmean = 0;
  for (let i = 0; i < floatData.length; i++) {
    const v = floatData[i];
    if (v < dmin) dmin = v;
    if (v > dmax) dmax = v;
    dmean += v;
  }
  dmean /= floatData.length;
  let rms = 0;
  for (let i = 0; i < floatData.length; i++) {
    const d = floatData[i] - dmean;
    rms += d * d;
  }
  rms = Math.sqrt(rms / floatData.length);

  // CCP4 header: 256 4-byte words = 1024 bytes
  const headerSize = 1024;
  const totalSize = headerSize + floatData.length * 4;
  const buf = new ArrayBuffer(totalSize);
  const intView = new Int32Array(buf, 0, 256);
  const floatView = new Float32Array(buf, 0, 256);

  // Columns (NX), Rows (NY), Sections (NZ)
  intView[0] = nx;
  intView[1] = ny;
  intView[2] = nz;
  intView[3] = 2; // MODE = 2 (32-bit float)
  intView[4] = 0; // NXSTART
  intView[5] = 0; // NYSTART
  intView[6] = 0; // NZSTART
  intView[7] = nx; // MX (intervals along X)
  intView[8] = ny; // MY
  intView[9] = nz; // MZ
  floatView[10] = cell.a;
  floatView[11] = cell.b;
  floatView[12] = cell.c;
  floatView[13] = cell.alpha;
  floatView[14] = cell.beta;
  floatView[15] = cell.gamma;
  intView[16] = 1; // MAPC (axis for columns = X)
  intView[17] = 2; // MAPR (axis for rows = Y)
  intView[18] = 3; // MAPS (axis for sections = Z)
  floatView[19] = dmin;
  floatView[20] = dmax;
  floatView[21] = dmean;
  intView[22] = 1; // ISPG (space group)
  intView[23] = 0; // NSYMBT (no symmetry bytes)
  // Words 24-52: unused, leave as 0
  // Word 49-51: origin (leave 0)
  // Word 52: MAP string
  const dv = new DataView(buf);
  dv.setUint8(52 * 4, 77);     // 'M'
  dv.setUint8(52 * 4 + 1, 65); // 'A'
  dv.setUint8(52 * 4 + 2, 80); // 'P'
  dv.setUint8(52 * 4 + 3, 32); // ' '
  // Word 53: machine stamp (little-endian)
  dv.setUint8(53 * 4, 68);
  dv.setUint8(53 * 4 + 1, 65);
  // Word 54: RMS
  floatView[54] = rms;

  // Copy density data after header
  const dataView = new Float32Array(buf, headerSize);
  dataView.set(floatData);

  return buf;
}

/**
 * Render electron density isosurfaces in a 3Dmol.js viewer.
 * @param {object} viewer - 3Dmol.js viewer
 * @param {object} mapData - { map2fofc, mapFofc } from parseMtz or { ccp4Map } from parseCcp4
 * @param {object} opts - { sigma2fofc, sigmaFofc, showFofc, radius, center }
 * @returns {object} - { isoIds } for later removal
 */
export function renderDensityMap(viewer, mapData, opts) {
  const sigma2fofc = opts.sigma2fofc || 1.5;
  const sigmaFofc = opts.sigmaFofc || 3.0;
  const showFofc = opts.showFofc || false;

  const result = { isoIds: [] };

  // 2Fo-Fc map (blue mesh)
  const map = mapData.map2fofc || mapData.ccp4Map;
  if (map) {
    const ccp4Buf = buildCcp4Buffer(map);
    const volData = new $3Dmol.VolumeData(ccp4Buf, "ccp4");
    const rms = map.rms || volData.data.reduce((s, v) => s + v * v, 0) / volData.data.length;
    const isoVal = map.mean + sigma2fofc * map.rms;
    const iso = viewer.addIsosurface(volData, {
      isoval: isoVal,
      color: "blue",
      alpha: 0.5,
      wireframe: true,
    });
    result.isoIds.push(iso);
  }

  // Fo-Fc map (green +σ, red -σ)
  if (showFofc && mapData.mapFofc) {
    const fofcBuf = buildCcp4Buffer(mapData.mapFofc);
    const fofcVol = new $3Dmol.VolumeData(fofcBuf, "ccp4");
    const fofcIsoVal = mapData.mapFofc.mean + sigmaFofc * mapData.mapFofc.rms;
    const isoPos = viewer.addIsosurface(fofcVol, {
      isoval: fofcIsoVal,
      color: "green",
      alpha: 0.5,
      wireframe: true,
    });
    const isoNeg = viewer.addIsosurface(fofcVol, {
      isoval: -fofcIsoVal + 2 * mapData.mapFofc.mean,
      color: "red",
      alpha: 0.5,
      wireframe: true,
    });
    result.isoIds.push(isoPos, isoNeg);
  }

  viewer.render();
  return result;
}

/**
 * Remove all density isosurfaces from the viewer.
 */
export function removeDensityMap(viewer) {
  viewer.removeAllIsosurfaces();
  viewer.render();
}
