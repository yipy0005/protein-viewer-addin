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
 * Returns the raw ArrayBuffer so 3Dmol.js can parse it natively.
 */
export function parseCcp4(gemmi, arrayBuffer) {
  // Validate it's a real CCP4 file by trying to read it with gemmi
  const map = gemmi.readCcp4Map(arrayBuffer, true);
  map.delete();
  return { rawCcp4Buffer: arrayBuffer };
}

/**
 * Create a VolumeData from a gemmi MtzMap by populating it directly.
 */
function volumeDataFromWasmMap(wasmMap) {
  const nx = wasmMap.nx;
  const ny = wasmMap.ny;
  const nz = wasmMap.nz;
  const cell = wasmMap.cell;
  const rawData = wasmMap.data();

  // Build a proper CCP4 binary that 3Dmol.js can parse
  const headerSize = 1024;
  const dataLen = nx * ny * nz;
  const totalSize = headerSize + dataLen * 4;
  const buf = new ArrayBuffer(totalSize);
  const intView = new Int32Array(buf, 0, 256);
  const floatView = new Float32Array(buf, 0, 256);
  const dv = new DataView(buf);

  // NX, NY, NZ (fast, medium, slow)
  intView[0] = nx;  // NC - columns (fastest)
  intView[1] = ny;  // NR - rows
  intView[2] = nz;  // NS - sections (slowest)
  intView[3] = 2;   // MODE 2 = 32-bit float

  // Start indices (map covers full unit cell)
  intView[4] = 0;   // NCSTART
  intView[5] = 0;   // NRSTART
  intView[6] = 0;   // NSSTART

  // Grid sampling along unit cell
  intView[7] = nx;   // MX
  intView[8] = ny;   // MY
  intView[9] = nz;   // MZ

  // Cell dimensions
  floatView[10] = cell.a;
  floatView[11] = cell.b;
  floatView[12] = cell.c;
  floatView[13] = cell.alpha;
  floatView[14] = cell.beta;
  floatView[15] = cell.gamma;

  // Axis correspondence: gemmi outputs data in X,Y,Z order
  intView[16] = 1;  // MAPC = X
  intView[17] = 2;  // MAPR = Y
  intView[18] = 3;  // MAPS = Z

  // Density stats
  let dmin = Infinity, dmax = -Infinity, dmean = 0;
  for (let i = 0; i < rawData.length; i++) {
    const v = rawData[i];
    if (v < dmin) dmin = v;
    if (v > dmax) dmax = v;
    dmean += v;
  }
  dmean /= rawData.length;
  floatView[19] = dmin;
  floatView[20] = dmax;
  floatView[21] = dmean;

  intView[22] = 1;  // ISPG (space group P1 for the map)
  intView[23] = 0;  // NSYMBT

  // MAP signature
  dv.setUint8(208, 77);  // 'M'
  dv.setUint8(209, 65);  // 'A'
  dv.setUint8(210, 80);  // 'P'
  dv.setUint8(211, 32);  // ' '

  // Machine stamp (little-endian)
  dv.setUint8(212, 68);  // 0x44
  dv.setUint8(213, 65);  // 0x41

  // RMS
  let rmsSum = 0;
  for (let i = 0; i < rawData.length; i++) {
    const d = rawData[i] - dmean;
    rmsSum += d * d;
  }
  floatView[54] = Math.sqrt(rmsSum / rawData.length);

  // Copy density data
  const dataView = new Float32Array(buf, headerSize);
  dataView.set(rawData);

  return new $3Dmol.VolumeData(new Int8Array(buf), "ccp4");
}

let currentIsoShapes = [];

/**
 * Render electron density isosurfaces in a 3Dmol.js viewer.
 */
export function renderDensityMap(viewer, mapData, opts) {
  const sigma2fofc = opts.sigma2fofc || 1.5;
  const sigmaFofc = opts.sigmaFofc || 3.0;
  const showFofc = opts.showFofc || false;

  // 2Fo-Fc map (blue mesh)
  if (mapData.rawCcp4Buffer) {
    // Native CCP4 file — let 3Dmol.js parse it directly
    const volData = new $3Dmol.VolumeData(new Int8Array(mapData.rawCcp4Buffer), "ccp4");
    const mean = volData.data.reduce((s, v) => s + v, 0) / volData.data.length;
    let rmsSum = 0;
    for (let i = 0; i < volData.data.length; i++) { const d = volData.data[i] - mean; rmsSum += d * d; }
    const rms = Math.sqrt(rmsSum / volData.data.length);
    const isoVal = mean + sigma2fofc * rms;
    const iso = viewer.addIsosurface(volData, { isoval: isoVal, color: "blue", alpha: 0.5, wireframe: true });
    currentIsoShapes.push(iso);
  } else if (mapData.map2fofc) {
    const volData = volumeDataFromWasmMap(mapData.map2fofc);
    const isoVal = mapData.map2fofc.mean + sigma2fofc * mapData.map2fofc.rms;
    const iso = viewer.addIsosurface(volData, { isoval: isoVal, color: "blue", alpha: 0.5, wireframe: true });
    currentIsoShapes.push(iso);
  }

  // Fo-Fc map (green +σ, red -σ)
  if (showFofc && mapData.mapFofc) {
    const fofcVol = volumeDataFromWasmMap(mapData.mapFofc);
    const fofcIsoVal = mapData.mapFofc.mean + sigmaFofc * mapData.mapFofc.rms;
    const isoPos = viewer.addIsosurface(fofcVol, { isoval: fofcIsoVal, color: "green", alpha: 0.5, wireframe: true });
    const isoNeg = viewer.addIsosurface(fofcVol, { isoval: -(fofcIsoVal - 2 * mapData.mapFofc.mean), color: "red", alpha: 0.5, wireframe: true });
    currentIsoShapes.push(isoPos, isoNeg);
  }

  viewer.render();
}

/**
 * Remove all density isosurfaces from the viewer.
 */
export function removeDensityMap(viewer) {
  for (const shape of currentIsoShapes) {
    viewer.removeShape(shape);
  }
  currentIsoShapes = [];
  viewer.render();
}
