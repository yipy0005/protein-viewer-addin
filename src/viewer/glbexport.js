/* global THREE */

/**
 * GLB Exporter — converts 3Dmol.js model to GLB via Three.js.
 * Requires Three.js r128 + BufferGeometryUtils + GLTFExporter loaded via CDN.
 */

const ELEM_COLORS = {
  H: 0xffffff, C: 0x909090, N: 0x3050f8, O: 0xff0d0d,
  S: 0xffff30, P: 0xff8000, F: 0x90e050, Cl: 0x1ff01f,
  Br: 0xa62929, I: 0x940094, Fe: 0xe06633, Zn: 0x7d80b0,
};

const ELEM_RADII = {
  H: 0.25, C: 0.7, N: 0.65, O: 0.6,
  S: 1.0, P: 1.0, F: 0.5, Cl: 0.8, Br: 0.9, I: 1.0,
  Fe: 0.65, Zn: 0.65,
};

export async function exportToGLB(viewerInstance) {
  const model = viewerInstance.getModel();
  const allAtoms = model.selectedAtoms({});
  if (!allAtoms || !allAtoms.length) throw new Error("No atoms to export");

  const atoms = allAtoms.filter((a) => a.elem !== "H");
  const scene = new THREE.Scene();

  scene.add(new THREE.AmbientLight(0x606060));
  const light = new THREE.DirectionalLight(0xffffff, 0.8);
  light.position.set(10, 10, 10);
  scene.add(light);

  const sphereGeo = new THREE.SphereGeometry(1, 10, 8);
  const atomScale = 0.35;
  const bondRadius = 0.06;

  // Group atoms by element, merge geometries
  const groups = {};
  for (const a of atoms) {
    const e = a.elem || "C";
    if (!groups[e]) groups[e] = [];
    groups[e].push(a);
  }

  for (const [elem, elemAtoms] of Object.entries(groups)) {
    const r = (ELEM_RADII[elem] || 0.7) * atomScale;
    const geos = elemAtoms.map((a) => {
      const g = sphereGeo.clone();
      g.scale(r, r, r);
      g.translate(a.x, a.y, a.z);
      return g;
    });
    const merged = THREE.BufferGeometryUtils.mergeBufferGeometries(geos, false);
    if (merged) {
      scene.add(new THREE.Mesh(merged, new THREE.MeshStandardMaterial({
        color: ELEM_COLORS[elem] || 0x808080, roughness: 0.4, metalness: 0.1,
      })));
    }
  }

  // Bonds
  const serialMap = {};
  allAtoms.forEach((a, i) => { serialMap[i] = a; });
  const seen = new Set();
  const yAxis = new THREE.Vector3(0, 1, 0);
  const bondGeos = [];

  for (let i = 0; i < allAtoms.length; i++) {
    const a1 = allAtoms[i];
    if (a1.elem === "H" || !a1.bonds) continue;
    for (const j of a1.bonds) {
      const a2 = serialMap[j];
      if (!a2 || a2.elem === "H") continue;
      const key = Math.min(i, j) + ":" + Math.max(i, j);
      if (seen.has(key)) continue;
      seen.add(key);

      const start = new THREE.Vector3(a1.x, a1.y, a1.z);
      const end = new THREE.Vector3(a2.x, a2.y, a2.z);
      const dir = new THREE.Vector3().subVectors(end, start);
      const len = dir.length();
      if (len < 0.1 || len > 4.0) continue;

      const mid = new THREE.Vector3().addVectors(start, end).multiplyScalar(0.5);
      const geo = new THREE.CylinderGeometry(bondRadius, bondRadius, len, 4, 1);
      const quat = new THREE.Quaternion().setFromUnitVectors(yAxis, dir.clone().normalize());
      geo.applyMatrix4(new THREE.Matrix4().makeRotationFromQuaternion(quat));
      geo.translate(mid.x, mid.y, mid.z);
      bondGeos.push(geo);
    }
  }

  if (bondGeos.length > 0) {
    const merged = THREE.BufferGeometryUtils.mergeBufferGeometries(bondGeos, false);
    if (merged) {
      scene.add(new THREE.Mesh(merged, new THREE.MeshStandardMaterial({
        color: 0x606060, roughness: 0.5, metalness: 0.1,
      })));
    }
  }

  // Center
  const box = new THREE.Box3().setFromObject(scene);
  scene.position.sub(box.getCenter(new THREE.Vector3()));

  // Export
  return new Promise((resolve, reject) => {
    try {
      const exporter = new THREE.GLTFExporter();
      exporter.parse(scene, (result) => {
        if (result instanceof ArrayBuffer) resolve(result);
        else reject(new Error("GLB export returned JSON instead of binary"));
      }, { binary: true });
    } catch (err) { reject(err); }
  });
}

export function downloadGLB(buffer, filename) {
  const blob = new Blob([buffer], { type: "model/gltf-binary" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename || "molecule.glb";
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
}
