// Research-only observation of a native producer's drawing snapshot. This is
// neither an ODF importer nor a layout algorithm. Names are association evidence,
// never stable semantic identities; transformed geometry stays uninterpreted.
const units = { pt: 1, pc: 12, in: 72, cm: 72 / 2.54, mm: 72 / 25.4 };
const encoder = new TextEncoder();
const invalid = message => { throw new TypeError(message); };
function plain(value, fields) {
  if (!value || ![Object.prototype, null].includes(Object.getPrototypeOf(value)) || Array.isArray(value)) invalid("Expected a plain geometry observation record.");
  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
    if (!fields.includes(key) || !("value" in descriptor)) invalid(`Unknown observation field/accessor: ${key}.`);
  }
}
export function nativeLengthPoints(value) {
  if (typeof value !== "string") return undefined;
  const match = /^([+-]?(?:\d+(?:\.\d*)?|\.\d+))(pt|pc|in|cm|mm)$/.exec(value.trim());
  if (!match) return value.trim() === "0" ? 0 : undefined;
  const result = Number(match[1]) * units[match[2]];
  return Number.isFinite(result) && Math.abs(result) <= 1e9 ? result : undefined;
}

export function observeDiagramGeometry(snapshot, model, options = {}) {
  plain(snapshot, ["frame", "shapes"]); plain(options, ["maxShapes", "maxWork", "maxTextBytes", "tolerancePoints"]);
  plain(snapshot.frame, ["x", "y", "width", "height"]);
  const limits = { maxShapes: 10000, maxWork: 200000, maxTextBytes: 1000000, tolerancePoints: 0.1, ...options };
  for (const key of ["maxShapes", "maxWork", "maxTextBytes"]) if (!Number.isSafeInteger(limits[key]) || limits[key] < 1) invalid(`Invalid ${key}.`);
  if (!Number.isFinite(limits.tolerancePoints) || limits.tolerancePoints < 0 || limits.tolerancePoints > 100) invalid("Invalid geometry tolerance.");
  for (const [key, value] of Object.entries(snapshot.frame)) if (!Number.isFinite(value) || Math.abs(value) > 1e9 || (["width", "height"].includes(key) && value <= 0)) invalid(`Invalid frame ${key}.`);
  if (Object.keys(snapshot.frame).length !== 4 || !Array.isArray(snapshot.shapes) || snapshot.shapes.length > limits.maxShapes) invalid("Incomplete frame or over-budget drawing snapshot.");
  let work = 0, textBytes = 0;
  const step = () => { if (++work > limits.maxWork) throw new RangeError("Geometry observation exceeds maxWork."); };
  const index = new Map(), sourceIndices = new Set();
  if (!model || !Array.isArray(model.points)) invalid("Source model points are required.");
  for (const point of model.points) {
    step();
    if (!Number.isSafeInteger(point.elementIndex) || point.elementIndex < 1 || sourceIndices.has(point.elementIndex)) invalid("Source point indices must be distinct positive integers.");
    sourceIndices.add(point.elementIndex);
    if (typeof point.modelId === "string") { const bucket = index.get(point.modelId) ?? []; bucket.push(point); index.set(point.modelId, bucket); }
  }
  const occurrences = new Set(), names = new Map(), captured = [];
  for (const raw of snapshot.shapes) {
    step(); plain(raw, ["ordinal", "kind", "name", "x", "y", "width", "height", "transform", "ancestorTransforms", "paragraphs", "preset", "textAreas", "mirrorHorizontal", "mirrorVertical"]);
    if (!Number.isSafeInteger(raw.ordinal) || raw.ordinal < 1 || occurrences.has(raw.ordinal)) invalid("Drawing occurrences need unique positive source ordinals.");
    occurrences.add(raw.ordinal);
    for (const key of ["kind", "name", "x", "y", "width", "height", "transform", "preset", "textAreas", "mirrorHorizontal", "mirrorVertical"]) {
      if (raw[key] !== undefined && typeof raw[key] !== "string") invalid(`Invalid raw ${key}.`);
      if (raw[key] !== undefined) { textBytes += encoder.encode(raw[key]).byteLength; if (textBytes > limits.maxTextBytes) throw new RangeError("Geometry observation exceeds maxTextBytes."); }
    }
    if (typeof raw.kind !== "string" || !Array.isArray(raw.paragraphs) || !Array.isArray(raw.ancestorTransforms)) invalid("Drawing kind, paragraphs and ancestor transforms are required.");
    for (const value of [...raw.paragraphs, ...raw.ancestorTransforms]) {
      step(); if (typeof value !== "string") invalid("Drawing text/transforms must be strings.");
      textBytes += encoder.encode(value).byteLength; if (textBytes > limits.maxTextBytes) throw new RangeError("Geometry observation exceeds maxTextBytes.");
    }
    const drawing = { ...raw, paragraphs: [...raw.paragraphs], ancestorTransforms: [...raw.ancestorTransforms] };
    if (drawing.name !== undefined && drawing.name !== "") { const bucket = names.get(drawing.name) ?? []; bucket.push(drawing.ordinal); names.set(drawing.name, bucket); }
    captured.push(drawing);
  }
  captured.sort((a, b) => { step(); return a.ordinal - b.ordinal; });
  for (const ordinals of names.values()) ordinals.sort((a, b) => { step(); return a - b; });
  const shapes = captured.map(raw => {
    step();
    const bounds = Object.fromEntries(["x", "y", "width", "height"].map(key => [key, nativeLengthPoints(raw[key])]));
    const transformed = !!raw.transform?.trim() || raw.ancestorTransforms.some(value => value.trim());
    const complete = Object.values(bounds).every(value => value !== undefined) && bounds.width >= 0 && bounds.height >= 0;
    const geometryStatus = transformed ? "transform-source-only" : complete ? "axis-aligned" : "incomplete";
    const matches = index.get(raw.name) ?? [], repeatedNameOrdinals = names.get(raw.name) ?? [];
    const associationStatus = model.status !== "inspected" ? "source-model-incomplete" : raw.name === undefined || raw.name === "" ? "missing-name"
      : matches.length === 0 ? "unmatched-name" : matches.length > 1 || matches.some(p => p.identityStatus !== "unique") ? "ambiguous-source-identity"
      : matches[0].kind !== "pres" ? "non-presentation-name" : "presentation-name-match";
    let outsideFrame;
    if (geometryStatus === "axis-aligned") {
      const f = snapshot.frame, t = limits.tolerancePoints;
      outsideFrame = bounds.x < f.x - t || bounds.y < f.y - t || bounds.x + bounds.width > f.x + f.width + t || bounds.y + bounds.height > f.y + f.height + t;
    }
    return { occurrence: raw.ordinal, raw, geometryStatus,
      ...(geometryStatus === "axis-aligned" ? { boundsPoints: bounds, outsideFrame } : {}),
      association: { status: associationStatus, evidence: "exact-producer-name-only",
        sourceElementIndices: matches.map(p => { step(); return p.elementIndex; }),
        repeatedNameOrdinals: repeatedNameOrdinals.map(ordinal => { step(); return ordinal; }) } };
  });
  return { profile: "native-diagram-geometry-observation-v1", coordinateSpace: "page-points", framePoints: { ...snapshot.frame }, shapes,
    summary: { shapes: shapes.length, axisAligned: shapes.filter(s => s.geometryStatus === "axis-aligned").length,
      transformed: shapes.filter(s => s.geometryStatus === "transform-source-only").length,
      incomplete: shapes.filter(s => s.geometryStatus === "incomplete").length,
      outsideFrame: shapes.filter(s => s.outsideFrame).length,
      repeatedNames: [...names.values()].filter(bucket => bucket.length > 1).length },
    limitations: ["Shape names provide source association evidence, not stable authoring/node identities.", "Transforms, enhanced geometry equations, text-area expressions and clipping are retained without evaluation.", "Axis-aligned boxes do not establish glyph bounds, legibility, connector topology, cache freshness or native editability."] };
}

export function compareNativeGeometryObservations(left, right) {
  for (const value of [left, right]) if (value?.profile !== "native-diagram-geometry-observation-v1" || value.coordinateSpace !== "page-points") invalid("Expected versioned native geometry observations.");
  const signature = observation => ({ frame: observation.framePoints, shapes: observation.shapes.map(s => ({
    geometryStatus: s.geometryStatus, boundsPoints: s.boundsPoints,
    // Transformed boxes are not resolved, so their raw sizes/positions must
    // participate: matching rotation strings alone cannot prove equal geometry.
    rawFrame: s.geometryStatus === "axis-aligned" ? undefined : [s.raw.x, s.raw.y, s.raw.width, s.raw.height],
    sourceElementIndices: s.association.sourceElementIndices, associationStatus: s.association.status,
    kind: s.raw.kind, preset: s.raw.preset, paragraphs: s.raw.paragraphs,
    transform: s.raw.transform, ancestorTransforms: s.raw.ancestorTransforms,
    textAreas: s.raw.textAreas, mirrorHorizontal: s.raw.mirrorHorizontal, mirrorVertical: s.raw.mirrorVertical,
  })) });
  return { exactObservedGeometryEqual: JSON.stringify(signature(left)) === JSON.stringify(signature(right)),
    policy: "Drawing order, source association evidence, untransformed point bounds and raw transformed frames/transform/text-area/mirror/text records; unmatched producer names and raw source ordinals are excluded." };
}
