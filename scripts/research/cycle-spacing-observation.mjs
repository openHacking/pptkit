// Diagnose an observed ordered set of circles. Inferred parameters describe the
// snapshot only; never feed them back as an automatic native-layout correction.
const invalid = message => { throw new TypeError(message); };
function plain(value, fields) {
  if (!value || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) invalid("Expected a plain spacing observation.");
  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) if (!fields.includes(key) || !("value" in descriptor)) invalid(`Unknown spacing field/accessor ${key}.`);
}
export function observeCycleSpacing(snapshot, options = {}) {
  plain(snapshot, ["frame", "circles"]); plain(options, ["quantumPoints", "fitTolerancePoints", "comparisonTolerancePoints", "maxNodes", "maxWork"]);
  const q = options.quantumPoints ?? 1 / 12700, tolerance = options.fitTolerancePoints ?? 0.001, comparisonTolerance = options.comparisonTolerancePoints ?? 0.1;
  for (const value of [q, tolerance, comparisonTolerance]) if (!Number.isFinite(value) || value <= 0 || value > 10) invalid("Invalid spacing tolerance/quantum.");
  const maxNodes = options.maxNodes ?? 10000, maxWork = options.maxWork ?? 100000;
  for (const value of [maxNodes, maxWork]) if (!Number.isSafeInteger(value) || value < 1) invalid("Invalid spacing budget.");
  let work = 0; const step = () => { if (++work > maxWork) throw new RangeError("Spacing observation exceeds maxWork."); };
  const box = value => { plain(value, ["x", "y", "width", "height"]); for (const key of ["x", "y", "width", "height"]) if (!Number.isFinite(value[key]) || Math.abs(value[key]) > 1e7 || (["width", "height"].includes(key) && value[key] <= 0)) invalid("Invalid spacing box."); return { ...value }; };
  const frame = box(snapshot.frame);
  if (!Array.isArray(snapshot.circles) || snapshot.circles.length < 3 || snapshot.circles.length > maxNodes) invalid("Ordered three-or-more circle observations required.");
  const ids = new Set(), circles = snapshot.circles.map(value => {
    step(); plain(value, ["id", "box", "proposed", "strokeWidth", "strokeAlignment"]);
    if (typeof value.id !== "string" || !value.id.length || value.id.length > 256 || ids.has(value.id)) invalid("Invalid/duplicate observed identity."); ids.add(value.id);
    const b = box(value.box), proposed = box(value.proposed);
    if (Math.abs(b.width - b.height) > tolerance) invalid("Noncircular geometry needs another profile.");
    if (value.strokeWidth !== undefined && (!Number.isFinite(value.strokeWidth) || value.strokeWidth < 0 || value.strokeWidth > 10000)) invalid("Invalid observed stroke.");
    if (value.strokeAlignment !== undefined && typeof value.strokeAlignment !== "string") invalid("Invalid raw stroke alignment.");
    return { id: value.id, box: b, proposed, ...(value.strokeWidth === undefined ? {} : { strokeWidth: value.strokeWidth }), ...(value.strokeAlignment === undefined ? {} : { strokeAlignment: value.strokeAlignment }), x: b.x + b.width / 2, y: b.y + b.height / 2 };
  });
  const n = circles.length, cx = circles.reduce((sum, c) => sum + c.x, 0) / n, cy = circles.reduce((sum, c) => sum + c.y, 0) / n;
  const diameter = circles.reduce((sum, c) => sum + c.box.width, 0) / n;
  let radius = 0;
  for (const [index, c] of circles.entries()) { step(); const angle = 2 * Math.PI * index / n; radius += ((c.x - cx) * Math.sin(angle) - (c.y - cy) * Math.cos(angle)) / n; }
  let maxFitResidual = 0, maxCoordinateDelta = 0, roundedMaxCoordinateDelta = 0, maxRoundingAdjustment = 0, minimumBoxInset = Infinity;
  const rows = circles.map((c, i) => {
    step(); const angle = 2 * Math.PI * i / n;
    const expected = { x: cx + radius * Math.sin(angle) - diameter / 2, y: cy - radius * Math.cos(angle) - diameter / 2, width: diameter, height: diameter };
    const rounding = {};
    for (const key of ["x", "y", "width", "height"]) {
      step(); maxFitResidual = Math.max(maxFitResidual, Math.abs(c.box[key] - expected[key]));
      const value = Math.round(c.proposed[key] / q) * q;
      maxCoordinateDelta = Math.max(maxCoordinateDelta, Math.abs(c.proposed[key] - c.box[key])); roundedMaxCoordinateDelta = Math.max(roundedMaxCoordinateDelta, Math.abs(value - c.box[key]));
      maxRoundingAdjustment = Math.max(maxRoundingAdjustment, Math.abs(value - c.proposed[key])); rounding[key] = value;
    }
    const inset = Math.min(c.box.x - frame.x, c.box.y - frame.y, frame.x + frame.width - c.box.x - c.box.width, frame.y + frame.height - c.box.y - c.box.height);
    minimumBoxInset = Math.min(minimumBoxInset, inset);
    return { id: c.id, fitted: expected, nearestQuantumProposed: rounding, minimumBoxInset: inset,
      ...(c.strokeWidth === undefined ? {} : { strokeWidth: c.strokeWidth, strokeAlignment: c.strokeAlignment ?? "unspecified", centeredStrokeInsetHypothesis: inset - c.strokeWidth / 2 }) };
  });
  const consistent = radius > 0 && maxFitResidual <= tolerance;
  return { profile: "cycle-spacing-observation-v1", consistentOrderedCircleRing: consistent, circleCount: n, center: { x: cx, y: cy }, diameter, radius,
    ...(consistent ? { observedGapRatio: 2 * radius * Math.sin(Math.PI / n) / diameter - 1 } : {}), maxFitResidual, fitTolerancePoints: tolerance, minimumBoxInset,
    rounding: { quantumPoints: q, maxCoordinateDelta, roundedMaxCoordinateDelta, maxRoundingAdjustment, comparisonTolerancePoints: comparisonTolerance, nearestQuantumMeetsComparison: roundedMaxCoordinateDelta <= comparisonTolerance }, rows, work,
    limitations: ["Parameters are observed from one ordered snapshot, not native solver rules or correction constants.", "Stroke inset is explicitly a centered-stroke hypothesis when raw alignment is unspecified.", "Ring fitting and EMU rounding do not explain text fitting, arrows, paints or cache freshness."] };
}
