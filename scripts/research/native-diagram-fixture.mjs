import { diagramResearchProfiles } from "./diagram-layout-profiles.mjs";
import { diagramPresentationSpecimen } from "./native-diagram-presentation.mjs";
import { reconcileDiagramIdentities } from "./diagram-identity-state.mjs";

// Research specimens only. No geometry/layout/cache is inferred or borrowed from
// previous content. Full native admission is deliberately outside this writer.
const D = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DSP = "http://schemas.microsoft.com/office/drawing/2008/diagram";
const STYLE = "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1";
const COLORS = "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2";
const encoder = new TextEncoder();
const invalid = message => { throw new TypeError(message); };
function record(value, keys) {
  if (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) invalid("Expected a plain research record.");
  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
    if (!keys.includes(key) || !("value" in descriptor)) invalid(`Unsupported field/accessor ${key}.`);
  }
}
function xml(value) {
  if (typeof value !== "string") invalid("Text must be a string.");
  for (const char of value) {
    const cp = char.codePointAt(0);
    if (!(cp === 9 || cp === 10 || cp === 13 || cp >= 32 && cp <= 0xd7ff || cp >= 0xe000 && cp <= 0xfffd || cp >= 0x10000 && cp <= 0x10ffff)) invalid("Text contains an XML 1.0-invalid scalar.");
  }
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&apos;").replaceAll("\r", "&#13;");
}
function id(value) { if (typeof value !== "string" || !value.length || value.length > 256) invalid("Candidate IDs must be nonempty strings of at most 256 code units."); return value; }

export function nativeDiagramSpecimen(candidate, options = {}) {
  record(candidate, ["family", "roots", "nodes", "closure"]);
  record(options, ["cache", "presentation", "modelIds", "identityState", "drawingId", "frame", "limits"]);
  const profile = diagramResearchProfiles.find(p => p.family === candidate.family);
  if (!profile) invalid("Unknown research family.");
  if (candidate.family === "cycle" ? candidate.closure !== "layout-generated" : candidate.closure !== undefined) invalid("Invalid family closure policy.");
  const cache = options.cache ?? "absent";
  if (!["empty", "absent"].includes(cache)) invalid("Only explicit empty or absent cache specimens are available.");
  const presentation = options.presentation ?? "content-only";
  if (!["content-only", "layout"].includes(presentation)) invalid("Unknown presentation research policy.");
  const modelIds = options.modelIds ?? "int32";
  if (!["int32", "guid"].includes(modelIds)) invalid("Unknown model identity research policy.");
  if (options.identityState !== undefined && modelIds !== "guid") invalid("Persistent identity state requires explicit GUID policy.");
  // Deliberately scoped specimen identities, not persistent managed author IDs.
  // A named PowerPoint Cycle control retains its content GUIDs on first edit/
  // save, whereas integer specimens receive new GUIDs. Other gates stay open.
  let resolvedIds;
  const modelId = serial => resolvedIds ? resolvedIds.get(serial) : modelIds === "int32" ? String(serial) : `{00000000-0000-4000-8000-${serial.toString(16).toUpperCase().padStart(12, "0")}}`;
  const limits = { maxNodes: 10000, maxDepth: 256, maxWork: 200000, maxTextBytes: 1000000, maxOutputBytes: 4000000, ...options.limits };
  if (options.limits !== undefined) record(options.limits, Object.keys(limits).filter(k => ["maxNodes", "maxDepth", "maxWork", "maxTextBytes", "maxOutputBytes"].includes(k)));
  for (const [key, number] of Object.entries(limits)) if (!["maxNodes", "maxDepth", "maxWork", "maxTextBytes", "maxOutputBytes"].includes(key) || !Number.isSafeInteger(number) || number < 1) invalid(`Invalid ${key}.`);
  const drawingId = options.drawingId ?? 3;
  if (!Number.isSafeInteger(drawingId) || drawingId < 1 || drawingId > 4294967295) invalid("Invalid drawing identity.");
  const frame = options.frame ?? { x: 1270000, y: 1270000, cx: 6604000, cy: 3810000 };
  record(frame, ["x", "y", "cx", "cy"]);
  for (const key of ["x", "y", "cx", "cy"]) if (!Number.isSafeInteger(frame[key]) || frame[key] > 2147483647 || frame[key] < (key === "x" || key === "y" ? -2147483648 : 1)) invalid(`Invalid frame ${key}.`);
  if (!Array.isArray(candidate.nodes) || !Array.isArray(candidate.roots) || candidate.nodes.length > Math.min(limits.maxNodes, 536870911)) invalid("Invalid/over-budget node or root collection.");
  let work = 0, textBytes = 0;
  const step = () => { if (++work > limits.maxWork) throw new RangeError("Native diagram specimen exceeds maxWork."); };
  const nodes = new Map(), parents = new Map();
  for (const node of candidate.nodes) {
    step(); record(node, ["id", "role", "paragraphs", "children"]);
    const key = id(node.id); if (nodes.has(key)) invalid("Duplicate candidate node ID.");
    if (!["content", "assistant"].includes(node.role) || node.role === "assistant" && candidate.family !== "hierarchy") invalid("Unadmitted node role.");
    if (!Array.isArray(node.children) || !Array.isArray(node.paragraphs) || !node.paragraphs.length) invalid("Each node needs children and at least one plain paragraph.");
    const paragraphs = [];
    for (const paragraph of node.paragraphs) {
      step(); if (typeof paragraph !== "string") invalid("Rich text/fields are outside the native specimen profile.");
      textBytes += encoder.encode(paragraph).byteLength;
      if (textBytes > limits.maxTextBytes) throw new RangeError("Native diagram specimen exceeds maxTextBytes.");
      paragraphs.push(paragraph);
    }
    nodes.set(key, { id: key, role: node.role, paragraphs, children: [] });
    for (const child of node.children) { step(); id(child); if (parents.has(child)) invalid("Multiple parent/duplicate child references."); parents.set(child, key); nodes.get(key).children.push(child); }
  }
  const roots = candidate.roots.map(root => {
    step(); id(root); if (parents.has(root)) invalid("Repeated root or root with parent."); parents.set(root, null); return root;
  });
  for (const [child, parent] of parents) { step(); if (!nodes.has(child) || parent !== null && !nodes.has(parent)) invalid("Missing child/root identity."); }
  const order = [], seen = new Set(), stack = roots.slice().reverse().map(key => ({ key, depth: 1 }));
  while (stack.length) {
    step(); const { key, depth } = stack.pop(), node = nodes.get(key);
    if (seen.has(key)) invalid("Cycle/repeated traversal."); seen.add(key);
    if (depth > limits.maxDepth) throw new RangeError("Native diagram specimen exceeds maxDepth.");
    if (candidate.family === "cycle" && node.children.length) invalid("Nested Cycle nodes remain unadmitted.");
    if (node.role === "assistant" && (parents.get(key) === null || node.children.length)) invalid("Assistant roots/descendants remain unadmitted.");
    order.push(node); for (let i = node.children.length - 1; i >= 0; i--) stack.push({ key: node.children[i], depth: depth + 1 });
  }
  if (seen.size !== nodes.size) invalid("Disconnected nodes/cycles are not repaired.");
  const nativeIds = new Map(order.map((node, i) => [node.id, i + 1])), count = order.length;
  const edges = [];
  for (const [parent, children] of [[null, roots], ...order.map(n => [n.id, n.children])]) for (const [index, child] of children.entries()) {
    step(); const serial = edges.length;
    edges.push({ source: parent === null ? 0 : nativeIds.get(parent), destination: nativeIds.get(child), order: index,
      id: count + serial + 1, parentTransition: 2 * count + serial + 1, siblingTransition: 3 * count + serial + 1 });
  }
  let nextId = 4 * count;
  const allocate = () => { step(); if (++nextId > 2147483647) throw new RangeError("Model identity exceeds int32."); return nextId; };
  const displays = presentation === "layout" ? diagramPresentationSpecimen(candidate.family, roots, order, nativeIds, edges, allocate, step) : { points: [], connections: [] };
  let identityState;
  if (options.identityState !== undefined) {
    // Semantic keys exclude traversal ordinals, sibling order, text and style.
    // Parent/transition identity follows the uniquely parented content child;
    // presentation identity follows source ownership plus exact layout role.
    const keys = new Map([[0, ["document"]]]);
    for (const node of order) { step(); keys.set(nativeIds.get(node.id), ["content", node.id]); }
    for (const edge of edges) {
      step(); const child = keys.get(edge.destination)[1];
      keys.set(edge.id, ["parent-edge", child]);
      keys.set(edge.parentTransition, ["parent-transition", child]);
      keys.set(edge.siblingTransition, ["sibling-transition", child]);
    }
    for (const point of displays.points) { step(); keys.set(point.id, ["presentation", ...keys.get(point.source), point.name]); }
    for (const edge of displays.connections) { step(); keys.set(edge.id, ["presentation-edge", edge.type, JSON.stringify(keys.get(edge.source)), JSON.stringify(keys.get(edge.destination))]); }
    step();
    const active = [...keys], reconciled = reconcileDiagramIdentities(options.identityState, candidate.family, active.map(([, key]) => JSON.stringify(key)), { maxWork: Math.max(1, limits.maxWork - work), maxKeyBytes: limits.maxOutputBytes });
    work += reconciled.work;
    resolvedIds = new Map(active.map(([serial], index) => [serial, reconciled.identities[index].modelId]));
    identityState = reconciled.state;
  }
  let outputBytes = 0;
  const append = (chunks, value) => {
    step(); outputBytes += encoder.encode(value).byteLength;
    if (outputBytes > limits.maxOutputBytes) throw new RangeError("Native diagram specimen exceeds maxOutputBytes.");
    chunks.push(value);
  };
  const body = paragraphs => `<dgm:t><a:bodyPr/><a:lstStyle/>${paragraphs.map(paragraph => `<a:p>${paragraph.split("\n").map((line, i) => `${i ? "<a:br/>" : ""}<a:r><a:t xml:space="preserve">${xml(line)}</a:t></a:r>`).join("")}</a:p>`).join("")}</dgm:t>`;
  const parts = [];
  append(parts, `<dgm:dataModel xmlns:dgm="${D}" xmlns:a="${A}" xmlns:r="${R}" xmlns:dsp="${DSP}"><dgm:ptLst>`);
  append(parts, `<dgm:pt modelId="${modelId(0)}" type="doc"><dgm:prSet loTypeId="${profile.layoutIdentifier}" loCatId="${candidate.family}" qsTypeId="${STYLE}" qsCatId="simple" csTypeId="${COLORS}" csCatId="accent1"/><dgm:spPr/>${body([""])}</dgm:pt>`);
  // Plain authored text has no custom formatting to lock. custT would let the
  // native layout ignore automatic text formatting (including fitting).
  for (const node of order) append(parts, `<dgm:pt modelId="${modelId(nativeIds.get(node.id))}"${node.role === "assistant" ? ' type="asst"' : ""}><dgm:prSet phldr="0"/><dgm:spPr/>${body(node.paragraphs)}</dgm:pt>`);
  for (const edge of edges) for (const [kind, pointId] of [["parTrans", edge.parentTransition], ["sibTrans", edge.siblingTransition]]) append(parts, `<dgm:pt modelId="${modelId(pointId)}" type="${kind}" cxnId="${modelId(edge.id)}"><dgm:prSet/><dgm:spPr/>${body([""])}</dgm:pt>`);
  for (const point of displays.points) append(parts, `<dgm:pt modelId="${modelId(point.id)}" type="pres"><dgm:prSet presAssocID="${modelId(point.source)}" presName="${point.name}"${point.label === undefined ? "" : ` presStyleLbl="${point.label}" presStyleIdx="${point.index}"`} presStyleCnt="${point.count}"/><dgm:spPr/>${body([""])}</dgm:pt>`);
  append(parts, "</dgm:ptLst><dgm:cxnLst>");
  for (const edge of edges) append(parts, `<dgm:cxn modelId="${modelId(edge.id)}" type="parOf" srcId="${modelId(edge.source)}" destId="${modelId(edge.destination)}" srcOrd="${edge.order}" destOrd="0" parTransId="${modelId(edge.parentTransition)}" sibTransId="${modelId(edge.siblingTransition)}"/>`);
  for (const edge of displays.connections) append(parts, `<dgm:cxn modelId="${modelId(edge.id)}" type="${edge.type}" srcId="${modelId(edge.source)}" destId="${modelId(edge.destination)}" srcOrd="${edge.order}" destOrd="${edge.targetOrder}"/>`);
  append(parts, `</dgm:cxnLst><dgm:bg/><dgm:whole/>${cache === "empty" ? `<dgm:extLst><a:ext uri="${DSP}"><dsp:dataModelExt relId="pptkitDiagramDrawing"/></a:ext></dgm:extLst>` : ""}</dgm:dataModel>`);
  const frameParts = [];
  append(frameParts, `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${drawingId}" name="PPTKit native diagram specimen"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${frame.x}" y="${frame.y}"/><a:ext cx="${frame.cx}" cy="${frame.cy}"/></p:xfrm><a:graphic><a:graphicData uri="${D}"><dgm:relIds xmlns:dgm="${D}" r:dm="pptkitDiagramData" r:lo="pptkitDiagramLayout" r:cs="pptkitDiagramColors" r:qs="pptkitDiagramStyle"/></a:graphicData></a:graphic></p:graphicFrame>`);
  const cacheParts = [];
  if (cache === "empty") append(cacheParts, `<dsp:drawing xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${frame.cx}" cy="${frame.cy}"/><a:chOff x="0" y="0"/><a:chExt cx="${frame.cx}" cy="${frame.cy}"/></a:xfrm></dsp:grpSpPr></dsp:spTree></dsp:drawing>`);
  return { dataXml: parts.join(""), frameXml: frameParts.join(""), ...(cache === "empty" ? { cacheXml: cacheParts.join("") } : {}),
    family: profile.family, layoutIdentifier: profile.layoutIdentifier, cachePolicy: cache, presentationPolicy: presentation, modelIdPolicy: modelIds,
    ...(identityState === undefined ? {} : { identityState }),
    identities: order.map(node => ({ id: node.id, modelId: modelId(nativeIds.get(node.id)) })),
    nativeAdmission: "untested" };
}
