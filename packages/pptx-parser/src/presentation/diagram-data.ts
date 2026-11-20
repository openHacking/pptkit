import type { structureContext } from "./structure-context.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { PptxDiagramDataModelInfo, PptxDiagramModelFinding, PptxDiagramModelReference, PptxDiagramPointInfo, PptxDiagramConnectionInfo, PptxDiagramTextParagraph } from "../types/smartart.js";
import { attribute, children } from "./structure-xml.js";
import { diagramModelIdentity as identity } from "./diagram-model-id.js";

type Context = Awaited<ReturnType<typeof structureContext>>;
const trim = (value: string) => value.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "");
const hasOpaqueCharacters = (node: XmlElement) => node.children.some(child => (child.kind === "text" || child.kind === "cdata") && /[^\t\r\n ]/.test(child.value));
function scalar(value: string | undefined, min: number, max: number): number | undefined {
  if (value === undefined || !/^[+-]?\d+$/.test(trim(value))) return undefined;
  const number = Number(value); return Number.isSafeInteger(number) && number >= min && number <= max ? number : undefined;
}

/** One model per located data part. No Core tree, layout evaluation or source-edit authority. */
export async function readDiagramData(c: Context, part: string, containsMce: boolean, checkpoint: () => Promise<void>, chargeFinding: () => void): Promise<PptxDiagramDataModelInfo> {
  const root = c.rawXml.get(part)!, dgm = c.a.replace(/\/main$/, "/diagram");
  const findings: PptxDiagramModelFinding[] = [];
  const note = async (node: XmlElement, code: PptxDiagramModelFinding["code"], reason: string, attr?: string) => {
    await checkpoint(); chargeFinding();
    findings.push({ code, elementIndex: (await c.sourceElement(part, node)).elementIndex, ...(attr ? { attribute: attr } : {}), reason });
  };
  const empty = async (reason: string): Promise<PptxDiagramDataModelInfo> => {
    await note(root, "source-only", reason);
    return { part, coverage: "identities-references-direct-text", status: "source-only", points: [], connections: [], findings };
  };
  if (containsMce) return empty("Data-model MCE needs a separate semantic branch policy; no graph is inferred from raw branches.");
  const lists = (name: string) => children(root).filter(n => n.namespaceUri === dgm && n.local === name);
  const pointLists = lists("ptLst"), connectionLists = lists("cxnLst");
  if (pointLists.length !== 1 || connectionLists.length > 1) return empty("Expected one direct point list and at most one direct connection list; ambiguous lists are not merged.");
  const pointNodes = children(pointLists[0]!).filter(n => n.namespaceUri === dgm && n.local === "pt");
  const connectionNodes = connectionLists[0] ? children(connectionLists[0]).filter(n => n.namespaceUri === dgm && n.local === "cxn") : [];
  for (const list of [...pointLists, ...connectionLists]) for (const node of children(list)) {
    await checkpoint();
    if (node.namespaceUri !== dgm || node.local !== (list.local === "ptLst" ? "pt" : "cxn")) await note(node, "uninterpreted-structure", "List child is outside the point/connection source profile.");
  }
  const entities = new Map<string, { kind: "point" | "connection"; node: XmlElement; count: number }>();
  for (const [kind, nodes] of [["point", pointNodes], ["connection", connectionNodes]] as const) for (const node of nodes) {
    await checkpoint();
    const key = identity(attribute(node, "modelId")); if (key === undefined) continue;
    const found = entities.get(key); if (found) found.count++; else entities.set(key, { kind, node, count: 1 });
  }
  const ownIdentity = async (node: XmlElement): Promise<PptxDiagramPointInfo["identityStatus"]> => {
    const value = attribute(node, "modelId"), key = identity(value);
    const status = value === undefined ? "missing" : key === undefined ? "unrecognized" : entities.get(key)!.count > 1 ? "duplicate" : "unique";
    if (status !== "unique") await note(node, status === "duplicate" ? "duplicate-identity" : status === "missing" ? "missing-field" : "unrecognized-field", `Model identity is ${status} within the int32/uppercase-braced-GUID profile.`, "modelId");
    return status;
  };
  const reference = async (node: XmlElement, attr: string, kind: "point" | "connection", required = false): Promise<PptxDiagramModelReference> => {
    const value = attribute(node, attr), key = identity(value), found = key === undefined ? undefined : entities.get(key);
    const status = value === undefined ? "absent" : key === undefined ? "unrecognized" : !found ? "missing" : found.count > 1 ? "ambiguous" : found.kind !== kind ? "wrong-kind" : "resolved";
    if (status !== "resolved" && (status !== "absent" || required)) await note(node, "unresolved-reference", `${attr} is ${status}; expected a unique ${kind} identity.`, attr);
    return { ...(value !== undefined ? { value } : {}), status, ...(status === "resolved" ? { targetElementIndex: (await c.sourceElement(part, found!.node)).elementIndex } : {}) };
  };
  const one = async (node: XmlElement, ns: string, local: string): Promise<XmlElement | undefined> => {
    const found = children(node).filter(n => n.namespaceUri === ns && n.local === local);
    if (found.length > 1) { await note(node, "uninterpreted-structure", `Multiple ${local} children are ambiguous; none is selected.`); return undefined; }
    return found[0];
  };
  const text = async (body: XmlElement): Promise<PptxDiagramTextParagraph[]> => {
    const result: PptxDiagramTextParagraph[] = [];
    if (hasOpaqueCharacters(body)) { await note(body, "uninterpreted-structure", "Text body contains direct characters outside paragraphs; text is not inferred."); return result; }
    for (const para of children(body)) {
      await checkpoint();
      if (para.namespaceUri !== c.a || para.local !== "p") { if (!(para.namespaceUri === c.a && ["bodyPr", "lstStyle"].includes(para.local))) await note(para, "uninterpreted-structure", "Text-body child is outside direct paragraph text."); continue; }
      if (hasOpaqueCharacters(para)) { await note(para, "uninterpreted-structure", "Paragraph contains direct characters outside runs; text is not inferred."); continue; }
      const pPr = await one(para, c.a, "pPr"), level = pPr ? attribute(pPr, "lvl") : undefined;
      if (level !== undefined && scalar(level, 0, 8) === undefined) await note(pPr!, "unrecognized-field", "Paragraph level is outside the source profile's range 0–8.", "lvl");
      const runs: { elementIndex: number; kind: "text" | "break" | "field"; text: string; fieldId?: string; fieldType?: string }[] = [];
      for (const run of children(para)) {
        await checkpoint();
        if (run.namespaceUri === c.a && ["pPr", "endParaRPr"].includes(run.local)) continue;
        if (run.namespaceUri !== c.a || !["r", "br", "fld"].includes(run.local)) { await note(run, "uninterpreted-structure", "Paragraph child is outside direct run/break/field text."); continue; }
        if (hasOpaqueCharacters(run)) { await note(run, "uninterpreted-structure", "Run contains direct characters outside its text leaf; text is not inferred."); continue; }
        const leaf = run.local === "br" ? undefined : await one(run, c.a, "t");
        if (run.local !== "br" && (!leaf || children(leaf).length)) { await note(run, "uninterpreted-structure", "Text run/field lacks one scalar text leaf; text is not guessed."); continue; }
        const id = run.local === "fld" ? attribute(run, "id") : undefined, type = run.local === "fld" ? attribute(run, "type") : undefined;
        runs.push({ elementIndex: (await c.sourceElement(part, run)).elementIndex, kind: run.local === "br" ? "break" : run.local === "fld" ? "field" : "text",
          text: run.local === "br" ? "\n" : leaf!.children.map(n => n.kind === "text" || n.kind === "cdata" ? n.value : "").join(""), ...(id !== undefined ? { fieldId: id } : {}), ...(type !== undefined ? { fieldType: type } : {}) });
      }
      result.push({ elementIndex: (await c.sourceElement(part, para)).elementIndex, ...(level !== undefined ? { level } : {}), runs });
    }
    return result;
  };
  const points: PptxDiagramPointInfo[] = [], connections: PptxDiagramConnectionInfo[] = [];
  const pointKinds = new Set(["node", "asst", "doc", "pres", "parTrans", "sibTrans"]);
  for (const node of pointNodes) {
    await checkpoint();
    const modelId = attribute(node, "modelId"), type = attribute(node, "type"), kind = type === undefined ? "node" : trim(type);
    if (!pointKinds.has(kind)) await note(node, "unrecognized-field", "Point type is outside the recognized source kinds.", "type");
    const props = await one(node, dgm, "prSet"), body = await one(node, dgm, "t");
    points.push({ elementIndex: (await c.sourceElement(part, node)).elementIndex, ...(modelId !== undefined ? { modelId } : {}), identityStatus: await ownIdentity(node),
      ...(type !== undefined ? { type } : {}), kind: pointKinds.has(kind) ? kind as PptxDiagramPointInfo["kind"] : "unrecognized",
      ...(attribute(node, "cxnId") !== undefined ? { connection: kind === "parTrans" || kind === "sibTrans" ? await reference(node, "cxnId", "connection") : { status: "uninterpreted" as const, value: attribute(node, "cxnId")! } } : {}),
      ...(props ? { properties: props.attributes.filter(a => a.namespaceUri !== "http://www.w3.org/2000/xmlns/").map(a => ({ namespaceUri: a.namespaceUri, local: a.local, value: a.value })),
        ...(attribute(props, "presAssocID") !== undefined ? { presentationAssociation: await reference(props, "presAssocID", "point") } : {}) } : {}),
      ...(body ? { text: await text(body) } : {}) });
    for (const child of children(node)) if (child.namespaceUri !== dgm || !["prSet", "t", "spPr", "extLst"].includes(child.local)) await note(child, "uninterpreted-structure", "Point child is outside the source graph profile.");
  }
  const connectionKinds = new Set(["parOf", "presOf", "presParOf", "unknownRelationship"]);
  for (const node of connectionNodes) {
    await checkpoint();
    const modelId = attribute(node, "modelId"), type = attribute(node, "type"), kind = type === undefined ? "parOf" : trim(type);
    if (!connectionKinds.has(kind)) await note(node, "unrecognized-field", "Connection type is outside the recognized source kinds.", "type");
    const sourceOrder = attribute(node, "srcOrd"), destinationOrder = attribute(node, "destOrd"), sourceOrderValue = scalar(sourceOrder, 0, 4294967295), destinationOrderValue = scalar(destinationOrder, 0, 4294967295);
    for (const [attr, value, parsed] of [["srcOrd", sourceOrder, sourceOrderValue], ["destOrd", destinationOrder, destinationOrderValue]] as const) if (parsed === undefined) await note(node, value === undefined ? "missing-field" : "unrecognized-field", "Order must be an explicit unsigned 32-bit source integer.", attr);
    const presentationId = attribute(node, "presId");
    connections.push({ elementIndex: (await c.sourceElement(part, node)).elementIndex, ...(modelId !== undefined ? { modelId } : {}), identityStatus: await ownIdentity(node),
      ...(type !== undefined ? { type } : {}), kind: connectionKinds.has(kind) ? kind as PptxDiagramConnectionInfo["kind"] : "unrecognized",
      source: await reference(node, "srcId", "point", true), destination: await reference(node, "destId", "point", true),
      ...(attribute(node, "parTransId") !== undefined ? { parentTransition: await reference(node, "parTransId", "point") } : {}),
      ...(attribute(node, "sibTransId") !== undefined ? { siblingTransition: await reference(node, "sibTransId", "point") } : {}),
      ...(sourceOrder !== undefined ? { sourceOrder } : {}), ...(destinationOrder !== undefined ? { destinationOrder } : {}),
      ...(sourceOrderValue !== undefined ? { sourceOrderValue } : {}), ...(destinationOrderValue !== undefined ? { destinationOrderValue } : {}), ...(presentationId !== undefined ? { presentationId } : {}) });
  }
  return { part, coverage: "identities-references-direct-text", status: findings.length ? "partial" : "inspected", points, connections, findings };
}
