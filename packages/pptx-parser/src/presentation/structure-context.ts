import { readOpcSource } from "../opc/read-package.js";
import { readMetadata } from "../metadata/read-metadata.js";
import { identifyPresentation, ROOT_NAMESPACES } from "./inspect.js";
import { captureConfiguration, expandedKey, MCE } from "../compatibility/configuration.js";
import { processCompatibility } from "../compatibility/process.js";
import type { OpcRelationshipInfo } from "../types/opc.js";
import type { MarkupCompatibilityReport } from "../types/compatibility.js";
import type { XmlElement } from "../xml/read-xml.js";
import { invalidStructure } from "./structure-xml.js";
import { PackageInspectionError } from "../types/archive.js";
import { isSectionExtension, SECTION_NS } from "./sections.js";
import { captureDiagramLimits, captureDiagramTextComparison } from "./diagram-limits.js";
import type { PptxInspectionOptions } from "../types/structure.js";

const CONTENT_PREFIX = "application/vnd.openxmlformats-officedocument.";
export async function structureContext(bytes: Uint8Array, options?: PptxInspectionOptions, retainPayloads = false) {
  const diagramLimits = captureDiagramLimits(options?.diagramLimits);
  const diagramTextComparison = captureDiagramTextComparison(options?.diagramTextComparison);
  const signal = options?.signal;
  const source = await readOpcSource(bytes, options, true, retainPayloads), entry = identifyPresentation(source.inspection, false);
  const p = ROOT_NAMESPACES[entry.dialect];
  const r = entry.dialect === "strict" ? "http://purl.oclc.org/ooxml/officeDocument/relationships" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const drawing = entry.dialect === "strict" ? "http://purl.oclc.org/ooxml/drawingml" : "http://schemas.openxmlformats.org/drawingml/2006";
  const a = `${drawing}/main`;
  const configuration = captureConfiguration({ understoodNamespaces: [p, r, a, ...["chart", "diagram", "picture", "lockedCanvas", "compatibility"].map((name) => `${drawing}/${name}`)], extensionElements: [p, a, `${drawing}/chart`].map(namespaceUri => ({ namespaceUri, local: "ext" })) });
  const sectionConfiguration = { understood: new Set([...configuration.understood, SECTION_NS]), extensions: new Set([...configuration.extensions, expandedKey("http://schemas.openxmlformats.org/presentationml/2006/main", "ext")]) };
  const parts = new Map(source.inspection.parts.map((part) => [part.name, part]));
  const bySource = new Map<string, Map<string, OpcRelationshipInfo>>();
  for (const relation of source.inspection.relationships) { let map = bySource.get(relation.source); if (!map) { map = new Map(); bySource.set(relation.source, map); } map.set(relation.id, relation); }
  const reports: { part: string; report: MarkupCompatibilityReport }[] = [], effective = new Map<string, Promise<XmlElement>>();
  const origins = new Map<string, WeakMap<XmlElement, XmlElement>>(), rawIndices = new Map<string, Promise<WeakMap<XmlElement, { elementIndex: number; parent?: XmlElement }>>>();
  const abort = (): void => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "Presentation structure inspection was aborted."); };
  let work = 0;
  const checkpoint = async (): Promise<void> => { abort(); if (++work % 64 === 0) { await new Promise<void>((resolve) => setTimeout(resolve, 0)); abort(); } };
  const root = async (part: string, local: string, type?: string, namespaceUri = p): Promise<XmlElement> => {
    await checkpoint();
    const inventory = parts.get(part);
    if (!inventory || (type && inventory.contentType.toLowerCase() !== `${CONTENT_PREFIX}${type}+xml`.toLowerCase())) invalidStructure(`Unexpected content type for ${local} target.`, part);
    let pending = effective.get(part);
    if (!pending) {
      const original = source.xml.get(part);
      if (!original) invalidStructure(`Expected XML for ${local} part.`, part);
      pending = (async () => {
        if (!inventory.xml?.namespaceUris.includes(MCE)) return original!;
        const main = part === entry.mainPart;
        const projected = await processCompatibility(original!, main ? sectionConfiguration : configuration, part, signal, main ? (node, parent, grandparent) => grandparent?.namespaceUri === p && grandparent.local === "presentation" && parent?.namespaceUri === p && parent.local === "extLst" && isSectionExtension(node, p) : undefined);
        const roots = projected.nodes.filter((node): node is XmlElement => node.kind === "element");
        if (roots.length !== 1 || projected.nodes.some((node) => (node.kind === "text" || node.kind === "cdata") && !/^[\t\r\n ]*$/.test(node.value))) invalidStructure("Presentation part must project to one root.", part);
        origins.set(part, projected.origins); reports.push({ part, report: projected.report });
        return roots[0]!;
      })();
      effective.set(part, pending);
    }
    let tree: XmlElement;
    try { tree = await pending; } catch (error) { if (effective.get(part) === pending) effective.delete(part); throw error; }
    if (tree.local !== local || tree.namespaceUri !== namespaceUri) invalidStructure(`Expected ${local} root in the presentation dialect.`, part);
    return tree;
  };
  const relationships = (part: string, type: string): OpcRelationshipInfo[] => [...(bySource.get(part)?.values() ?? [])].filter((relation) => relation.type === `${r}/${type}`);
  const target = (relation: OpcRelationshipInfo, type: string): string => {
    if (relation.type !== `${r}/${type}` || relation.targetMode !== "Internal" || relation.fragment !== undefined || !relation.targetPart) invalidStructure(`Expected an internal whole-part ${type} relationship.`, relation.source);
    return relation.targetPart!;
  };
  const findRelationship = (part: string, id: string | undefined): OpcRelationshipInfo | undefined => id === undefined ? undefined : bySource.get(part)?.get(id);
  const relationship = (part: string, id: string | undefined, type: string): OpcRelationshipInfo => {
    const relation = findRelationship(part, id);
    if (!relation) invalidStructure(`Missing ${type} relationship for r:id ${id ?? "(absent)"}.`, part);
    if (relation!.type !== `${r}/${type}`) invalidStructure(`Unexpected relationship type for ${type} reference.`, part);
    return relation!;
  };
  const reference = (part: string, id: string | undefined, type: string): { part: string; relationshipId: string } => {
    const relation = relationship(part, id, type);
    return { part: target(relation!, type), relationshipId: relation!.id };
  };
  // Only content import retains verified archive payloads. No second inflate,
  // public byte getter, filesystem access or external relationship resolution.
  const payload = (part: string) => {
    const inventory = parts.get(part), bytes = inventory ? source.payloads?.get(inventory.entryName) : undefined;
    if (!inventory || !bytes) invalidStructure("Content payload is unavailable.", part);
    return { bytes: bytes!, contentType: inventory!.contentType };
  };
  const optionalTarget = (part: string, type: string): string | undefined => {
    const matches = relationships(part, type);
    if (matches.length > 1) invalidStructure(`Multiple ${type} relationships where at most one is permitted.`, part);
    return matches[0] ? target(matches[0], type) : undefined;
  };
  const sourceElement = async (part: string, node: XmlElement): Promise<{ node: XmlElement; elementIndex: number; parent?: XmlElement }> => {
    let pending = rawIndices.get(part);
    if (!pending) {
      const root = source.xml.get(part);
      if (!root) invalidStructure("Raw source XML is unavailable.", part);
      pending = (async () => {
        const indices = new WeakMap<XmlElement, { elementIndex: number; parent?: XmlElement }>(), stack: { node: XmlElement; parent?: XmlElement }[] = [{ node: root! }]; let index = 0;
        while (stack.length) {
          await checkpoint();
          const item = stack.pop()!; indices.set(item.node, { elementIndex: ++index, ...(item.parent ? { parent: item.parent } : {}) });
          for (let i = item.node.children.length - 1; i >= 0; i--) {
            const child = item.node.children[i]!; if (child.kind === "element") stack.push({ node: child, parent: item.node });
          }
        }
        return indices;
      })();
      rawIndices.set(part, pending);
    }
    let indices: WeakMap<XmlElement, { elementIndex: number; parent?: XmlElement }>;
    try { indices = await pending; } catch (error) { if (rawIndices.get(part) === pending) rawIndices.delete(part); throw error; }
    const raw = origins.get(part)?.get(node) ?? node, identity = indices.get(raw);
    if (!identity) invalidStructure("Effective XML node has no verified source identity.", part);
    return { node: raw, ...identity! };
  };
  const inspected = await readMetadata(source, entry.dialect === "strict", checkpoint, signal);
  reports.push(...inspected.reports);
  return { entry, p, r, a, root, sourceElement, rawXml: source.xml, reference, relationship, findRelationship, payload, optionalTarget, relationships, target, reports, checkpoint, abort, metadata: inspected.metadata, customMetadataRoot: inspected.customRoot, diagramLimits, diagramTextComparison };
}
