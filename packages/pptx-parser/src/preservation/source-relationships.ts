import { MCE, XMLNS } from "../compatibility/configuration.js";
import { asciiKey, iriReference, relationshipSource, resolveInternal } from "../opc/names.js";
import { RELATIONSHIPS_NS, XML_ID } from "../opc/markup.js";
import { PackageInspectionError } from "../types/archive.js";
import type { OpcPackageInspection, OpcPartInfo } from "../types/opc.js";
import type { OpcClosureFinding, OpcSourceRelationship } from "../types/closure.js";
import type { XmlDocument, XmlElement, XmlNode } from "../xml/read-xml.js";

const directives = new Set(["Ignorable", "MustUnderstand", "ProcessContent", "PreserveElements", "PreserveAttributes"]);
const fields = new Set(["Id", "Type", "Target", "TargetMode"]);
interface BranchCondition { readonly alternate: number; readonly branch: number; readonly parent: BranchCondition | undefined }
const order = (a: string, b: string): number => a < b ? -1 : a > b ? 1 : 0;
const key = (source: string, id: string | undefined, type: string | undefined, target: string | undefined, mode: string | undefined, targetPart: string | undefined, fragment: string | undefined) =>
  JSON.stringify([source, id ?? null, type ?? null, target ?? null, mode ?? null, targetPart ?? null, fragment ?? null]);

/** Source union, not a second MCE projection. Inactive references cannot disappear. */
export async function sourceRelationships(
  inspection: OpcPackageInspection,
  documents: ReadonlyMap<string, XmlDocument>,
  work: (units: number) => Promise<void>,
  addFinding: (finding: OpcClosureFinding) => void,
  reserveRelationship: () => void,
): Promise<{ edges: OpcSourceRelationship[]; sets: Map<string, string> }> {
  const parts = new Map(inspection.parts.map((part) => [asciiKey(part.name), part]));
  const effective = new Set<string>();
  for (const edge of inspection.relationships) {
    await work(edge.source.length + edge.id.length + 1);
    effective.add(key(edge.source, edge.id, edge.type, edge.target, edge.targetMode, edge.targetPart, edge.fragment));
  }
  const edges: OpcSourceRelationship[] = [], sets = new Map<string, string>();
  for (const part of inspection.parts.filter((part) => part.relationshipPart).sort((a, b) => order(a.name, b.name))) {
    const declaredSource = relationshipSource(part.name)!;
    const source = declaredSource === "/" ? "/" : parts.get(asciiKey(declaredSource))!.name;
    sets.set(source, part.name);
    const document = documents.get(part.name)!;
    const identities = new Map<string, (BranchCondition | undefined)[]>();
    const stack = [{ nodes: document.children as readonly XmlNode[], index: 0, withinSet: false, withinRelationship: false, parentIndex: 0, alternateIndex: 0, condition: undefined as BranchCondition | undefined }];
    let elementIndex = 0;
    while (stack.length) {
      const frame = stack[stack.length - 1]!;
      if (frame.index === frame.nodes.length) { stack.pop(); continue; }
      const node = frame.nodes[frame.index++]!;
      await work(1);
      if (node.kind !== "element") {
        if ((node.kind === "text" || node.kind === "cdata") && !frame.withinRelationship && !/^[\t\r\n ]*$/.test(node.value)) {
          addFinding({ source, relationshipPart: part.name, elementIndex: frame.parentIndex, code: "opaque-source-markup", reason: "Source relationship control character content is uninterpreted." });
        }
        continue;
      }
      const index = ++elementIndex, rel = node.namespaceUri === RELATIONSHIPS_NS && node.local === "Relationship";
      const set = node.namespaceUri === RELATIONSHIPS_NS && node.local === "Relationships";
      const mce = node.namespaceUri === MCE && ["AlternateContent", "Choice", "Fallback"].includes(node.local);
      const location = { source, relationshipPart: part.name, elementIndex: index };
      const opaque = !rel && !set && !mce || !knownAttributes(node);
      if (opaque) addFinding({ ...location, code: "opaque-source-markup", reason: "Source relationship markup contains an unrecognized element or attribute; hidden dependencies cannot be excluded." });
      if ((set && (frame.withinSet || frame.withinRelationship)) || (rel && (!frame.withinSet || frame.withinRelationship))) {
        addFinding({ ...location, code: "invalid-source-relationship", reason: "Source relationship markup has unsupported placement." });
      }
      if (rel) {
        reserveRelationship();
        const attrs = new Map(node.attributes.filter((a) => !a.namespaceUri).map((a) => [a.local, a.value]));
        await work([...attrs.values()].reduce((sum, value) => sum + value.length, 0) + 1);
        const rawId = attrs.get("Id"), id = rawId?.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "").replace(/[\t\r\n ]+/g, " ");
        const type = attrs.get("Type"), target = attrs.get("Target"), rawMode = attrs.get("TargetMode") ?? "Internal";
        const targetMode = rawMode === "Internal" || rawMode === "External" ? rawMode : undefined;
        let targetPart: string | undefined, fragment: string | undefined;
        let invalid = !id || !XML_ID.test(id) || !type || !target || !targetMode;
        if (id && XML_ID.test(id)) {
          let conditions = identities.get(id);
          if (!conditions) { conditions = []; identities.set(id, conditions); }
          for (const previous of conditions) {
            if (await mayCoexist(previous, frame.condition, work)) {
              addFinding({ ...location, code: "ambiguous-source-identity", reason: "Repeated source relationship IDs may coexist in one branch combination; safe identity reconciliation is unproven." });
              break;
            }
          }
          if (!conditions.includes(frame.condition)) conditions.push(frame.condition);
        }
        if (type) {
          try { iriReference(type, part.entryName); }
          catch (error) { if (!(error instanceof PackageInspectionError)) throw error; invalid = true; }
        }
        if (target && targetMode) {
          try {
            iriReference(target, part.entryName);
            if (targetMode === "Internal") {
              const resolved = resolveInternal(source, target, part.entryName), destination: OpcPartInfo | undefined = parts.get(asciiKey(resolved.name));
              if (!destination) addFinding({ ...location, code: "missing-source-target", reason: "An internal target in source relationship markup is absent." });
              else if (destination.relationshipPart) invalid = true;
              else { targetPart = destination.name; fragment = resolved.fragment; }
            }
          } catch (error) { if (!(error instanceof PackageInspectionError)) throw error; invalid = true; }
        }
        if (node.children.some((child) => child.kind === "element")) invalid = true;
        if (invalid) addFinding({ ...location, code: "invalid-source-relationship", reason: "Source relationship fields, target or nested content are outside the supported OPC profile." });
        edges.push(Object.freeze({
          ...location, ...(id !== undefined ? { id } : {}), ...(type !== undefined ? { type } : {}),
          ...(target !== undefined ? { target } : {}), ...(targetMode ? { targetMode } : {}),
          ...(targetPart !== undefined ? { targetPart } : {}), ...(fragment !== undefined ? { fragment } : {}),
          matchesEffectiveRelationship: effective.has(key(source, id, type, target, targetMode, targetPart, fragment)),
        }));
      }
      const condition = frame.alternateIndex && mce && (node.local === "Choice" || node.local === "Fallback")
        ? { alternate: frame.alternateIndex, branch: index, parent: frame.condition } : frame.condition;
      stack.push({ nodes: node.children, index: 0, withinSet: frame.withinSet || set, withinRelationship: frame.withinRelationship || rel, parentIndex: index, alternateIndex: mce && node.local === "AlternateContent" ? index : 0, condition });
    }
  }
  return { edges, sets };
}

/** Different branches of one AlternateContent cannot coexist; independent scopes conservatively may. */
async function mayCoexist(a: BranchCondition | undefined, b: BranchCondition | undefined, work: (units: number) => Promise<void>): Promise<boolean> {
  await work(1);
  if (a === b) return true;
  const choices = new Map<number, number>();
  for (let condition = a; condition; condition = condition.parent) {
    await work(1); choices.set(condition.alternate, condition.branch);
  }
  for (let condition = b; condition; condition = condition.parent) {
    await work(1);
    const branch = choices.get(condition.alternate);
    if (branch !== undefined && branch !== condition.branch) return false;
  }
  return true;
}

function knownAttributes(node: XmlElement): boolean {
  return node.attributes.every((attribute) => {
    if (attribute.namespaceUri === XMLNS) return true;
    if (attribute.namespaceUri === MCE) return directives.has(attribute.local);
    if (attribute.namespaceUri) return false;
    if (node.namespaceUri === RELATIONSHIPS_NS) return node.local === "Relationship" && fields.has(attribute.local);
    return node.namespaceUri === MCE && node.local === "Choice" && attribute.local === "Requires";
  });
}
