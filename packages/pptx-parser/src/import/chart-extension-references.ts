import type { readOpcSource } from "../opc/read-package.js";
import type { XmlNode } from "../xml/read-xml.js";

export interface ChartExtensionReference { readonly part: string; readonly uniqueId: string; readonly kind: "attribute" | "text" | "comment" | "processing-instruction" }

/** Inspect retained, unprojected XML only; opaque and nested payloads remain unproved. */
export async function inspectChartExtensionReferences(source: Awaited<ReturnType<typeof readOpcSource>>, chartPart: string, uniqueIds: readonly string[], maxWork: number, checkpoint: () => Promise<void>): Promise<{ readonly references: readonly ChartExtensionReference[]; readonly opaqueParts: readonly string[]; readonly workUnits: number }> {
  if (!source || typeof source !== "object" || !source.xmlDocuments || typeof chartPart !== "string" || !chartPart || !Array.isArray(uniqueIds)) throw new TypeError("Invalid chart extension reference input.");
  if (!Number.isSafeInteger(maxWork) || maxWork <= 0) throw new RangeError("Chart extension reference inspection requires a positive work limit.");
  if (!uniqueIds.length || uniqueIds.length > maxWork) throw new RangeError("Chart extension identities exceed the work limit.");
  const ids = uniqueIds.map(value => typeof value === "string" ? value.toUpperCase() : "");
  if (ids.some(value => !/^\{[0-9A-F]{8}-(?:[0-9A-F]{4}-){3}[0-9A-F]{12}\}$/u.test(value)) || new Set(ids).size !== ids.length) throw new TypeError("Chart extension identifiers must be unique GUIDs.");
  if (!source.xmlDocuments.has('/[Content_Types].xml') || source.inspection.parts.some(part => part.xml && !source.xmlDocuments.has(part.name))) throw new TypeError("Chart extension inspection requires every original XML document, including inactive branches.");
  let workUnits = 0;
  const charge = async (units: number) => { workUnits += units; if (!Number.isSafeInteger(workUnits) || workUnits > maxWork) throw new RangeError("Chart extension reference inspection exceeds its work limit."); await checkpoint(); };
  const references: ChartExtensionReference[] = [], declarations = new Map(ids.map(id => [id, 0])), expectedNamespace = "http://schemas.microsoft.com/office/drawing/2014/chart";
  const inspect = async (part: string, value: string, kind: ChartExtensionReference["kind"], declaration = false) => {
    await charge(value.length * ids.length + 1); const upper = value.toUpperCase();
    for (const id of ids) if (upper.includes(id)) {
      if (declaration && upper === id) declarations.set(id, declarations.get(id)! + 1);
      else references.push({ part, uniqueId: id, kind });
    }
  };
  for (const [part, document] of source.xmlDocuments) {
    const pending: XmlNode[] = [...document.children].reverse();
    while (pending.length) {
      await charge(1); const node = pending.pop()!;
      if (node.kind === "element") {
        for (const attribute of node.attributes) await inspect(part, attribute.value, "attribute", part === chartPart && node.namespaceUri === expectedNamespace && node.local === "uniqueId" && attribute.namespaceUri === "" && attribute.local === "val");
        for (let i = node.children.length - 1; i >= 0; i--) pending.push(node.children[i]!);
      } else if (node.kind === "processing-instruction") { await inspect(part, node.target, "processing-instruction"); await inspect(part, node.value, "processing-instruction"); }
      else await inspect(part, node.value, node.kind === "comment" ? "comment" : "text");
    }
  }
  if (ids.some(id => declarations.get(id) !== 1)) throw new TypeError("Each source unique ID must have exactly one chart declaration.");
  const opaqueParts = source.inspection.parts.filter(part => !part.xml).map(part => part.name);
  return { references, opaqueParts, workUnits };
}
