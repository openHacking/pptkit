import type { OpcPackageInspection } from "../types/opc.js";
import type { PptxDocumentMetadata } from "../types/metadata.js";
import type { MarkupCompatibilityReport } from "../types/compatibility.js";
import type { XmlElement } from "../xml/read-xml.js";
import { captureConfiguration, MCE } from "../compatibility/configuration.js";
import { processCompatibility } from "../compatibility/process.js";
import { invalidStructure } from "../presentation/structure-xml.js";
import { CORE_NS, trimmed } from "./xml.js";
import { coreProperties } from "./core.js";
import { customProperties, extendedProperties } from "./properties.js";

const T = "http://schemas.openxmlformats.org/officeDocument/2006", S = "http://purl.oclc.org/ooxml/officeDocument";
const CORE_REL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
export async function readMetadata(source: { inspection: OpcPackageInspection; xml: ReadonlyMap<string, XmlElement> }, strict: boolean, checkpoint: () => Promise<void>, signal?: AbortSignal) {
  const metadata: { -readonly [K in keyof PptxDocumentMetadata]: PptxDocumentMetadata[K] } = {};
  let customRoot: XmlElement | undefined;
  const reports: { part: string; report: MarkupCompatibilityReport }[] = [];
  const base = strict ? S : T, vt = `${base}/docPropsVTypes`;
  const specs = [
    { key: "core", namespaceUri: CORE_NS, local: "coreProperties", mime: "application/vnd.openxmlformats-package.core-properties+xml", rel: CORE_REL, alternatives: [CORE_REL] },
    ...(["extended", "custom"] as const).map((key) => ({ key, namespaceUri: `${base}/${strict ? `${key}Properties` : `${key}-properties`}`, local: "Properties", mime: `application/vnd.openxmlformats-officedocument.${key}-properties+xml`, rel: `${base}/relationships/${strict ? `${key}Properties` : `${key}-properties`}`, alternatives: [`${T}/relationships/${key}-properties`, `${S}/relationships/${key}Properties`] })),
  ];
  for (const spec of specs) {
    await checkpoint();
    const parts = source.inspection.parts.filter((part) => part.contentType.toLowerCase() === spec.mime);
    const relations = source.inspection.relationships.filter((r) => r.source === "/" && spec.alternatives.includes(r.type));
    if (parts.length > 1 || relations.length > 1) invalidStructure(`Multiple ${spec.key} metadata parts/relationships.`, "/");
    if (!parts.length && !relations.length) continue;
    const part = parts[0], relation = relations[0];
    if (!part || !relation || relation.type !== spec.rel || relation.targetMode !== "Internal" || relation.fragment !== undefined || relation.targetPart !== part.name) invalidStructure(`Expected a package relationship to one internal ${spec.key} metadata part in the presentation dialect.`, part?.name ?? "/");
    const name = part!.name;
    let root = source.xml.get(name);
    if (!root) invalidStructure("Metadata part must contain XML.", name);
    // OPC core properties forbid MCE markup. Other properties use an explicit profile.
    if (spec.key !== "core" && part!.xml?.namespaceUris.includes(MCE)) {
      const projected = await processCompatibility(root!, captureConfiguration({ understoodNamespaces: [spec.namespaceUri, vt] }), name, signal);
      const roots = projected.nodes.filter((node): node is XmlElement => node.kind === "element");
      if (roots.length !== 1 || projected.nodes.some((node) => (node.kind === "text" || node.kind === "cdata") && trimmed(node.value))) invalidStructure("Metadata must project to one root.", name);
      root = roots[0]!; reports.push({ part: name, report: projected.report });
    }
    if (root!.namespaceUri !== spec.namespaceUri || root!.local !== spec.local) invalidStructure("Unexpected metadata root namespace/name.", name);
    const ref = { part: name, relationshipId: relation!.id };
    if (spec.key === "core") metadata.core = { ...ref, properties: await coreProperties(root!, name, checkpoint) };
    else if (spec.key === "extended") metadata.extended = { ...ref, properties: extendedProperties(root!, vt, name) };
    else { customRoot = root!; metadata.custom = { ...ref, properties: await customProperties(root!, vt, name, checkpoint) }; }
  }
  return { metadata, reports, customRoot };
}
