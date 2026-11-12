import type { CorePropertyInfo, MetadataKeywordSegment } from "../types/metadata.js";
import type { XmlElement } from "../xml/read-xml.js";
import { attribute, children, invalidStructure } from "../presentation/structure-xml.js";
import { XML } from "../compatibility/configuration.js";
import { attributes, CORE_NS, DC, TERMS, XSI, noText, resolvedType, textValue } from "./xml.js";

const CP_NAMES = new Set(["category", "contentStatus", "keywords", "lastModifiedBy", "lastPrinted", "revision", "version"]);
const DC_NAMES = new Set(["creator", "description", "identifier", "language", "subject", "title"]);
export async function coreProperties(root: XmlElement, part: string, checkpoint: () => Promise<void>): Promise<CorePropertyInfo[]> {
  attributes(root, [], part); noText(root, part);
  const properties: CorePropertyInfo[] = [], seen = new Set<string>();
  for (const node of children(root)) {
    await checkpoint();
    const valid = node.namespaceUri === CORE_NS ? CP_NAMES.has(node.local) : node.namespaceUri === DC ? DC_NAMES.has(node.local) : node.namespaceUri === TERMS && ["created", "modified"].includes(node.local);
    if (!valid || seen.has(node.local)) invalidStructure(`Unknown or duplicate core property ${node.name}.`, part);
    seen.add(node.local);
    const name = node.local as CorePropertyInfo["name"];
    if (name === "keywords") {
      attributes(node, [`${XML}|lang`], part);
      const language = attribute(node, "lang", XML), segments: MetadataKeywordSegment[] = [];
      for (const child of node.children) {
        await checkpoint();
        if (child.kind === "element") {
          if (child.namespaceUri !== CORE_NS || child.local !== "value") invalidStructure("Unknown keyword child.", part);
          attributes(child, [`${XML}|lang`], part);
          const own = attribute(child, "lang", XML), effective = own ?? language;
          segments.push({ kind: "value", value: textValue(child, part), ...(effective !== undefined ? { language: effective } : {}) });
        } else if (child.kind === "text" || child.kind === "cdata") segments.push({ kind: "text", value: child.value, ...(language !== undefined ? { language } : {}) });
      }
      properties.push({ name, value: segments.map((segment) => segment.value).join(""), keywordSegments: segments, ...(language !== undefined ? { language } : {}) });
    } else {
      attributes(node, node.namespaceUri === TERMS ? [`${XSI}|type`] : [], part);
      if (node.namespaceUri === TERMS) {
        const type = resolvedType(root, node, part);
        if (type.namespaceUri !== TERMS || type.local !== "W3CDTF") invalidStructure("Date property must declare dcterms:W3CDTF.", part);
      }
      properties.push({ name, value: textValue(node, part) });
    }
  }
  return properties;
}
