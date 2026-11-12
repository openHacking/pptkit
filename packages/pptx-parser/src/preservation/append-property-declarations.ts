import { CONTENT_TYPES_NS, RELATIONSHIPS_NS } from "../opc/markup.js";
import { asciiKey, partName } from "../opc/names.js";
import { PackageInspectionError } from "../types/archive.js";
import type { XmlInspectionOptions } from "../types/xml.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { DEFAULT_XML_LIMITS, readXml, type XmlElement } from "../xml/read-xml.js";
import { appendXmlRootChild } from "./append-xml-root-child.js";
import { XmlSourceEditError } from "./xml-text-source.js";

const TRANSITIONAL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
const STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships/customProperties";
const MIME = "application/vnd.openxmlformats-officedocument.custom-properties+xml";

type Request = { readonly part: string; readonly relationshipId: string; readonly dialect: "strict" | "transitional" };
type Options = XmlInspectionOptions & { readonly maxOutputBytes: number };

/** Add only the declarations; callers must separately prove package closure and identity. */
export async function appendPropertyDeclarations(relationships: Uint8Array, contentTypes: Uint8Array, request: Request & { readonly kind: "core" | "custom" | "extended" }, options: Options): Promise<{ relationships: Uint8Array; contentTypes: Uint8Array }> {
  if (!(relationships instanceof Uint8Array) || !(contentTypes instanceof Uint8Array) || !request || !options) throw new TypeError("Invalid property declaration request.");
  const { part, relationshipId, dialect, kind } = request, { signal, partName: sourceName, maxOutputBytes } = options, limits = { ...DEFAULT_XML_LIMITS, ...options.limits };
  if (typeof part !== "string" || typeof relationshipId !== "string" || !["strict", "transitional"].includes(dialect) || !Number.isSafeInteger(maxOutputBytes) || maxOutputBytes <= 0 || kind !== "core" && kind !== "custom" && kind !== "extended") throw new TypeError("Invalid property declaration request.");
  if (part.length + relationshipId.length > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "Property declaration exceeds the output limit.");
  if (!/^\/[A-Za-z0-9_.-]+(?:\/[A-Za-z0-9_.-]+)*$/.test(part) || part.slice(1).split("/").some(segment => segment === "." || segment === ".." || segment.endsWith(".")) || !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(relationshipId)) throw new XmlSourceEditError("xml-edit-unsupported", "Property declarations require a safe ASCII part and relationship ID.");
  const coreRel = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
  const transitionalRel = kind === "extended" ? "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" : TRANSITIONAL;
  const strictRel = kind === "extended" ? "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties" : STRICT;
  const allowedRelationships = kind === "core" ? [coreRel] : [transitionalRel, strictRel];
  const selectedRel = kind === "core" ? coreRel : dialect === "strict" ? strictRel : transitionalRel;
  const mimeType = kind === "core" ? "application/vnd.openxmlformats-package.core-properties+xml" : kind === "extended" ? "application/vnd.openxmlformats-officedocument.extended-properties+xml" : MIME;
  const check = () => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "Property declaration insertion was aborted.", sourceName); };
  check();
  if (byteViewInfo(relationships).shared || byteViewInfo(contentTypes).shared) throw new TypeError("XML insertion requires ordinary byte buffers.");
  if (byteViewInfo(relationships).length > limits.maxInputBytes || byteViewInfo(contentTypes).length > limits.maxInputBytes) throw new PackageInspectionError("xml-limit-exceeded", "Declaration XML exceeds maxInputBytes.", sourceName);
  const relBytes = copyByteView(relationships), typeBytes = copyByteView(contentTypes), captured = { limits, ...(signal ? { signal } : {}), ...(sourceName !== undefined ? { partName: sourceName } : {}) };
  const rel = await readXml(relBytes, captured, true), types = await readXml(typeBytes, captured, true); check();
  if (!rel.root || rel.root.namespaceUri !== RELATIONSHIPS_NS || rel.root.local !== "Relationships" || !types.root || types.root.namespaceUri !== CONTENT_TYPES_NS || types.root.local !== "Types") throw new XmlSourceEditError("xml-edit-conflict", "Expected OPC declaration roots.");
  const mce = "http://schemas.openxmlformats.org/markup-compatibility/2006";
  if (rel.inspection.namespaceUris.includes(mce) || types.inspection.namespaceUris.includes(mce)) throw new XmlSourceEditError("xml-edit-unsupported", "Compatibility markup in OPC declarations is not supported for insertion.");
  const key = asciiKey(partName(part, sourceName ?? "[Content_Types].xml"));
  const scan = async (root: XmlElement, kind: "relationships" | "types") => {
    const stack = [root]; let visited = 0;
    while (stack.length) {
      const node = stack.pop()!;
      const attr = (local: string) => node.attributes.find(item => !item.namespaceUri && item.local === local)?.value;
      if (kind === "relationships" && node.namespaceUri === RELATIONSHIPS_NS && node.local === "Relationship") {
        const id = attr("Id")?.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "").replace(/[\t\r\n ]+/g, " "), type = attr("Type")?.trim();
        if (id === relationshipId || allowedRelationships.includes(type ?? "")) throw new XmlSourceEditError("xml-edit-conflict", "Raw relationship ID or property type already exists.");
      }
      if (kind === "types" && node.namespaceUri === CONTENT_TYPES_NS && (node.local === "Override" || node.local === "Default")) {
        const name = node.local === "Override" ? attr("PartName") : undefined, mime = attr("ContentType");
        if (name !== undefined && asciiKey(partName(name, sourceName ?? "[Content_Types].xml")) === key || mime?.trim().toLowerCase() === mimeType) throw new XmlSourceEditError("xml-edit-conflict", "Raw property part or content type already exists.");
      }
      for (let i = node.children.length - 1; i >= 0; i--) { const child = node.children[i]!; if (child.kind === "element") stack.push(child); }
      if (++visited % 256 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
    }
  };
  await scan(rel.root, "relationships"); await scan(types.root, "types");
  const relPrefix = rel.root.prefix ? `${rel.root.prefix}:` : "", typePrefix = types.root.prefix ? `${types.root.prefix}:` : "";
  const relChild = `<${relPrefix}Relationship Id="${relationshipId}" Type="${selectedRel}" Target="${part.slice(1)}"/>`;
  const typeChild = `<${typePrefix}Override PartName="${part}" ContentType="${mimeType}"/>`;
  const relResult = await appendXmlRootChild(relBytes, rel.root, rel.inspection.encoding, relChild, maxOutputBytes, captured);
  const typeResult = await appendXmlRootChild(typeBytes, types.root, types.inspection.encoding, typeChild, maxOutputBytes, captured);
  const child = (root: XmlElement) => root.children.filter((item): item is XmlElement => item.kind === "element").at(-1);
  const relNode = child(relResult.root), typeNode = child(typeResult.root);
  const field = (node: XmlElement | undefined, name: string) => node?.attributes.find(item => !item.namespaceUri && item.local === name)?.value;
  if (relResult.root.children.length !== rel.root.children.length + 1 || typeResult.root.children.length !== types.root.children.length + 1 || !relNode || relNode.namespaceUri !== RELATIONSHIPS_NS || relNode.local !== "Relationship" || field(relNode, "Id") !== relationshipId || field(relNode, "Type") !== selectedRel || field(relNode, "Target") !== part.slice(1) || !typeNode || typeNode.namespaceUri !== CONTENT_TYPES_NS || typeNode.local !== "Override" || field(typeNode, "PartName") !== part || field(typeNode, "ContentType") !== mimeType) throw new XmlSourceEditError("xml-edit-conflict", "Property declarations failed verification.");
  return { relationships: relResult.bytes, contentTypes: typeResult.bytes };
}
