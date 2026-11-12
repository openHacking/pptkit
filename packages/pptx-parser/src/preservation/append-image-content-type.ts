import { CONTENT_TYPES_NS } from "../opc/markup.js";
import { asciiKey, partName } from "../opc/names.js";
import { PackageInspectionError } from "../types/archive.js";
import type { XmlInspectionOptions } from "../types/xml.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { DEFAULT_XML_LIMITS, readXml, type XmlElement } from "../xml/read-xml.js";
import { appendXmlRootChild } from "./append-xml-root-child.js";
import { XmlSourceEditError } from "./xml-text-source.js";

/** Append an explicit raster override. It grants no package write permission. */
export async function appendImageContentType(input: Uint8Array, request: { readonly part: string; readonly mimeType: "image/png" | "image/jpeg" | "image/gif" }, options: XmlInspectionOptions & { readonly maxOutputBytes: number }): Promise<Uint8Array> {
  if (!(input instanceof Uint8Array) || !request || typeof request !== "object" || !options || typeof options !== "object") throw new TypeError("Invalid image content type input.");
  const { part, mimeType } = request, { signal, partName: sourceName, maxOutputBytes } = options, limits = { ...DEFAULT_XML_LIMITS, ...options.limits };
  if (typeof part !== "string" || !["image/png", "image/jpeg", "image/gif"].includes(mimeType)) throw new TypeError("Invalid image content type request.");
  if (!Number.isSafeInteger(maxOutputBytes) || maxOutputBytes <= 0) throw new RangeError("Invalid content type output limit.");
  if (part.length > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "Image part name exceeds the output budget.");
  if (!/^\/[A-Za-z0-9_.-]+(?:\/[A-Za-z0-9_.-]+)*$/.test(part) || part.slice(1).split("/").some(segment => segment === "." || segment === ".." || segment.endsWith("."))) throw new XmlSourceEditError("xml-edit-unsupported", "Image override requires a safe absolute ASCII part name.");
  const key = asciiKey(partName(part, sourceName ?? "[Content_Types].xml"));
  const check = () => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "Image content type insertion was aborted.", sourceName); };
  check(); const view = byteViewInfo(input);
  if (view.shared) throw new TypeError("Content type insertion requires an ordinary byte buffer.");
  if (!Number.isSafeInteger(limits.maxInputBytes) || limits.maxInputBytes <= 0) throw new RangeError("Invalid XML limit maxInputBytes.");
  if (view.length > limits.maxInputBytes) throw new PackageInspectionError("xml-limit-exceeded", "Content type source exceeds maxInputBytes.", sourceName);
  const bytes = copyByteView(input), captured = { limits, ...(sourceName !== undefined ? { partName: sourceName } : {}), ...(signal ? { signal } : {}) };
  const parsed = await readXml(bytes, captured, true); check(); const root = parsed.root!;
  if (root.namespaceUri !== CONTENT_TYPES_NS || root.local !== "Types") throw new XmlSourceEditError("xml-edit-conflict", "Expected the OPC content types root.");
  const stack: XmlElement[] = [root]; let visited = 0;
  while (stack.length) {
    const node = stack.pop()!;
    if (node.namespaceUri === CONTENT_TYPES_NS && node.local === "Override") {
      const value = node.attributes.find(attribute => !attribute.namespaceUri && attribute.local === "PartName")?.value;
      if (value !== undefined && asciiKey(partName(value, sourceName ?? "[Content_Types].xml")) === key) throw new XmlSourceEditError("xml-edit-conflict", "Image content type override already exists in source markup.");
    }
    for (let i = node.children.length - 1; i >= 0; i--) { const child = node.children[i]!; if (child.kind === "element") stack.push(child); }
    if (++visited % 256 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
  }
  const prefix = root.prefix ? `${root.prefix}:` : "", child = `<${prefix}Override PartName="${part}" ContentType="${mimeType}"/>`;
  const verified = await appendXmlRootChild(bytes, root, parsed.inspection.encoding, child, maxOutputBytes, captured);
  const children = verified.root.children.filter((node): node is XmlElement => node.kind === "element"), added = children.at(-1);
  const fields = new Map(added?.attributes.filter(attribute => !attribute.namespaceUri).map(attribute => [attribute.local, attribute.value]));
  if (!added || added.namespaceUri !== CONTENT_TYPES_NS || added.local !== "Override" || fields.get("PartName") !== part || fields.get("ContentType") !== mimeType || children.length !== root.children.filter(node => node.kind === "element").length + 1) throw new XmlSourceEditError("xml-edit-unsupported", "Image content type candidate failed verification.");
  return verified.bytes;
}
