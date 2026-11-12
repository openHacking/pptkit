import { RELATIONSHIPS_NS } from "../opc/markup.js";
import { PackageInspectionError } from "../types/archive.js";
import type { XmlInspectionOptions } from "../types/xml.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { DEFAULT_XML_LIMITS, readXml, type XmlElement } from "../xml/read-xml.js";
import { XmlSourceEditError } from "./xml-text-source.js";
import { appendXmlRootChild } from "./append-xml-root-child.js";

/** Private lexical operation; package-level ownership, targets and closure require separate proof. */
export async function appendImageRelationship(input: Uint8Array, request: { readonly id: string; readonly target: string; readonly dialect: "strict" | "transitional" }, options: XmlInspectionOptions & { readonly maxOutputBytes: number }): Promise<Uint8Array> {
  if (!(input instanceof Uint8Array) || !request || typeof request !== "object" || !options || typeof options !== "object") throw new TypeError("Invalid relationship insertion input.");
  const { id, target, dialect } = request, { signal, partName, maxOutputBytes } = options, limits = { ...DEFAULT_XML_LIMITS, ...options.limits };
  if (typeof id !== "string" || typeof target !== "string" || (dialect !== "strict" && dialect !== "transitional")) throw new TypeError("Invalid image relationship request.");
  if (!Number.isSafeInteger(maxOutputBytes) || maxOutputBytes <= 0) throw new RangeError("Invalid relationship output limit.");
  if (id.length + target.length > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "Relationship fields exceed the output budget.");
  if (!/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(id) || !/^(?:\.\.\/)*[A-Za-z0-9_.-]+(?:\/[A-Za-z0-9_.-]+)*$/.test(target) || target.replace(/^(?:\.\.\/)+/, "").split("/").some(segment => segment === "." || segment === "..")) throw new XmlSourceEditError("xml-edit-unsupported", "Image relationship requires a safe ID and relative ASCII target.");
  const check = () => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "Image relationship insertion was aborted.", partName); };
  check(); const view = byteViewInfo(input);
  if (view.shared) throw new TypeError("Relationship insertion requires an ordinary byte buffer.");
  if (!Number.isSafeInteger(limits.maxInputBytes) || limits.maxInputBytes <= 0) throw new RangeError("Invalid XML limit maxInputBytes.");
  if (view.length > limits.maxInputBytes) throw new PackageInspectionError("xml-limit-exceeded", "Relationship source exceeds maxInputBytes.", partName);
  const bytes = copyByteView(input), captured = { limits, ...(partName !== undefined ? { partName } : {}), ...(signal ? { signal } : {}) };
  const parsed = await readXml(bytes, captured, true); check();
  const root = parsed.root!;
  if (root.namespaceUri !== RELATIONSHIPS_NS || root.local !== "Relationships") throw new XmlSourceEditError("xml-edit-conflict", "Expected the OPC Relationships root.");
  const stack: XmlElement[] = [root]; let visited = 0;
  while (stack.length) {
    const node = stack.pop()!;
    if (node.namespaceUri === RELATIONSHIPS_NS && node.local === "Relationship") {
      const existing = node.attributes.find(attribute => !attribute.namespaceUri && attribute.local === "Id")?.value.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "").replace(/[\t\r\n ]+/g, " ");
      if (existing === id) throw new XmlSourceEditError("xml-edit-conflict", "Relationship ID already exists in raw source markup.");
    }
    for (let i = node.children.length - 1; i >= 0; i--) { const child = node.children[i]!; if (child.kind === "element") stack.push(child); }
    if (++visited % 256 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
  }
  const prefix = root.prefix ? `${root.prefix}:` : "", type = dialect === "strict" ? "http://purl.oclc.org/ooxml/officeDocument/relationships/image" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
  const child = `<${prefix}Relationship Id="${id}" Type="${type}" Target="${target}"/>`;
  const verified = await appendXmlRootChild(bytes, root, parsed.inspection.encoding, child, maxOutputBytes, captured);
  const children = verified.root.children.filter((child): child is XmlElement => child.kind === "element"), addedNode = children.at(-1);
  const fields = new Map(addedNode?.attributes.filter(attribute => !attribute.namespaceUri).map(attribute => [attribute.local, attribute.value]));
  if (!addedNode || addedNode.namespaceUri !== RELATIONSHIPS_NS || addedNode.local !== "Relationship" || fields.get("Id") !== id || fields.get("Type") !== type || fields.get("Target") !== target || children.length !== root.children.filter(child => child.kind === "element").length + 1) throw new XmlSourceEditError("xml-edit-unsupported", "Relationship candidate failed verification.");
  return verified.bytes;
}
