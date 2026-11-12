import { PackageInspectionError } from "../types/archive.js";
import type { XmlElement } from "../xml/read-xml.js";
export const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
export const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
export const RELATIONSHIPS_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
const XMLNS = "http://www.w3.org/2000/xmlns/";
export function packageInvalid(message: string, entry?: string): never { throw new PackageInspectionError("package-invalid", message, entry); }
export function attributesOf(element: XmlElement, allowed: readonly string[], required: readonly string[], entry: string, extensionPolicy: "invalid" | "unsupported" = "invalid"): Map<string, string> {
  const result = new Map<string, string>();
  for (const attribute of element.attributes) {
    if (attribute.namespaceUri === XMLNS) continue;
    if (attribute.namespaceUri || !allowed.includes(attribute.local)) {
      if (attribute.namespaceUri && attribute.namespaceUri !== "http://www.w3.org/XML/1998/namespace" && extensionPolicy === "unsupported") throw new PackageInspectionError("package-unsupported", "Relationship markup extensions require an MCE processor.", entry);
      packageInvalid("Unexpected attribute in OPC control markup.", entry);
    }
    result.set(attribute.local, attribute.value);
  }
  for (const name of required) if (!result.get(name)) packageInvalid(`Missing or empty OPC attribute ${name}.`, entry);
  return result;
}
export function elementsOf(element: XmlElement, entry: string, allowText = false): XmlElement[] {
  const children: XmlElement[] = [];
  for (const child of element.children) {
    if (child.kind === "element") children.push(child);
    else if (!allowText && (child.kind === "text" || child.kind === "cdata") && !/^[\t\n\r ]*$/.test(child.value)) packageInvalid("OPC control markup cannot contain character content.", entry);
  }
  return children;
}
const TOKEN = "[!#$%&'*+.^_`|~0-9A-Za-z-]+";
const MEDIA = new RegExp(`^(${TOKEN}/${TOKEN})((?:[ \\t]*;[ \\t]*${TOKEN}=(?:${TOKEN}|"(?:[\\t !#-\\[\\]-~]|\\\\[\\t -~])*"))*)$`);
export function mediaType(value: string, entry: string): string {
  const match = MEDIA.exec(value);
  if (!match) packageInvalid("Invalid MIME content type.", entry);
  return match![1]!.toLowerCase();
}
export { XML_ID } from "../xml/names.js";


export function controlEncoding(declared: string | undefined, entry: string): void {
  if (declared && !["utf-8", "utf-16"].includes(declared.toLowerCase())) packageInvalid("OPC-defined XML declarations must name UTF-8 or UTF-16.", entry);
}
