import type { MarkupCompatibilityOptions } from "../types/compatibility.js";
import { XML_ID } from "../xml/names.js";

export const MCE = "http://schemas.openxmlformats.org/markup-compatibility/2006";
export const XMLNS = "http://www.w3.org/2000/xmlns/";
export const XML = "http://www.w3.org/XML/1998/namespace";
export const expandedKey = (uri: string, local: string): string => JSON.stringify([uri, local]);
export function captureConfiguration(options: MarkupCompatibilityOptions) {
  if (!options || typeof options !== "object" || Array.isArray(options) || !Array.isArray(options.understoodNamespaces)) throw new TypeError("MCE understoodNamespaces must be an array");
  const understood = new Set<string>([XML, XMLNS]);
  for (const uri of options.understoodNamespaces) {
    if (typeof uri !== "string" || !uri || /[\s\u0000-\u001f]/u.test(uri) || uri === MCE) throw new TypeError("Invalid understood namespace URI");
    understood.add(uri);
  }
  if (options.understandNoNamespace !== undefined && typeof options.understandNoNamespace !== "boolean") throw new TypeError("understandNoNamespace must be boolean");
  if (options.understandNoNamespace !== false) understood.add("");
  if (options.extensionElements !== undefined && !Array.isArray(options.extensionElements)) throw new TypeError("extensionElements must be an array");
  const extensions = new Set<string>();
  for (const name of options.extensionElements ?? []) {
    if (!name || typeof name.namespaceUri !== "string" || /[\s\u0000-\u001f]/u.test(name.namespaceUri) || name.namespaceUri === MCE || typeof name.local !== "string" || !XML_ID.test(name.local)) throw new TypeError("Invalid extension expanded name");
    extensions.add(expandedKey(name.namespaceUri, name.local));
  }
  return { understood, extensions };
}
export type CompatibilityConfiguration = ReturnType<typeof captureConfiguration>;
