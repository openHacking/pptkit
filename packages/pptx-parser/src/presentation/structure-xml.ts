import { PackageInspectionError } from "../types/archive.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { PptxSize } from "../types/structure.js";

export function invalidStructure(message: string, part: string): never { throw new PackageInspectionError("presentation-invalid", message, part); }
export const attribute = (node: XmlElement, local: string, namespaceUri = ""): string | undefined => node.attributes.find((a) => a.local === local && a.namespaceUri === namespaceUri)?.value;
export const children = (node: XmlElement): XmlElement[] => node.children.filter((n): n is XmlElement => n.kind === "element");
export function singleChild(node: XmlElement, namespaceUri: string, local: string, part: string, required = false): XmlElement | undefined {
  const matches = children(node).filter((n) => n.namespaceUri === namespaceUri && n.local === local);
  if (matches.length > 1 || (required && !matches.length)) invalidStructure(`Expected ${required ? "one" : "at most one"} ${local} child in ${node.local}.`, part);
  return matches[0];
}
export function integer(value: string | undefined, min: number, max: number, label: string, part: string): number {
  const normalized = value?.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "");
  const number = Number(normalized);
  if (!normalized || !/^[+-]?\d+$/.test(normalized) || !Number.isSafeInteger(number) || number < min || number > max) invalidStructure(`Invalid ${label}: expected integer in [${min}, ${max}].`, part);
  return number;
}
export function size(node: XmlElement, min: number, max: number, part: string): PptxSize {
  return { widthEmu: integer(attribute(node, "cx"), min, max, `${node.local}/@cx`, part), heightEmu: integer(attribute(node, "cy"), min, max, `${node.local}/@cy`, part) };
}
export function hiddenSlide(node: XmlElement, part: string): boolean {
  const show = attribute(node, "show")?.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "");
  if (show === undefined || show === "true" || show === "1") return false;
  if (show === "false" || show === "0") return true;
  return invalidStructure("Invalid slide show boolean.", part);
}
