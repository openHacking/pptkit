import type { XmlElement } from "../xml/read-xml.js";
import { attribute, children, invalidStructure } from "../presentation/structure-xml.js";
import { XML, XMLNS } from "../compatibility/configuration.js";
import { XML_ID } from "../xml/names.js";

export const CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
export const DC = "http://purl.org/dc/elements/1.1/", TERMS = "http://purl.org/dc/terms/", XSI = "http://www.w3.org/2001/XMLSchema-instance";
export const trimmed = (text: string): string => text.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "");
export function textValue(node: XmlElement, part: string): string {
  if (children(node).length) invalidStructure(`Nested markup in scalar property ${node.local}.`, part);
  return node.children.map((child) => child.kind === "text" || child.kind === "cdata" ? child.value : "").join("");
}
export function noText(node: XmlElement, part: string): void {
  if (node.children.some((child) => (child.kind === "text" || child.kind === "cdata") && trimmed(child.value))) invalidStructure(`Unexpected text in ${node.local}.`, part);
}
export function attributes(node: XmlElement, allowed: readonly string[], part: string): void {
  for (const a of node.attributes) if (a.namespaceUri !== XMLNS && !allowed.includes(`${a.namespaceUri}|${a.local}`)) invalidStructure(`Unexpected ${a.name} attribute on ${node.local}.`, part);
}
export function resolvedType(root: XmlElement, node: XmlElement, part: string): { namespaceUri: string; local: string } {
  const lexical = trimmed(attribute(node, "type", XSI) ?? ""), pieces = lexical.split(":");
  if (pieces.length > 2 || pieces.some((piece) => !XML_ID.test(piece))) invalidStructure("Invalid xsi:type QName.", part);
  const prefix = pieces.length === 2 ? pieces[0]! : "", local = pieces.at(-1)!;
  const scope = new Map<string, string>([["xml", XML]]);
  for (const element of [root, node]) for (const a of element.attributes) if (a.namespaceUri === XMLNS) scope.set(a.prefix === "xmlns" ? a.local : "", a.value);
  const namespaceUri = scope.get(prefix);
  if (!namespaceUri) invalidStructure("Unbound xsi:type QName.", part);
  return { namespaceUri: namespaceUri!, local };
}
export function boolValue(value: string, part: string): void {
  if (!["true", "false", "1", "0"].includes(trimmed(value))) invalidStructure("Invalid metadata boolean.", part);
}
export function integerValue(value: string, min: bigint, max: bigint, part: string): void {
  const lexical = trimmed(value);
  if (!/^[+-]?\d+$/.test(lexical)) invalidStructure("Invalid metadata integer.", part);
  const sign = lexical.startsWith("-") ? "-" : "", digits = lexical.replace(/^[+-]/, "").replace(/^0+/, "") || "0";
  // Strip leading zeros and bound conversion work before using BigInt.
  if (digits.length > 20) invalidStructure("Metadata integer is out of range.", part);
  const number = BigInt(sign + digits);
  if (number < min || number > max) invalidStructure("Metadata integer is out of range.", part);
}
