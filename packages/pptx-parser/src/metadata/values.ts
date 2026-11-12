import type { MetadataPropertyValue, MetadataScalarKind } from "../types/metadata.js";
import type { XmlElement } from "../xml/read-xml.js";
import { invalidStructure } from "../presentation/structure-xml.js";
import { attributes, boolValue, integerValue, textValue, trimmed } from "./xml.js";

const INTEGER_BITS: Readonly<Record<string, number>> = { i1: 8, i2: 16, i4: 32, i8: 64, int: 32, ui1: 8, ui2: 16, ui4: 32, ui8: 64, uint: 32 };
const OPAQUE = new Set(["vector", "array", "blob", "oblob", "cy", "error", "stream", "ostream", "storage", "ostorage", "vstream", "clsid"]);
export const uninterpreted = (node: XmlElement): MetadataPropertyValue => ({ status: "uninterpreted", namespaceUri: node.namespaceUri, local: node.local });
export function scalar(node: XmlElement, kind: MetadataScalarKind, part: string): MetadataPropertyValue {
  attributes(node, [], part);
  const value = textValue(node, part), lexical = trimmed(value);
  if (kind === "boolean") boolValue(value, part);
  if ((kind === "empty" || kind === "null") && lexical) invalidStructure("Empty/null property contains text.", part);
  if (kind === "decimal" && !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(lexical)) invalidStructure("Invalid decimal property.", part);
  if (kind === "real" && !/^(?:[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|INF|-INF|NaN)$/.test(lexical)) invalidStructure("Invalid real property.", part);
  // Dates are source lexical values, not a calendar-validation or timezone-conversion API.
  return { status: "scalar", kind, value };
}
export function variant(node: XmlElement, namespaceUri: string, part: string): MetadataPropertyValue {
  if (node.namespaceUri !== namespaceUri) invalidStructure("Unexpected custom value namespace.", part);
  if (OPAQUE.has(node.local)) return uninterpreted(node);
  const bits = INTEGER_BITS[node.local];
  if (typeof bits === "number") {
    const result = scalar(node, "integer", part);
    const unsigned = node.local.startsWith("u"), width = BigInt(bits), upper = 1n << (unsigned ? width : width - 1n);
    integerValue(textValue(node, part), unsigned ? 0n : -upper, upper - 1n, part);
    return result;
  }
  const kind: MetadataScalarKind | undefined = ["lpstr", "lpwstr", "bstr"].includes(node.local) ? "string" : ["date", "filetime"].includes(node.local) ? "date" : ["r4", "r8"].includes(node.local) ? "real" : node.local === "bool" ? "boolean" : ["decimal", "empty", "null"].includes(node.local) ? node.local as MetadataScalarKind : undefined;
  if (!kind) invalidStructure(`Unknown custom value type ${node.local}.`, part);
  return scalar(node, kind!, part);
}
