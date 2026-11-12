import { escapeXml } from "./xml.js";

/** Encode Xstring escapes and XML-forbidden controls before ordinary XML escaping. */
export function chartStringXml(value: string): string {
  const encoded = value.replace(/_(?=x[0-9a-fA-F]{4}_)|[\u0000-\u0008\u000b-\u000d\u000e-\u001f\ufffe\uffff]/g,
    match => match === "_" ? "_x005F_" : `_x${match.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}_`);
  return escapeXml(encoded);
}
