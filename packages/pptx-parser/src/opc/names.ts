import { PackageInspectionError } from "../types/archive.js";
const invalid = (message: string, entry: string): never => { throw new PackageInspectionError("package-invalid", message, entry); };
export const asciiKey = (value: string): string => value.replace(/[A-Z]/g, (character) => character.toLowerCase());
const unreserved = /^[A-Za-z0-9._~-]$/;
function iriUnicode(code: number): boolean {
  return (code >= 0xa0 && code <= 0xd7ff) || (code >= 0xf900 && code <= 0xfdcf) || (code >= 0xfdf0 && code <= 0xffef) || (code >= 0x10000 && code <= 0xefffd && (code & 0xffff) <= 0xfffd && !(code >= 0xe0000 && code < 0xe1000));
}
/** Logical names use literal Unicode and preserved escaped ASCII reserved characters. */
export function partName(value: string, entry: string, fromZip = false): string {
  if (fromZip && /[^\x00-\x7f]/.test(value)) throw new PackageInspectionError("package-unsupported", "OPC ZIP names must encode non-ASCII characters as UTF-8 percent sequences.", entry);
  if (!value.startsWith("/") || value.endsWith("/") || /%(?![\da-f]{2})/i.test(value)) invalid("Invalid absolute OPC part name.", entry);
  for (const match of value.matchAll(/%([\da-f]{2})/gi)) {
    const byte = Number.parseInt(match[1]!, 16), character = String.fromCharCode(byte);
    if (byte < 128 && (unreserved.test(character) || character === "/" || character === "\\" || byte < 32 || byte === 127)) invalid("Forbidden percent-encoded character in OPC part name.", entry);
  }
  let logical: string;
  try { logical = value.replace(/(?:%[89a-f][\da-f])+/gi, (sequence) => decodeURIComponent(sequence)).replace(/%[\da-f]{2}/gi, (sequence) => sequence.toUpperCase()); }
  catch { return invalid("Invalid UTF-8 in OPC part name.", entry); }
  for (const segment of logical.slice(1).split("/")) {
    if (!segment || segment.endsWith(".")) invalid("OPC part segments must be nonempty and cannot end in a dot.", entry);
    for (let i = 0; i < segment.length;) {
      const code = segment.codePointAt(i)!, character = String.fromCodePoint(code);
      if (character === "%") { i += 3; continue; }
      if (!(unreserved.test(character) || /^[!$&'()*+,;=:@]$/.test(character) || iriUnicode(code))) invalid("Forbidden character in OPC part name.", entry);
      i += character.length;
    }
  }
  return logical;
}

/** Returns undefined for ordinary parts; '/' for the package relationship set. */
export function relationshipSource(name: string): string | undefined {
  const segments = name.split("/"), last = segments[segments.length - 1]!;
  if (asciiKey(segments[segments.length - 2] ?? "") !== "_rels" || !asciiKey(last).endsWith(".rels")) return undefined;
  if (asciiKey(name) === "/_rels/.rels") return "/";
  if (last.length === 5) invalid("Invalid relationship part name.", name);
  return [...segments.slice(0, -2), last.slice(0, -5)].join("/");
}

/** Screen literal IRI-reference spelling without resolving, fetching or normalizing it. */
export function iriReference(value: string, entry: string): void {
  if (!value || /[\u0000-\u0020\u007f-\u009f<>"{}|\\^`]/u.test(value) || /%(?![\da-f]{2})/i.test(value) || value.indexOf("#") !== value.lastIndexOf("#")) invalid("Invalid relationship IRI reference.", entry);
  const first = value.split(/[/?#]/, 1)[0]!;
  if (first.includes(":") && !/^[a-z][a-z\d+.-]*:/i.test(value)) invalid("Invalid IRI scheme or first path segment.", entry);
}
export function resolveInternal(source: string, target: string, entry: string): { name: string; fragment?: string } {
  iriReference(target, entry);
  if (/^[a-z][a-z\d+.-]*:/i.test(target) || target.startsWith("//") || target.split("#", 1)[0]!.includes("?")) invalid("Internal relationship target cannot contain a scheme, authority or query.", entry);
  const hash = target.indexOf("#"), path = hash < 0 ? target : target.slice(0, hash), fragment = hash < 0 ? undefined : target.slice(hash + 1);
  if (!path) {
    if (source === "/") invalid("Package relationship cannot target the package itself.", entry);
    return { name: source, ...(fragment !== undefined ? { fragment } : {}) };
  }
  const segments = path.startsWith("/") ? [] : source.split("/").slice(1, -1);
  const input = path.startsWith("/") ? path.slice(1) : path;
  for (const segment of input.split("/")) {
    if (segment === ".") continue;
    if (segment === "..") { if (!segments.length) invalid("Internal relationship escapes the package root.", entry); segments.pop(); }
    else { if (!segment) invalid("Internal relationship contains an empty path segment.", entry); segments.push(segment); }
  }
  if (/(?:^|\/)\.{1,2}$/.test(path)) invalid("Internal relationship targets a directory rather than a part.", entry);
  const name = partName(`/${segments.join("/")}`, entry);
  return { name, ...(fragment !== undefined ? { fragment } : {}) };
}
