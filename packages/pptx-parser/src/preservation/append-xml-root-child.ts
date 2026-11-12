import type { XmlInspectionOptions, XmlPartInspection } from "../types/xml.js";
import { PackageInspectionError } from "../types/archive.js";
import { readXml, type XmlElement } from "../xml/read-xml.js";
import { XmlSourceEditError } from "./xml-text-source.js";

/** Insert one already-constructed child into a validated root; callers own its semantics. */
export async function appendXmlRootChild(bytes: Uint8Array, root: XmlElement, encoding: XmlPartInspection["encoding"], child: string, maxOutputBytes: number, options: XmlInspectionOptions): Promise<{ bytes: Uint8Array; root: XmlElement }> {
  const check = () => { if (options.signal?.aborted) throw new PackageInspectionError("operation-aborted", "XML insertion was aborted.", options.partName); };
  check();
  const text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
  const bomLength = encoding !== "utf-8" ? 2 : bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  let position = 0, depth = 0, insertion = -1, selfClosing = false, tokens = 0;
  while (position < text.length) {
    const open = text.indexOf("<", position); if (open < 0) break;
    if (++tokens % 256 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
    if (text.startsWith("<!--", open)) { position = text.indexOf("-->", open + 4) + 3; continue; }
    if (text.startsWith("<![CDATA[", open)) { position = text.indexOf("]]>", open + 9) + 3; continue; }
    if (text.startsWith("<?", open)) { position = text.indexOf("?>", open + 2) + 2; continue; }
    const closing = text[open + 1] === "/"; let end = open + 1, quote = "";
    for (; end < text.length; end++) { const c = text[end]!; if (quote) { if (c === quote) quote = ""; } else if (c === "'" || c === '"') quote = c; else if (c === ">") break; }
    if (end >= text.length) throw new XmlSourceEditError("xml-edit-unsupported", "XML root lexical boundary is unavailable.");
    if (closing) { if (--depth === 0) { insertion = open; break; } }
    else if (text[end - 1] === "/") { if (depth === 0) { insertion = end - 1; selfClosing = true; break; } }
    else depth++;
    position = end + 1;
  }
  if (insertion < 0) throw new XmlSourceEditError("xml-edit-unsupported", "XML root lexical boundary is unavailable.");
  const added = selfClosing ? `>${child}</${root.name}>` : child, removed = selfClosing ? 2 : 0;
  const outputSize = bytes.length + (encoding === "utf-8" ? new TextEncoder().encode(added).length - removed : (added.length - removed) * 2);
  if (outputSize > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML candidate exceeds maxOutputBytes.");
  const candidate = text.slice(0, insertion) + added + text.slice(insertion + removed);
  let output: Uint8Array;
  if (encoding === "utf-8") { const encoded = new TextEncoder().encode(candidate); output = new Uint8Array(encoded.length + bomLength); output.set(bytes.subarray(0, bomLength)); output.set(encoded, bomLength); }
  else { output = new Uint8Array(outputSize); output.set(bytes.subarray(0, bomLength)); for (let i = 0; i < candidate.length; i++) { const value = candidate.charCodeAt(i), offset = bomLength + 2 * i; output[offset + (encoding === "utf-16le" ? 0 : 1)] = value & 255; output[offset + (encoding === "utf-16le" ? 1 : 0)] = value >>> 8; } }
  check(); const verified = await readXml(output, options, true); check();
  return { bytes: output, root: verified.root! };
}
