import { PackageInspectionError } from "../types/archive.js";
import type { XmlInspectionOptions } from "../types/xml.js";
import { readXml } from "../xml/read-xml.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { captureXmlTextSource, XmlSourceEditError } from "./xml-text-source.js";

export interface RootChildRemoval {
  readonly elementIndex: number;
  readonly namespaceUri: string;
  readonly local: string;
}

/** Remove explicitly identified raw direct children without reserializing their neighbors. */
export async function removeXmlRootChildren(input: Uint8Array, removals: readonly RootChildRemoval[], rootName: { readonly namespaceUri: string; readonly local: string }, options: XmlInspectionOptions & { readonly maxEdits: number; readonly maxOutputBytes: number }): Promise<Uint8Array> {
  if (!(input instanceof Uint8Array) || !Array.isArray(removals) || !rootName || !options) throw new TypeError("Invalid XML removal request.");
  const { signal, maxEdits, maxOutputBytes, partName } = options;
  const xmlOptions: XmlInspectionOptions = { ...(options.limits ? { limits: { ...options.limits } } : {}), ...(signal ? { signal } : {}), ...(partName !== undefined ? { partName } : {}) };
  if (!Number.isSafeInteger(maxEdits) || maxEdits <= 0 || !Number.isSafeInteger(maxOutputBytes) || maxOutputBytes <= 0) throw new RangeError("Invalid XML removal limit.");
  if (!removals.length || removals.length > maxEdits) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML removal count exceeds the limit.");
  const expectedRoot = { namespaceUri: rootName.namespaceUri, local: rootName.local };
  if (typeof expectedRoot.namespaceUri !== "string" || typeof expectedRoot.local !== "string" || !expectedRoot.local) throw new TypeError("Invalid expected XML root.");
  const captured = removals.map(item => {
    if (!item || typeof item !== "object") throw new TypeError("Invalid XML removal target.");
    const { elementIndex, namespaceUri, local } = item;
    if (!Number.isSafeInteger(elementIndex) || elementIndex <= 1 || typeof namespaceUri !== "string" || typeof local !== "string" || !local) throw new TypeError("Invalid XML removal target.");
    return { elementIndex, namespaceUri, local };
  });
  if (new Set(captured.map(item => item.elementIndex)).size !== captured.length) throw new XmlSourceEditError("xml-edit-conflict", "Duplicate XML removal target.");
  const check = () => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "XML removal was aborted.", partName); };
  check(); if (byteViewInfo(input).shared) throw new TypeError("XML removal requires an ordinary byte buffer.");
  const bytes = copyByteView(input), xml = await captureXmlTextSource(bytes, xmlOptions);
  try {
    check(); const root = xml.elements[0];
    if (!root || root.elementIndex !== 1 || root.parentElementIndex !== undefined || root.namespaceUri !== expectedRoot.namespaceUri || root.local !== expectedRoot.local) throw new XmlSourceEditError("xml-edit-conflict", "XML removal root identity does not match.");
    const spans = captured.map(item => {
      const element = xml.elements[item.elementIndex - 1];
      if (!element || element.elementIndex !== item.elementIndex || element.parentElementIndex !== 1 || element.namespaceUri !== item.namespaceUri || element.local !== item.local) throw new XmlSourceEditError("xml-edit-conflict", "XML removal target identity or ancestry does not match.");
      return { start: element.outerStart, end: element.outerEnd };
    }).sort((a, b) => a.start - b.start);
    const encoding = xml.inspection.encoding, text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
    let cursor = 0; const chunks: string[] = [];
    for (const span of spans) {
      if (span.start < cursor || span.end < span.start || span.end > text.length) throw new XmlSourceEditError("xml-edit-conflict", "XML removal spans overlap or exceed the source.");
      chunks.push(text.slice(cursor, span.start)); cursor = span.end;
    }
    chunks.push(text.slice(cursor)); const candidate = chunks.join("");
    const bomLength = encoding !== "utf-8" ? 2 : bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
    let result: Uint8Array;
    if (encoding === "utf-8") {
      let size = bomLength, work = 0;
      for (const character of candidate) {
        const code = character.codePointAt(0)!;
        size += code < 0x80 ? 1 : code < 0x800 ? 2 : code < 0x10000 ? 3 : 4;
        if (size > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML removal output exceeds the limit.");
        if (++work % 16384 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
      }
      const encoded = new TextEncoder().encode(candidate);
      result = new Uint8Array(size); result.set(bytes.subarray(0, bomLength)); result.set(encoded, bomLength);
    } else {
      const size = bomLength + candidate.length * 2;
      if (size > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML removal output exceeds the limit.");
      result = new Uint8Array(size); result.set(bytes.subarray(0, bomLength));
      for (let i = 0; i < candidate.length; i++) {
        if (i && i % 16384 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
        const unit = candidate.charCodeAt(i), offset = bomLength + 2 * i;
        result[offset + (encoding === "utf-16le" ? 0 : 1)] = unit & 255; result[offset + (encoding === "utf-16le" ? 1 : 0)] = unit >>> 8;
      }
    }
    check(); const verified = await readXml(result, xmlOptions, true); check();
    if (verified.root?.namespaceUri !== expectedRoot.namespaceUri || verified.root.local !== expectedRoot.local) throw new XmlSourceEditError("xml-edit-conflict", "XML removal changed the root identity.");
    return result;
  } finally { xml.dispose(); }
}
