import { SaxesParser } from "saxes";
import { PackageInspectionError } from "../types/archive.js";
import type { XmlExpandedName, XmlInspectionLimits, XmlInspectionOptions, XmlPartInspection } from "../types/xml.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";

export const DEFAULT_XML_LIMITS: Readonly<XmlInspectionLimits> = Object.freeze({ maxInputBytes: 8 * 1024 * 1024, maxCharacters: 4 * 1024 * 1024, maxNodes: 100000, maxDepth: 256, maxAttributesPerElement: 256, maxTextCharacters: 2 * 1024 * 1024, maxPendingCharacters: 128 * 1024 });
export interface XmlAttribute extends XmlExpandedName { readonly value: string }
export interface XmlElement extends XmlExpandedName { readonly kind: "element"; readonly attributes: readonly XmlAttribute[]; readonly children: XmlNode[] }
export type XmlNode = XmlElement | { kind: "text" | "cdata" | "comment"; value: string } | { kind: "processing-instruction"; target: string; value: string };
/** Includes prolog/epilog content; format readers ordinarily consume only root. */
export interface XmlDocument { readonly children: XmlNode[]; readonly standalone?: string }

function limitsFor(options: XmlInspectionOptions): XmlInspectionLimits {
  if (!options || typeof options !== "object" || Array.isArray(options) || (options.limits !== undefined && (!options.limits || typeof options.limits !== "object" || Array.isArray(options.limits)))) throw new TypeError("Invalid XML inspection options");
  if (options.partName !== undefined && typeof options.partName !== "string") throw new TypeError("XML partName must be a string");
  const limits = { ...DEFAULT_XML_LIMITS, ...options.limits };
  for (const [name, value] of Object.entries(limits)) if (!Number.isSafeInteger(value) || value <= 0) throw new RangeError(`Invalid XML limit ${name}`);
  return limits;
}
function detectEncoding(bytes: Uint8Array, part: string | undefined): XmlPartInspection["encoding"] {
  if ((bytes[0] === 0 && bytes[1] === 0 && bytes[2] === 0xfe && bytes[3] === 0xff) || (bytes[0] === 0xff && bytes[1] === 0xfe && bytes[2] === 0 && bytes[3] === 0)) throw new PackageInspectionError("xml-unsupported", "UTF-32 XML is outside the supported encoding profile.", part);
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return "utf-16le";
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return "utf-16be";
  if ((bytes[0] === 0 && bytes[1] === 0x3c) || (bytes[0] === 0x3c && bytes[1] === 0)) throw new PackageInspectionError("xml-unsupported", "UTF-16 XML requires a byte-order mark in this profile.", part);
  return "utf-8";
}

/** Internal bounded XML reader. No AST or library parser instance crosses the public API. */
export async function readXml(input: Uint8Array, options: XmlInspectionOptions = {}, retain = false, observe?: (metric: "characters" | "nodes" | "textCharacters", increment: number) => void): Promise<{ inspection: XmlPartInspection; root: XmlElement | undefined; document: XmlDocument | undefined; declaredEncoding: string | undefined }> {
  const limits = limitsFor(options), signal = options.signal, part = options.partName;
  const abort = (): void => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "XML inspection was aborted.", part); };
  if (!(input instanceof Uint8Array)) throw new TypeError("XML input must be Uint8Array");
  abort();
  const view = byteViewInfo(input);
  if (view.shared) throw new TypeError("XML input must not use SharedArrayBuffer.");
  if (view.length > limits.maxInputBytes) throw new PackageInspectionError("xml-limit-exceeded", "XML input exceeds maxInputBytes.", part);
  const bytes = copyByteView(input), encoding = detectEncoding(bytes, part);
  const decoder = new TextDecoder(encoding, { fatal: true });
  const parser = new SaxesParser({ xmlns: true, defaultXMLVersion: "1.0", forceXMLVersion: true });
  let characters = 0, nodes = 0, elements = 0, attributes = 0, depth = 0, maxDepth = 0, textCharacters = 0, tagAttributes = 0, lastEvent = 0;
  let lastKind: "text" | "other" = "other";
  let rootName: XmlExpandedName | undefined, root: XmlElement | undefined, declaredEncoding: string | undefined;
  const namespaces = new Set<string>(), stack: XmlElement[] = [];
  const documentChildren: XmlNode[] = [];
  let standalone: string | undefined;
  const fail = (code: "xml-invalid" | "xml-unsupported" | "xml-limit-exceeded", message: string, offset = parser.position): never => {
    throw new PackageInspectionError(code, message, part, { line: parser.line, column: parser.column + 1, offset });
  };
  const pending = (offset = parser.position): void => { if (offset - lastEvent > limits.maxPendingCharacters) fail("xml-limit-exceeded", "XML source span exceeds maxPendingCharacters.", offset); };
  const event = (): void => { pending(); lastEvent = parser.position; };
  const node = (): void => { observe?.("nodes", 1); if (++nodes > limits.maxNodes) fail("xml-limit-exceeded", "XML node count exceeds maxNodes."); };
  const text = (value: string): void => { observe?.("textCharacters", value.length); textCharacters += value.length; if (textCharacters > limits.maxTextCharacters) fail("xml-limit-exceeded", "XML content exceeds maxTextCharacters."); };
  const namespace = (uri: string): void => { if (uri) namespaces.add(uri); };
  const append = (value: XmlNode): void => {
    if (retain) (stack.length ? stack[stack.length - 1]!.children : documentChildren).push(value);
  };
  parser.on("error", (error) => fail("xml-invalid", `Malformed XML: ${error.message.slice(0, 300)}`));
  parser.on("doctype", () => fail("xml-unsupported", "DTD declarations and user-defined entities are not supported."));
  parser.on("xmldecl", (declaration) => {
    event(); declaredEncoding = declaration.encoding;
    standalone = declaration.standalone;
    if (declaration.version !== "1.0") fail("xml-unsupported", "Only XML 1.0 is supported.");
    if (declaration.encoding) {
      const declared = declaration.encoding.toLowerCase();
      if (declared !== encoding && !(declared === "utf-16" && encoding !== "utf-8")) fail("xml-unsupported", "XML encoding declaration disagrees with the supported byte encoding.");
    }
  });
  parser.on("opentagstart", () => {
    event(); lastKind = "other"; tagAttributes = 0; node(); elements++; depth++; maxDepth = Math.max(maxDepth, depth);
    if (depth > limits.maxDepth) fail("xml-limit-exceeded", "XML nesting exceeds maxDepth.");
  });
  parser.on("attribute", (attribute) => {
    event(); attributes++; text(attribute.value);
    if (++tagAttributes > limits.maxAttributesPerElement) fail("xml-limit-exceeded", "XML attributes exceed maxAttributesPerElement.");
  });
  parser.on("opentag", (tag) => {
    event();
    const name: XmlExpandedName = { name: tag.name, prefix: tag.prefix, local: tag.local, namespaceUri: tag.uri };
    rootName ??= name;
    namespace(tag.uri);
    for (const attribute of Object.values(tag.attributes)) { namespace(attribute.uri); if (attribute.uri === "http://www.w3.org/2000/xmlns/") namespace(attribute.value); }
    if (retain) {
      const element: XmlElement = { ...name, kind: "element", attributes: Object.values(tag.attributes).map((attribute) => ({ name: attribute.name, prefix: attribute.prefix, local: attribute.local, namespaceUri: attribute.uri, value: attribute.value })), children: [] };
      append(element); root ??= element; stack.push(element);
    }
  });
  parser.on("closetag", () => { event(); lastKind = "other"; depth--; if (retain) stack.pop(); });
  parser.on("text", (value) => {
    event(); text(value);
    if (lastKind !== "text") node();
    lastKind = "text";
    if (retain) {
      const children = stack.length ? stack[stack.length - 1]!.children : documentChildren, previous = children[children.length - 1];
      if (previous?.kind === "text") previous.value += value;
      else append({ kind: "text", value });
    }
  });
  parser.on("cdata", (value) => { event(); lastKind = "other"; node(); text(value); append({ kind: "cdata", value }); });
  parser.on("comment", (value) => { event(); lastKind = "other"; node(); text(value); append({ kind: "comment", value }); });
  parser.on("processinginstruction", ({ target, body }) => { event(); lastKind = "other"; node(); text(body); append({ kind: "processing-instruction", target, value: body }); });
  const feed = (chunk: string): void => {
    if (characters === 0 && chunk.startsWith("\ufeff")) fail("xml-invalid", "XML contains a repeated leading byte-order mark.", 0);
    observe?.("characters", chunk.length);
    characters += chunk.length;
    if (characters > limits.maxCharacters) fail("xml-limit-exceeded", "XML source exceeds maxCharacters.", characters - chunk.length);
    parser.write(chunk);
    // saxes 6 position is an event-time cursor; after write it includes the
    // just-consumed chunk twice. Our decoded source count is authoritative here.
    pending(characters);
  };
  // Snapshot input/options before yielding; return nothing until parser.close succeeds.
  await Promise.resolve(); abort();
  try {
    let chunks = 0;
    for (let offset = 0; offset < bytes.length; offset += 1024) {
      abort();
      feed(decoder.decode(bytes.subarray(offset, Math.min(offset + 1024, bytes.length)), { stream: true }));
      if (++chunks % 32 === 0) { await new Promise<void>((resolve) => setTimeout(resolve, 0)); abort(); }
    }
    feed(decoder.decode()); parser.close(); abort();
  } catch (error) {
    if (error instanceof PackageInspectionError) throw error;
    fail("xml-invalid", "XML decoding or parsing failed.", characters);
  }
  if (!rootName) fail("xml-invalid", "XML document has no root element.");
  return { inspection: { root: rootName!, encoding, inputBytes: bytes.length, characters, nodes, elements, attributes, maxDepth, textCharacters, namespaceUris: [...namespaces].sort() }, root, document: retain ? { children: documentChildren, ...(standalone !== undefined ? { standalone } : {}) } : undefined, declaredEncoding };
}
