import { PackageInspectionError } from "../types/archive.js";
import type { XmlInspectionOptions } from "../types/xml.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { DEFAULT_XML_LIMITS, readXml, type XmlElement } from "../xml/read-xml.js";
import { XmlSourceEditError } from "./xml-text-source.js";

interface AttributeReplacement {
  readonly elementIndex: number; readonly elementNamespace: string; readonly elementLocal: string;
  readonly attributeNamespace: string; readonly attributeLocal: string; readonly expected: string; readonly replacement: string;
}
type OptionalAttributeReplacement = Omit<AttributeReplacement, "expected" | "replacement"> & { readonly expected?: string; readonly replacement?: string };
type AttributeOptions = XmlInspectionOptions & { readonly maxOutputBytes: number };

/** Private lexical primitive: neither wrapper establishes ownership or save permission. */
export async function replaceXmlRelationshipId(input: Uint8Array, request: AttributeReplacement, options: AttributeOptions): Promise<Uint8Array> {
  return replaceXmlAttribute(input, request, options, "relationship");
}

/** Replace an existing attribute with a canonical nonnegative signed-32-bit value. */
export async function replaceXmlNumericAttribute(input: Uint8Array, request: AttributeReplacement, options: AttributeOptions): Promise<Uint8Array> {
  return replaceXmlAttribute(input, request, options, "number");
}

export async function editXmlNumericAttribute(input: Uint8Array, request: OptionalAttributeReplacement, options: AttributeOptions): Promise<Uint8Array> {
  return replaceXmlAttribute(input, request, options, "number-optional");
}

/** Private character-valued attribute primitive; it grants no source edit authority. */
export async function editXmlTextAttribute(input: Uint8Array, request: OptionalAttributeReplacement, options: AttributeOptions): Promise<Uint8Array> {
  return replaceXmlAttribute(input, request, options, "text-optional");
}

function validXmlCharacters(value: string): boolean {
  for (const character of value) {
    const code = character.codePointAt(0)!;
    if (code !== 9 && code !== 10 && code !== 13 && (code < 32 || code >= 0xD800 && code <= 0xDFFF || code === 0xFFFE || code === 0xFFFF)) return false;
  }
  return true;
}

function escapeAttribute(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&apos;').replaceAll('\t', '&#x9;').replaceAll('\n', '&#xA;').replaceAll('\r', '&#xD;');
}

async function checkEscapedAttributeBudget(value: string, encoding: string, maxBytes: number, check: () => void): Promise<void> {
  let bytes = 0, work = 0;
  for (const character of value) {
    const code = character.codePointAt(0)!;
    const escapedLength = character === '&' ? 5 : character === '<' || character === '>' ? 4 : character === '"' || character === "'" ? 6 : character === '\t' || character === '\n' || character === '\r' ? 5 : 0;
    bytes += escapedLength ? escapedLength * (encoding === 'utf-8' ? 1 : 2) : encoding !== 'utf-8' ? character.length * 2 : code <= 0x7F ? 1 : code <= 0x7FF ? 2 : code <= 0xFFFF ? 3 : 4;
    if (bytes > maxBytes) throw new XmlSourceEditError('xml-edit-limit-exceeded', 'Escaped XML attribute exceeds maxOutputBytes.');
    if (++work % 65536 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
  }
}

async function replaceXmlAttribute(input: Uint8Array, request: OptionalAttributeReplacement, options: AttributeOptions, syntax: "relationship" | "number" | "number-optional" | "text-optional"): Promise<Uint8Array> {
  if (!(input instanceof Uint8Array) || !request || typeof request !== "object" || !options || typeof options !== "object") throw new TypeError("Invalid XML relationship edit input.");
  const { elementIndex, elementNamespace, elementLocal, attributeNamespace, attributeLocal, expected, replacement } = request;
  const { signal, partName, maxOutputBytes } = options, limits = { ...DEFAULT_XML_LIMITS, ...options.limits };
  if (!Number.isSafeInteger(elementIndex) || elementIndex <= 0 || [elementNamespace, elementLocal, attributeNamespace, attributeLocal].some(value => typeof value !== "string") || (expected !== undefined && typeof expected !== "string") || (replacement !== undefined && typeof replacement !== "string") || (syntax !== "number-optional" && syntax !== "text-optional" && (expected === undefined || replacement === undefined))) throw new TypeError("Invalid XML relationship edit request.");
  if (!Number.isSafeInteger(maxOutputBytes) || maxOutputBytes <= 0) throw new RangeError("Invalid XML relationship edit output limit.");
  if (replacement !== undefined && replacement.length > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML relationship replacement exceeds maxOutputBytes.");
  // Each wrapper admits a small ASCII grammar. Existing values may use
  // entities; comparison uses the XML reader's normalized attribute value.
  const valid = replacement === undefined ? syntax === "number-optional" || syntax === "text-optional" : syntax === "text-optional" ? validXmlCharacters(replacement) : syntax === "relationship" ? /^[A-Za-z_][A-Za-z0-9_.-]*$/.test(replacement) : /^(?:0|[1-9][0-9]*)$/.test(replacement) && Number.isSafeInteger(Number(replacement)) && Number(replacement) <= 2147483647;
  if (expected === undefined && (attributeNamespace !== "" || !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(attributeLocal))) throw new XmlSourceEditError("xml-edit-unsupported", "Attribute insertion requires an unqualified name.");
  if (!valid || attributeNamespace === "http://www.w3.org/2000/xmlns/") throw new XmlSourceEditError("xml-edit-unsupported", "Attribute replacement requires an admitted lexical value and a non-namespace attribute.");
  const check = () => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "XML relationship editing was aborted.", partName); };
  check();
  const view = byteViewInfo(input);
  if (view.shared) throw new TypeError("XML relationship editing requires an ordinary byte buffer.");
  if (!Number.isSafeInteger(limits.maxInputBytes) || limits.maxInputBytes <= 0) throw new RangeError("Invalid XML limit maxInputBytes.");
  if (view.length > limits.maxInputBytes) throw new PackageInspectionError("xml-limit-exceeded", "XML source exceeds maxInputBytes.", partName);
  const bytes = copyByteView(input), captured = { limits, ...(partName !== undefined ? { partName } : {}), ...(signal ? { signal } : {}) };
  const parsed = await readXml(bytes, captured, true); check();
  const elements: XmlElement[] = [], stack = [parsed.root!];
  while (stack.length) {
    const node = stack.pop()!; elements.push(node);
    for (let i = node.children.length - 1; i >= 0; i--) { const child = node.children[i]!; if (child.kind === "element") stack.push(child); }
  }
  const target = elements[elementIndex - 1];
  if (!target || target.namespaceUri !== elementNamespace || target.local !== elementLocal) throw new XmlSourceEditError("xml-edit-conflict", "Relationship element identity does not match the source.");
  const attribute = target.attributes.find(item => item.namespaceUri === attributeNamespace && item.local === attributeLocal);
  if (attribute?.value !== expected || (expected === undefined && attribute !== undefined)) throw new XmlSourceEditError("xml-edit-conflict", "Relationship attribute or expected value does not match the source.");
  if (replacement === expected) {
    if (bytes.length > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML relationship candidate exceeds maxOutputBytes.");
    return bytes;
  }
  const encoding = parsed.inspection.encoding, text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
  const bomLength = encoding !== "utf-8" ? 2 : bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  const whitespace = (c: string | undefined) => c === " " || c === "\t" || c === "\r" || c === "\n";
  let position = 0, index = 0, start = -1, end = -1;
  while (position < text.length) {
    const open = text.indexOf("<", position); if (open < 0) break;
    if (text.startsWith("<!--", open)) { position = text.indexOf("-->", open + 4) + 3; continue; }
    if (text.startsWith("<![CDATA[", open)) { position = text.indexOf("]]>", open + 9) + 3; continue; }
    if (text.startsWith("<?", open)) { position = text.indexOf("?>", open + 2) + 2; continue; }
    const closing = text[open + 1] === "/"; let cursor = open + (closing ? 2 : 1), nameStart = cursor;
    while (cursor < text.length && !whitespace(text[cursor]) && text[cursor] !== "/" && text[cursor] !== ">") cursor++;
    const name = text.slice(nameStart, cursor);
    if (closing) { position = text.indexOf(">", cursor) + 1; continue; }
    if (elements[index]?.name !== name) throw new XmlSourceEditError("xml-edit-unsupported", "XML source preorder does not match its validated tree.");
    index++;
    while (cursor < text.length) {
      while (whitespace(text[cursor])) cursor++;
      if (text[cursor] === ">") { if (index === elementIndex && !attribute) start = end = cursor; cursor++; break; }
      if (text[cursor] === "/" && text[cursor + 1] === ">") { if (index === elementIndex && !attribute) start = end = cursor; cursor += 2; break; }
      nameStart = cursor;
      while (cursor < text.length && !whitespace(text[cursor]) && text[cursor] !== "=") cursor++;
      const attributeName = text.slice(nameStart, cursor);
      while (whitespace(text[cursor])) cursor++;
      if (text[cursor++] !== "=") throw new XmlSourceEditError("xml-edit-unsupported", "XML attribute source could not be bound.");
      while (whitespace(text[cursor])) cursor++;
      const quote = text[cursor++];
      if (quote !== "'" && quote !== '"') throw new XmlSourceEditError("xml-edit-unsupported", "XML attribute quote could not be bound.");
      const valueStart = cursor, valueEnd = text.indexOf(quote, cursor);
      if (valueEnd < 0) throw new XmlSourceEditError("xml-edit-unsupported", "XML attribute value could not be bound.");
      if (index === elementIndex && attributeName === attribute?.name) { start = replacement === undefined ? nameStart : valueStart; end = replacement === undefined ? valueEnd + 1 : valueEnd; }
      cursor = valueEnd + 1;
    }
    position = cursor;
    if (index === elementIndex) break;
    if (index % 256 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
  }
  if (start < 0 || end < start) throw new XmlSourceEditError("xml-edit-unsupported", "XML relationship attribute has no verified source span.");
  if (replacement !== undefined && syntax === "text-optional") await checkEscapedAttributeBudget(replacement, encoding, maxOutputBytes, check);
  const lexicalValue = replacement === undefined ? undefined : syntax === "text-optional" ? escapeAttribute(replacement) : replacement;
  const edit = lexicalValue === undefined ? "" : attribute ? lexicalValue : ` ${attributeLocal}="${lexicalValue}"`;
  const outputSize = encoding === "utf-8" ? bytes.length - new TextEncoder().encode(text.slice(start, end)).length + new TextEncoder().encode(edit).length : bytes.length - (end - start) * 2 + edit.length * 2;
  if (outputSize > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML relationship candidate exceeds maxOutputBytes.");
  const candidate = text.slice(0, start) + edit + text.slice(end);
  let output: Uint8Array;
  if (encoding === "utf-8") {
    const encoded = new TextEncoder().encode(candidate);
    if (encoded.length + bomLength > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML relationship candidate exceeds maxOutputBytes.");
    output = new Uint8Array(encoded.length + bomLength); output.set(bytes.subarray(0, bomLength)); output.set(encoded, bomLength);
  } else {
    const size = bomLength + candidate.length * 2;
    if (size > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML relationship candidate exceeds maxOutputBytes.");
    output = new Uint8Array(size); output.set(bytes.subarray(0, bomLength));
    for (let i = 0; i < candidate.length; i++) { const value = candidate.charCodeAt(i), offset = bomLength + i * 2; output[offset + (encoding === "utf-16le" ? 0 : 1)] = value & 255; output[offset + (encoding === "utf-16le" ? 1 : 0)] = value >>> 8; }
  }
  check();
  const verified = await readXml(output, captured, true); check();
  const verifyStack = [verified.root!]; let verifiedTarget: XmlElement | undefined, count = 0;
  while (verifyStack.length) { const node = verifyStack.pop()!; if (++count === elementIndex) verifiedTarget = node; for (let i = node.children.length - 1; i >= 0; i--) { const child = node.children[i]!; if (child.kind === "element") verifyStack.push(child); } }
  if (count !== elements.length || verifiedTarget?.namespaceUri !== elementNamespace || verifiedTarget.local !== elementLocal || verifiedTarget.attributes.find(item => item.namespaceUri === attributeNamespace && item.local === attributeLocal)?.value !== replacement) throw new XmlSourceEditError("xml-edit-unsupported", "XML relationship candidate failed identity verification.");
  return output;
}

/** Atomically patch multiple independently bound character attributes with two XML parses. */
export async function editXmlTextAttributes(
  input: Uint8Array, requests: readonly OptionalAttributeReplacement[], options: AttributeOptions & { readonly maxEdits: number; readonly maxWorkUnits: number },
): Promise<{ bytes: Uint8Array; workUnits: number }> {
  if (!(input instanceof Uint8Array) || !Array.isArray(requests) || !options || typeof options !== 'object') throw new TypeError('Invalid XML attribute batch input.');
  const { signal, partName, maxOutputBytes, maxEdits, maxWorkUnits } = options, limits = { ...DEFAULT_XML_LIMITS, ...options.limits };
  if (![maxOutputBytes, maxEdits, maxWorkUnits, limits.maxInputBytes].every(v => Number.isSafeInteger(v) && v > 0)) throw new RangeError('Invalid XML attribute batch limit.');
  if (!requests.length || requests.length > maxEdits) throw new XmlSourceEditError('xml-edit-limit-exceeded', 'XML attribute batch exceeds maxEdits.');
  const check = () => { if (signal?.aborted) throw new PackageInspectionError('operation-aborted', 'XML attribute batch was aborted.', partName); };
  check(); const view = byteViewInfo(input); if (view.shared) throw new TypeError('XML attribute batch requires an ordinary byte buffer.');
  if (view.length > limits.maxInputBytes) throw new PackageInspectionError('xml-limit-exceeded', 'XML source exceeds maxInputBytes.', partName);
  const bytes = copyByteView(input), captured = { limits, ...(partName !== undefined ? { partName } : {}), ...(signal ? { signal } : {}) };
  const batch = requests.map(request => {
    if (!request || typeof request !== 'object') throw new TypeError('Invalid XML attribute batch request.');
    const { elementIndex, elementNamespace, elementLocal, attributeNamespace, attributeLocal, expected, replacement } = request;
    if (!Number.isSafeInteger(elementIndex) || elementIndex <= 0 || [elementNamespace, elementLocal, attributeNamespace, attributeLocal].some(v => typeof v !== 'string') || expected !== undefined && typeof expected !== 'string' || replacement !== undefined && typeof replacement !== 'string') throw new TypeError('Invalid XML attribute batch request.');
    const insertXmlSpace = attributeNamespace === 'http://www.w3.org/XML/1998/namespace' && attributeLocal === 'space';
    if (attributeNamespace === 'http://www.w3.org/2000/xmlns/' || expected === undefined && (!insertXmlSpace && attributeNamespace !== '' || !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(attributeLocal))) throw new XmlSourceEditError('xml-edit-unsupported', 'Attribute batch has unsupported names or characters.');
    if (insertXmlSpace && replacement !== undefined && replacement !== 'default' && replacement !== 'preserve') throw new XmlSourceEditError('xml-edit-unsupported', 'xml:space accepts only default or preserve.');
    if (replacement !== undefined && replacement.length > maxOutputBytes) throw new XmlSourceEditError('xml-edit-limit-exceeded', 'XML attribute replacement exceeds maxOutputBytes.');
    return { elementIndex, elementNamespace, elementLocal, attributeNamespace, attributeLocal, expected, replacement };
  });
  let workUnits = 0;
  const charge = (units: number) => { workUnits += units; if (!Number.isSafeInteger(workUnits) || workUnits > maxWorkUnits) throw new XmlSourceEditError('xml-edit-limit-exceeded', 'XML attribute batch exceeds maxWorkUnits.'); };
  for (const request of batch) {
    let visited = 0;
    for (const character of request.replacement ?? '') {
      const code = character.codePointAt(0)!;
      if (code !== 9 && code !== 10 && code !== 13 && (code < 32 || code >= 0xD800 && code <= 0xDFFF || code === 0xFFFE || code === 0xFFFF)) throw new XmlSourceEditError('xml-edit-unsupported', 'Attribute batch contains unsupported XML characters.');
      if (++visited % 16384 === 0) { charge(16384); await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
    }
    charge(visited % 16384);
  }
  charge(bytes.length); const parsed = await readXml(bytes, captured, true); check();
  const elements: XmlElement[] = [], stack = [parsed.root!];
  while (stack.length) { const node = stack.pop()!; elements.push(node); for (let i = node.children.length - 1; i >= 0; i--) { const child = node.children[i]!; if (child.kind === 'element') stack.push(child); } }
  type BatchTarget = { request: typeof batch[number]; attributeName?: string; start?: number; end?: number };
  const byIndex = new Map<number, BatchTarget[]>(), byQualifiedName = new Map<number, Map<string, BatchTarget>>(), attributesByIndex = new Map<number, Map<string, XmlElement['attributes'][number]>>(), identities = new Set<string>();
  for (const request of batch) {
    check(); charge(request.elementNamespace.length + request.elementLocal.length + request.attributeNamespace.length + request.attributeLocal.length + (request.expected?.length ?? 0) + (request.replacement?.length ?? 0) + 1);
    const identity = JSON.stringify([request.elementIndex, request.attributeNamespace, request.attributeLocal]); if (identities.has(identity)) throw new XmlSourceEditError('xml-edit-conflict', 'Duplicate XML attribute batch target.'); identities.add(identity);
    const target = elements[request.elementIndex - 1]; if (!target || target.namespaceUri !== request.elementNamespace || target.local !== request.elementLocal) throw new XmlSourceEditError('xml-edit-conflict', 'XML attribute batch element identity does not match.');
    let attributes = attributesByIndex.get(request.elementIndex);
    if (!attributes) { attributes = new Map(); for (const a of target.attributes) { charge(a.name.length + a.value.length + 1); attributes.set(JSON.stringify([a.namespaceUri, a.local]), a); } attributesByIndex.set(request.elementIndex, attributes); }
    const attribute = attributes.get(JSON.stringify([request.attributeNamespace, request.attributeLocal]));
    if (attribute?.value !== request.expected || request.expected === undefined && attribute !== undefined) throw new XmlSourceEditError('xml-edit-conflict', 'XML attribute batch value or presence does not match.');
    if (request.expected === request.replacement) continue;
    const list = byIndex.get(request.elementIndex) ?? [], item: BatchTarget = { request, ...(attribute ? { attributeName: attribute.name } : {}) }; list.push(item); byIndex.set(request.elementIndex, list);
    if (attribute) { const names = byQualifiedName.get(request.elementIndex) ?? new Map<string, BatchTarget>(); names.set(attribute.name, item); byQualifiedName.set(request.elementIndex, names); }
  }
  if (!byIndex.size) { if (bytes.length > maxOutputBytes) throw new XmlSourceEditError('xml-edit-limit-exceeded', 'XML candidate exceeds maxOutputBytes.'); return { bytes, workUnits }; }
  const encoding = parsed.inspection.encoding, text = new TextDecoder(encoding, { fatal: true }).decode(bytes), bomLength = encoding !== 'utf-8' ? 2 : bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  const whitespace = (c: string | undefined) => c === ' ' || c === '\t' || c === '\r' || c === '\n';
  let position = 0, index = 0;
  while (position < text.length) {
    const open = text.indexOf('<', position); if (open < 0) break;
    if (text.startsWith('<!--', open)) { position = text.indexOf('-->', open + 4) + 3; continue; }
    if (text.startsWith('<![CDATA[', open)) { position = text.indexOf(']]>', open + 9) + 3; continue; }
    if (text.startsWith('<?', open)) { position = text.indexOf('?>', open + 2) + 2; continue; }
    const closing = text[open + 1] === '/'; let cursor = open + (closing ? 2 : 1), start = cursor;
    while (cursor < text.length && !whitespace(text[cursor]) && text[cursor] !== '/' && text[cursor] !== '>') cursor++;
    const name = text.slice(start, cursor); if (closing) { position = text.indexOf('>', cursor) + 1; continue; }
    if (elements[index]?.name !== name) throw new XmlSourceEditError('xml-edit-unsupported', 'XML batch source preorder disagrees with the parsed tree.');
    index++; const targets = byIndex.get(index);
    while (cursor < text.length) {
      while (whitespace(text[cursor])) cursor++;
      if (text[cursor] === '>' || text[cursor] === '/' && text[cursor + 1] === '>') { if (targets) for (const target of targets) if (!target.attributeName) target.start = target.end = cursor; cursor += text[cursor] === '/' ? 2 : 1; break; }
      start = cursor; while (cursor < text.length && !whitespace(text[cursor]) && text[cursor] !== '=') cursor++;
      const attributeName = text.slice(start, cursor); while (whitespace(text[cursor])) cursor++; if (text[cursor++] !== '=') throw new XmlSourceEditError('xml-edit-unsupported', 'XML batch attribute could not be bound.');
      while (whitespace(text[cursor])) cursor++; const quote = text[cursor++]; if (quote !== "'" && quote !== '"') throw new XmlSourceEditError('xml-edit-unsupported', 'XML batch quote could not be bound.');
      const valueStart = cursor, valueEnd = text.indexOf(quote, cursor); if (valueEnd < 0) throw new XmlSourceEditError('xml-edit-unsupported', 'XML batch value could not be bound.');
      const target = byQualifiedName.get(index)?.get(attributeName); if (target) { target.start = target.request.replacement === undefined ? start : valueStart; target.end = target.request.replacement === undefined ? valueEnd + 1 : valueEnd; }
      cursor = valueEnd + 1;
    }
    position = cursor; if (index % 256 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
  }
  charge(bytes.length); const patches: { start: number; end: number; edit: string }[] = [];
  for (const targets of byIndex.values()) for (const target of targets) {
    check(); if (target.start === undefined || target.end === undefined) throw new XmlSourceEditError('xml-edit-unsupported', 'XML batch attribute has no verified source span.');
    const replacement = target.request.replacement; if (replacement !== undefined) await checkEscapedAttributeBudget(replacement, encoding, maxOutputBytes, check);
    const lexical = replacement === undefined ? '' : escapeAttribute(replacement), name = target.request.attributeNamespace === 'http://www.w3.org/XML/1998/namespace' ? `xml:${target.request.attributeLocal}` : target.request.attributeLocal;
    const edit = replacement !== undefined && !target.attributeName ? ` ${name}="${lexical}"` : lexical;
    charge(edit.length * (encoding === 'utf-8' ? 4 : 2) + 1); patches.push({ start: target.start, end: target.end, edit });
  }
  charge(patches.length * Math.ceil(Math.log2(patches.length + 1)));
  patches.sort((a, b) => a.start - b.start || a.end - b.end);
  const segments: string[] = [], encoder = new TextEncoder(); let cursor = 0, outputSize = bytes.length;
  for (const patch of patches) { if (patch.start < cursor) throw new XmlSourceEditError('xml-edit-conflict', 'XML batch source spans overlap.'); outputSize += encoding === 'utf-8' ? encoder.encode(patch.edit).length - encoder.encode(text.slice(patch.start, patch.end)).length : (patch.edit.length - (patch.end - patch.start)) * 2; segments.push(text.slice(cursor, patch.start), patch.edit); cursor = patch.end; }
  if (outputSize > maxOutputBytes) throw new XmlSourceEditError('xml-edit-limit-exceeded', 'XML batch candidate exceeds maxOutputBytes.');
  segments.push(text.slice(cursor)); const candidate = segments.join('');
  let output: Uint8Array;
  if (encoding === 'utf-8') { const encoded = new TextEncoder().encode(candidate); if (encoded.length + bomLength > maxOutputBytes) throw new XmlSourceEditError('xml-edit-limit-exceeded', 'XML batch candidate exceeds maxOutputBytes.'); output = new Uint8Array(encoded.length + bomLength); output.set(bytes.subarray(0, bomLength)); output.set(encoded, bomLength); }
  else { const size = bomLength + candidate.length * 2; if (size > maxOutputBytes) throw new XmlSourceEditError('xml-edit-limit-exceeded', 'XML batch candidate exceeds maxOutputBytes.'); output = new Uint8Array(size); output.set(bytes.subarray(0, bomLength)); for (let i = 0; i < candidate.length; i++) { const value = candidate.charCodeAt(i), offset = bomLength + i * 2; output[offset + (encoding === 'utf-16le' ? 0 : 1)] = value & 255; output[offset + (encoding === 'utf-16le' ? 1 : 0)] = value >>> 8; } }
  charge(output.length); check(); const verified = await readXml(output, captured, true); check();
  const verifiedElements: XmlElement[] = [], pending = [verified.root!];
  while (pending.length) { const node = pending.pop()!; verifiedElements.push(node); for (let i = node.children.length - 1; i >= 0; i--) { const child = node.children[i]!; if (child.kind === 'element') pending.push(child); } }
  if (verifiedElements.length !== elements.length) throw new XmlSourceEditError('xml-edit-unsupported', 'XML batch candidate changed element topology.');
  const verifiedAttributes = new Map<number, Map<string, string>>();
  let verifiedCount = 0;
  for (const request of batch) {
    if (++verifiedCount % 256 === 0) { await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
    const target = verifiedElements[request.elementIndex - 1]; if (!target || target.namespaceUri !== request.elementNamespace || target.local !== request.elementLocal) throw new XmlSourceEditError('xml-edit-unsupported', 'XML batch candidate failed element verification.');
    let attributes = verifiedAttributes.get(request.elementIndex); if (!attributes) { attributes = new Map(); for (const a of target.attributes) { charge(a.name.length + a.value.length + 1); attributes.set(JSON.stringify([a.namespaceUri, a.local]), a.value); } verifiedAttributes.set(request.elementIndex, attributes); }
    if (attributes.get(JSON.stringify([request.attributeNamespace, request.attributeLocal])) !== request.replacement) throw new XmlSourceEditError('xml-edit-unsupported', 'XML batch candidate failed attribute verification.');
  }
  return { bytes: output, workUnits };
}
