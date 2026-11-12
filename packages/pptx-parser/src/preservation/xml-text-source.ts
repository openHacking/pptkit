import { PackageInspectionError } from "../types/archive.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import type { XmlInspectionOptions, XmlPartInspection } from "../types/xml.js";
import { DEFAULT_XML_LIMITS, readXml, type XmlElement } from "../xml/read-xml.js";
import { splitAtGraphemeBoundary } from "./split-text-boundary.js";

/** Private primitive. A caller still has to prove format-level ownership/edit semantics. */
export interface XmlTextEdit {
  readonly elementIndex: number;
  readonly namespaceUri: string;
  readonly local: string;
  readonly expectedText: string;
  readonly text: string;
}
export interface XmlTextSourceElement {
  readonly elementIndex: number;
  /** Raw-tree parent identity; absent only for the document element. */
  readonly parentElementIndex?: number;
  /** Half-open UTF-16 offsets of the complete element in decoded XML, excluding the BOM. */
  readonly outerStart: number;
  readonly outerEnd: number;
  readonly name: string;
  readonly namespaceUri: string;
  readonly local: string;
  /** Absent when children include elements, comments or processing instructions. */
  readonly text?: string;
}
export interface XmlTextEditLimits {
  maxEdits: number;
  maxReplacementCharacters: number;
  maxOutputBytes: number;
  maxWorkUnits: number;
}
export interface XmlTextEditOptions {
  readonly limits?: Partial<XmlTextEditLimits>;
  readonly signal?: AbortSignal;
}
export interface XmlTextChange {
  readonly elementIndex: number;
  /** Half-open UTF-16 offsets into the decoded source, excluding its BOM. */
  readonly sourceStart: number;
  readonly sourceEnd: number;
}
export interface XmlTextSource {
  readonly inspection: Readonly<XmlPartInspection>;
  readonly elements: readonly XmlTextSourceElement[];
  /** Read one captured raw element without granting permission to modify it. */
  readElement(request: { readonly elementIndex: number; readonly namespaceUri: string; readonly local: string }, options?: { readonly maxCharacters?: number; readonly signal?: AbortSignal }): string;
  /** Inspect a direct paragraph/run/text split boundary; it does not prepare a write. */
  preflightRunSplit(request: { readonly elementIndex: number; readonly namespaceUri: string; readonly expectedText: string; readonly offset: number }, options?: { readonly maxCharacters?: number; readonly maxGraphemes?: number; readonly signal?: AbortSignal }): Promise<Readonly<{ left: string; right: string; runElementIndex: number; paragraphElementIndex: number }>>;
  /** Experimental lexical primitive for a simple text run; no package-level edit authority. */
  splitSimpleRun(request: { readonly elementIndex: number; readonly namespaceUri: string; readonly expectedText: string; readonly offset: number }, options?: { readonly maxCharacters?: number; readonly maxGraphemes?: number; readonly maxOutputBytes?: number; readonly maxWorkUnits?: number; readonly signal?: AbortSignal }): Promise<{ bytes: Uint8Array }>;
  replaceText(edits: readonly XmlTextEdit[], options?: XmlTextEditOptions): Promise<{ bytes: Uint8Array; changes: readonly XmlTextChange[] }>;
  dispose(): void;
}
export class XmlSourceEditError extends Error {
  constructor(readonly code: "xml-edit-conflict" | "xml-edit-unsupported" | "xml-edit-limit-exceeded" | "xml-source-disposed", message: string) {
    super(message); this.name = "XmlSourceEditError";
  }
}
export const DEFAULT_XML_TEXT_EDIT_LIMITS: Readonly<XmlTextEditLimits> = Object.freeze({
  maxEdits: 10000, maxReplacementCharacters: 2 * 1024 * 1024,
  maxOutputBytes: 8 * 1024 * 1024, maxWorkUnits: 32 * 1024 * 1024,
});

interface Span { start: number; end: number; outerStart: number; outerEnd: number; selfClosing: boolean; parentElementIndex?: number }
interface Patch extends Span { elementIndex: number; text: string }
function conflict(message: string): never { throw new XmlSourceEditError("xml-edit-conflict", message); }
function unsupported(message: string): never { throw new XmlSourceEditError("xml-edit-unsupported", message); }
function limit(message: string): never { throw new XmlSourceEditError("xml-edit-limit-exceeded", message); }
const textEscapes: Readonly<Record<string, string>> = Object.freeze({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\r": "&#xD;" });
const escapeText = (text: string): string => text.replace(/[&<>\r]/g, (value) => textEscapes[value]!);
const ws = (c: string | undefined): boolean => c === " " || c === "\t" || c === "\r" || c === "\n";

function orderedElements(root: XmlElement): XmlElement[] {
  const result: XmlElement[] = [], stack = [root];
  while (stack.length) {
    const element = stack.pop()!; result.push(element);
    for (let i = element.children.length - 1; i >= 0; i--) {
      const child = element.children[i]!;
      if (child.kind === "element") stack.push(child);
    }
  }
  return result;
}
function textOf(element: XmlElement): string | undefined {
  if (element.children.some((child) => child.kind !== "text" && child.kind !== "cdata")) return undefined;
  return element.children.map((child) => child.kind === "text" || child.kind === "cdata" ? child.value : "").join("");
}

/**
 * Bind lexical spans to the already validated parser tree, never to a text search.
 * This scanner does not accept XML independently; declarations/entities/encoding
 * and namespace correctness have already passed the bounded XML reader.
 */
async function sourceSpans(text: string, elements: readonly XmlElement[], check: () => void): Promise<Span[]> {
  const spans: Span[] = [], stack: { name: string; index: number }[] = [];
  let position = 0, index = 0, tokens = 0;
  while (position < text.length) {
    const start = text.indexOf("<", position);
    if (start < 0) break;
    if (++tokens % 256 === 0) { await new Promise<void>((resolve) => setTimeout(resolve, 0)); check(); }
    if (text.startsWith("<!--", start)) { position = text.indexOf("-->", start + 4) + 3; continue; }
    if (text.startsWith("<![CDATA[", start)) { position = text.indexOf("]]>", start + 9) + 3; continue; }
    if (text.startsWith("<?", start)) { position = text.indexOf("?>", start + 2) + 2; continue; }
    const closing = text[start + 1] === "/", nameStart = start + (closing ? 2 : 1);
    let nameEnd = nameStart;
    while (nameEnd < text.length && !ws(text[nameEnd]) && text[nameEnd] !== "/" && text[nameEnd] !== ">") nameEnd++;
    const name = text.slice(nameStart, nameEnd);
    let end = nameEnd, quote = "";
    for (; end < text.length; end++) {
      const c = text[end]!;
      if (quote) { if (c === quote) quote = ""; }
      else if (c === "'" || c === '"') quote = c;
      else if (c === ">") break;
    }
    if (end >= text.length) unsupported("Validated XML source could not be bound to lexical spans.");
    position = end + 1;
    if (closing) {
      const open = stack.pop();
      if (!open || open.name !== name) unsupported("Validated XML closing identity disagrees with source spans.");
      spans[open.index]!.end = start;
      spans[open.index]!.outerEnd = end + 1;
    } else {
      if (elements[index]?.name !== name) unsupported("Validated XML preorder disagrees with source spans.");
      const selfClosing = text[end - 1] === "/";
      const parent = stack.at(-1);
      spans.push({ start: selfClosing ? end - 1 : end + 1, end: selfClosing ? end + 1 : -1, outerStart: start, outerEnd: selfClosing ? end + 1 : -1, selfClosing, ...(parent ? { parentElementIndex: parent.index + 1 } : {}) });
      if (!selfClosing) stack.push({ name, index });
      index++;
    }
  }
  if (stack.length || index !== elements.length || spans.some((span) => span.end < span.start || span.outerEnd <= span.outerStart)) unsupported("Validated XML source spans are incomplete.");
  check();
  return spans;
}

/** UTF-8/UTF-16 encoding without newline, namespace, prefix or lexical rewriting. */
function encode(text: string, encoding: XmlPartInspection["encoding"], bom: Uint8Array, size: number): Uint8Array {
  if (encoding === "utf-8") {
    const output = new Uint8Array(size); output.set(bom); output.set(new TextEncoder().encode(text), bom.length); return output;
  }
  const output = new Uint8Array(size); output.set(bom);
  for (let i = 0; i < text.length; i++) {
    const value = text.charCodeAt(i), offset = bom.length + i * 2;
    output[offset + (encoding === "utf-16le" ? 0 : 1)] = value & 255;
    output[offset + (encoding === "utf-16le" ? 1 : 0)] = value >>> 8;
  }
  return output;
}

/** Captures one XML part only. No OPC writer, package permission or public API. */
export async function captureXmlTextSource(input: Uint8Array, options: XmlInspectionOptions = {}): Promise<XmlTextSource> {
  if (!options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("Invalid XML source options.");
  if (options.limits !== undefined && (!options.limits || typeof options.limits !== "object" || Array.isArray(options.limits))) throw new TypeError("Invalid XML source limits.");
  if (!(input instanceof Uint8Array)) throw new TypeError("XML source input must be Uint8Array.");
  const view = byteViewInfo(input);
  if (view.shared) unsupported("XML source capture requires an ordinary byte buffer.");
  // The reader validates these captured limits and the private byte snapshot.
  const captured = { ...options, limits: { ...options.limits } };
  const inputLimit = captured.limits.maxInputBytes ?? DEFAULT_XML_LIMITS.maxInputBytes;
  if (!Number.isSafeInteger(inputLimit) || inputLimit <= 0) throw new RangeError("Invalid XML limit maxInputBytes.");
  if (view.length > inputLimit) throw new PackageInspectionError("xml-limit-exceeded", "XML source exceeds maxInputBytes.", captured.partName);
  const bytes = copyByteView(input);
  const parsed = await readXml(bytes, captured, true), encoding = parsed.inspection.encoding;
  const text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
  const bomLength = encoding !== "utf-8" ? 2 : bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  const bom = bytes.slice(0, bomLength), nodes = orderedElements(parsed.root!);
  const checkCapture = () => { if (captured.signal?.aborted) throw new PackageInspectionError("operation-aborted", "XML source capture was aborted.", captured.partName); };
  const spans = await sourceSpans(text, nodes, checkCapture);
  const elements = Object.freeze(nodes.map((node, index): XmlTextSourceElement => {
    const value = textOf(node);
    const parentElementIndex = spans[index]!.parentElementIndex;
    return Object.freeze({ elementIndex: index + 1, ...(parentElementIndex !== undefined ? { parentElementIndex } : {}), outerStart: spans[index]!.outerStart, outerEnd: spans[index]!.outerEnd, name: node.name, namespaceUri: node.namespaceUri, local: node.local, ...(value !== undefined ? { text: value } : {}) });
  }));
  const inspection = Object.freeze({ ...parsed.inspection, root: Object.freeze({ ...parsed.inspection.root }), namespaceUris: Object.freeze([...parsed.inspection.namespaceUris]) });
  // All mutable source and tree records remain private; disposal releases source text.
  let sourceText: string | undefined = text, sourceNodes: XmlElement[] | undefined = nodes;
  const live = () => { if (sourceText === undefined) throw new XmlSourceEditError("xml-source-disposed", "The XML source has been disposed."); };
  const api: XmlTextSource = Object.freeze({
    inspection, elements,
    readElement(request: { readonly elementIndex: number; readonly namespaceUri: string; readonly local: string }, options: { readonly maxCharacters?: number; readonly signal?: AbortSignal } = {}) {
      live();
      if (!request || typeof request !== "object" || Array.isArray(request) || !options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("Invalid XML element read request.");
      const { elementIndex, namespaceUri, local } = request;
      const { maxCharacters = 2 * 1024 * 1024, signal } = options;
      live();
      if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "XML element reading was aborted.", captured.partName);
      live();
      if (!Number.isSafeInteger(maxCharacters) || maxCharacters <= 0) throw new RangeError("Invalid XML element read limit.");
      if (!Number.isSafeInteger(elementIndex) || elementIndex <= 0 || typeof namespaceUri !== "string" || typeof local !== "string") throw new TypeError("Invalid XML element identity.");
      const element = elements[elementIndex - 1];
      if (!element || element.namespaceUri !== namespaceUri || element.local !== local) conflict("XML source element identity does not match the requested expanded name.");
      if (element.outerEnd - element.outerStart > maxCharacters) limit("XML element source exceeds maxCharacters.");
      return sourceText!.slice(element.outerStart, element.outerEnd);
    },
    async preflightRunSplit(request: { readonly elementIndex: number; readonly namespaceUri: string; readonly expectedText: string; readonly offset: number }, options: { readonly maxCharacters?: number; readonly maxGraphemes?: number; readonly signal?: AbortSignal } = {}) {
      live();
      if (!request || typeof request !== "object" || Array.isArray(request) || !options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("Invalid run split preflight.");
      const { elementIndex, namespaceUri, expectedText, offset } = request;
      live();
      if (!Number.isSafeInteger(elementIndex) || elementIndex <= 0 || typeof namespaceUri !== "string" || typeof expectedText !== "string" || !Number.isSafeInteger(offset)) throw new TypeError("Invalid run split identity or offset.");
      const atom = elements[elementIndex - 1];
      if (!atom || atom.local !== "t" || atom.namespaceUri !== namespaceUri) conflict("Run split text identity does not match the source.");
      if (atom.text === undefined) unsupported("Run split preflight requires character-only text.");
      if (atom.text !== expectedText) conflict("Run split text does not match the expected baseline.");
      const run = atom.parentElementIndex === undefined ? undefined : elements[atom.parentElementIndex - 1];
      const paragraph = run?.parentElementIndex === undefined ? undefined : elements[run.parentElementIndex - 1];
      if (!run || run.local !== "r" || run.namespaceUri !== namespaceUri || !paragraph || paragraph.local !== "p" || paragraph.namespaceUri !== namespaceUri) unsupported("Run split preflight requires direct raw run and paragraph ancestry.");
      const { maxCharacters, maxGraphemes, signal } = options;
      live();
      const split = await splitAtGraphemeBoundary(expectedText, offset, { ...(maxCharacters !== undefined ? { maxCharacters } : {}), ...(maxGraphemes !== undefined ? { maxGraphemes } : {}), ...(signal !== undefined ? { signal } : {}), checkpoint: live });
      live();
      return Object.freeze({ ...split, runElementIndex: run.elementIndex, paragraphElementIndex: paragraph.elementIndex });
    },
    async splitSimpleRun(request: { readonly elementIndex: number; readonly namespaceUri: string; readonly expectedText: string; readonly offset: number }, options: { readonly maxCharacters?: number; readonly maxGraphemes?: number; readonly maxOutputBytes?: number; readonly maxWorkUnits?: number; readonly signal?: AbortSignal } = {}) {
      live();
      if (!request || typeof request !== "object" || Array.isArray(request) || !options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("Invalid plain run split request.");
      const { elementIndex, namespaceUri, expectedText, offset } = request;
      const { maxCharacters, maxGraphemes, maxOutputBytes = DEFAULT_XML_TEXT_EDIT_LIMITS.maxOutputBytes, maxWorkUnits = DEFAULT_XML_TEXT_EDIT_LIMITS.maxWorkUnits, signal } = options;
      live();
      if (!Number.isSafeInteger(maxOutputBytes) || maxOutputBytes <= 0 || !Number.isSafeInteger(maxWorkUnits) || maxWorkUnits <= 0) throw new RangeError("Invalid plain run split limits.");
      const check = () => { live(); if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "Plain run split was aborted.", captured.partName); live(); };
      check();
      const plan = await api.preflightRunSplit({ elementIndex, namespaceUri, expectedText, offset }, { ...(maxCharacters !== undefined ? { maxCharacters } : {}), ...(maxGraphemes !== undefined ? { maxGraphemes } : {}), ...(signal !== undefined ? { signal } : {}) });
      check();
      const runIndex = plan.runElementIndex - 1, textIndex = elementIndex - 1, run = sourceNodes![runIndex], atom = sourceNodes![textIndex];
      if (!run || !atom || run.attributes.length !== 0 || run.children.some(child => child.kind !== "element" && (child.kind !== "text" || !/^[\t\r\n ]*$/.test(child.value)))) unsupported("Simple run split rejects run attributes and opaque children.");
      const children = run.children.filter((child): child is XmlElement => child.kind === "element");
      const properties = children.length === 2 ? children[0] : undefined;
      if ((children.length !== 1 && children.length !== 2) || children.at(-1) !== atom || (properties && (properties.local !== "rPr" || properties.namespaceUri !== namespaceUri))) unsupported("Simple run split requires one text child and optional direct run properties.");
      if (properties && (properties.attributes.some(attribute => attribute.namespaceUri !== "" || !["b", "i", "sz", "lang"].includes(attribute.local)) || properties.children.some(child => child.kind !== "text" || !/^[\t\r\n ]*$/.test(child.value)))) unsupported("Simple run split rejects unknown run properties, hyperlinks and complex children.");
      if (atom.attributes.some(attribute => attribute.namespaceUri !== "http://www.w3.org/XML/1998/namespace" || attribute.local !== "space" || !["default", "preserve"].includes(attribute.value))) unsupported("Plain run split does not duplicate unsupported text attributes.");
      const runSpan = spans[runIndex]!, textSpan = spans[textIndex]!;
      if (textSpan.selfClosing || textSpan.start < runSpan.outerStart || textSpan.end > runSpan.outerEnd) unsupported("Plain run lexical spans are incompatible.");
      const snapshot = sourceText!;
      const before = snapshot.slice(runSpan.outerStart, textSpan.start), after = snapshot.slice(textSpan.end, runSpan.outerEnd);
      const segments = [
        { text: snapshot.slice(0, runSpan.outerStart), escape: false }, { text: before, escape: false }, { text: plan.left, escape: true }, { text: after, escape: false },
        { text: before, escape: false }, { text: plan.right, escape: true }, { text: after, escape: false }, { text: snapshot.slice(runSpan.outerEnd), escape: false },
      ];
      let outputBytes = bom.length, work = 0, pending = 0;
      for (const segment of segments) {
        for (let i = 0; i < segment.text.length; i++) {
          const value = segment.text.codePointAt(i)!; let units = 1;
          if (value > 0xffff) { i++; units++; }
          const escapedLength = segment.escape ? value === 38 ? 5 : value === 60 || value === 62 ? 4 : value === 13 ? 5 : 0 : 0;
          outputBytes += escapedLength ? escapedLength * (encoding === "utf-8" ? 1 : 2) : encoding !== "utf-8" ? units * 2 : value < 0x80 ? 1 : value < 0x800 ? 2 : value < 0x10000 ? 3 : 4;
          work += units; pending += units;
          if (outputBytes > maxOutputBytes || work > maxWorkUnits) limit("Plain run split exceeds its output or work budget.");
          if (pending >= 32768) { pending = 0; await new Promise<void>(resolve => setTimeout(resolve, 0)); check(); }
        }
      }
      check();
      const output = encode(segments.map(segment => segment.escape ? escapeText(segment.text) : segment.text).join(""), encoding, bom, outputBytes);
      const verified = await readXml(output, { ...(captured.partName !== undefined ? { partName: captured.partName } : {}), ...(signal ? { signal } : {}), limits: { ...DEFAULT_XML_LIMITS, ...captured.limits } }, true);
      check();
      const outputNodes = orderedElements(verified.root!);
      const subtreeSize = properties ? 3 : 2, atomOffset = textIndex - runIndex;
      if (atomOffset !== subtreeSize - 1 || outputNodes.length !== elements.length + subtreeSize) unsupported("Simple run split changed unexpected source elements.");
      for (let i = 0; i < outputNodes.length; i++) {
        const original = i < runIndex ? i : i < runIndex + 2 * subtreeSize ? runIndex + ((i - runIndex) % subtreeSize) : i - subtreeSize;
        if (outputNodes[i]!.name !== elements[original]!.name || outputNodes[i]!.namespaceUri !== elements[original]!.namespaceUri) unsupported("Simple run split changed an unexpected element identity.");
      }
      if (textOf(outputNodes[runIndex + atomOffset]!) !== plan.left || textOf(outputNodes[runIndex + subtreeSize + atomOffset]!) !== plan.right) unsupported("Simple run split text differs from the requested boundary.");
      check();
      return { bytes: output };
    },
    async replaceText(requests: readonly XmlTextEdit[], editOptions: XmlTextEditOptions = {}) {
      live();
      if (!editOptions || typeof editOptions !== "object" || Array.isArray(editOptions) || (editOptions.limits !== undefined && (!editOptions.limits || typeof editOptions.limits !== "object" || Array.isArray(editOptions.limits)))) throw new TypeError("Invalid XML edit options.");
      const limits = { ...DEFAULT_XML_TEXT_EDIT_LIMITS, ...editOptions.limits }, signal = editOptions.signal;
      for (const [name, value] of Object.entries(limits)) if (!Number.isSafeInteger(value) || value <= 0) throw new RangeError(`Invalid XML edit limit ${name}.`);
      const check = () => {
        live();
        if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "XML text editing was aborted.", captured.partName);
      };
      check();
      if (!Array.isArray(requests)) throw new TypeError("XML text edits must be an array.");
      if (requests.length > limits.maxEdits) limit("XML text edit count exceeds maxEdits.");
      // Copy scalars before yielding; caller mutations never affect the candidate.
      const edits = requests.map((edit) => ({ ...edit })), selected = new Set<number>(), patches: Patch[] = [];
      check();
      let replacementCharacters = 0, work = 0, pending = 0;
      const charge = async (units: number) => {
        work += units; pending += units;
        if (work > limits.maxWorkUnits) limit("XML text editing exceeds maxWorkUnits.");
        if (pending >= 32768) { pending = 0; await new Promise<void>((resolve) => setTimeout(resolve, 0)); check(); }
      };
      for (const edit of edits) {
        await charge(1); check();
        if (!Number.isSafeInteger(edit.elementIndex) || edit.elementIndex <= 0 || typeof edit.namespaceUri !== "string" || typeof edit.local !== "string" || typeof edit.expectedText !== "string" || typeof edit.text !== "string") throw new TypeError("Invalid XML text edit.");
        if (selected.has(edit.elementIndex)) conflict("An XML source element cannot be edited twice in one candidate.");
        selected.add(edit.elementIndex);
        const element = elements[edit.elementIndex - 1], span = spans[edit.elementIndex - 1];
        if (!element || element.namespaceUri !== edit.namespaceUri || element.local !== edit.local) conflict("XML source element identity does not match the requested expanded name.");
        if (element.text === undefined) unsupported("Text editing requires character-only content; child elements, comments and processing instructions must remain untouched.");
        if (element.text !== edit.expectedText) conflict("XML source text does not match the expected baseline.");
        replacementCharacters += edit.text.length;
        if (replacementCharacters > limits.maxReplacementCharacters) limit("XML replacement text exceeds maxReplacementCharacters.");
        if (edit.text !== element.text) patches.push({ ...span!, elementIndex: edit.elementIndex, text: edit.text });
      }
      // Canonical source order; caller ordering cannot alter the result.
      check();
      patches.sort((a, b) => a.start - b.start);
      let position = 0, outputBytes = bom.length;
      const segments: { text: string; escape: boolean }[] = [];
      for (const patch of patches) {
        if (patch.start < position) conflict("XML text edit regions overlap.");
        segments.push({ text: sourceText!.slice(position, patch.start), escape: false });
        if (patch.selfClosing) segments.push({ text: ">", escape: false });
        segments.push({ text: patch.text, escape: true });
        if (patch.selfClosing) segments.push({ text: `</${elements[patch.elementIndex - 1]!.name}>`, escape: false });
        position = patch.end;
      }
      segments.push({ text: sourceText!.slice(position), escape: false });
      await Promise.resolve(); check();
      // Count the actual encoded candidate before allocating escaped/output strings.
      for (const segment of segments) {
        let block = 0;
        for (let i = 0; i < segment.text.length; i++) {
          const value = segment.text.codePointAt(i)!;
          let units = 1;
          if (value > 0xffff) { i++; units++; }
          if (segment.escape && !((value >= 0x20 && value <= 0xd7ff) || (value >= 0xe000 && value <= 0xfffd) || (value >= 0x10000 && value <= 0x10ffff) || value === 9 || value === 10 || value === 13)) unsupported("Replacement text contains an invalid XML 1.0 character.");
          const escapedLength = segment.escape ? value === 38 ? 5 : value === 60 || value === 62 ? 4 : value === 13 ? 5 : 0 : 0;
          outputBytes += escapedLength ? escapedLength * (encoding === "utf-8" ? 1 : 2) : encoding !== "utf-8" ? units * 2 : value < 0x80 ? 1 : value < 0x800 ? 2 : value < 0x10000 ? 3 : 4;
          if (outputBytes > limits.maxOutputBytes) limit("XML candidate exceeds maxOutputBytes.");
          block += units;
          if (block >= 8192) { await charge(block); block = 0; }
        }
        await charge(block);
      }
      if (outputBytes > limits.maxOutputBytes) limit("XML candidate exceeds maxOutputBytes.");
      check();
      const candidate = segments.map((segment) => segment.escape ? escapeText(segment.text) : segment.text).join("");
      const output = encode(candidate, encoding, bom, outputBytes);
      check();
      // Revalidate under the original XML budgets as well as the edit budgets.
      const verified = await readXml(output, { ...(captured.partName !== undefined ? { partName: captured.partName } : {}), ...(signal ? { signal } : {}), limits: { ...DEFAULT_XML_LIMITS, ...captured.limits } }, true);
      check();
      const outputElements = orderedElements(verified.root!);
      if (outputElements.length !== elements.length) unsupported("XML candidate changed the source element structure.");
      for (let i = 0; i < outputElements.length; i++) {
        const before = elements[i]!, after = outputElements[i]!;
        if (before.name !== after.name || before.namespaceUri !== after.namespaceUri) unsupported("XML candidate changed a source element identity.");
      }
      for (const edit of edits) if (textOf(outputElements[edit.elementIndex - 1]!) !== edit.text) unsupported("XML candidate text disagrees with the requested replacement.");
      check();
      return { bytes: output, changes: Object.freeze(patches.map((patch) => Object.freeze({ elementIndex: patch.elementIndex, sourceStart: patch.start, sourceEnd: patch.end }))) };
    },
    dispose() { sourceText = undefined; sourceNodes = undefined; },
  });
  return api;
}
