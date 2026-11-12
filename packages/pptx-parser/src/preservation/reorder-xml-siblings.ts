import { PackageInspectionError } from "../types/archive.js";
import type { XmlInspectionOptions } from "../types/xml.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { captureXmlTextSource, XmlSourceEditError } from "./xml-text-source.js";

/** A private lexical primitive, not authority to reorder OOXML objects. */
export async function reorderXmlSiblings(input: Uint8Array, request: {
  readonly parentElementIndex: number;
  readonly namespaceUri: string;
  readonly local: string;
  readonly elementIndexes: readonly number[];
  readonly requestedIndexes: readonly number[];
}, options: XmlInspectionOptions & { readonly maxOutputBytes: number; readonly maxWorkUnits: number; readonly maxElements: number }): Promise<{ bytes: Uint8Array; workUnits: number; movedElementIndexes: ReadonlyMap<number, number> }> {
  if (!(input instanceof Uint8Array) || !request || typeof request !== "object" || !options || typeof options !== "object") throw new TypeError("Invalid XML sibling reorder input.");
  const { parentElementIndex, namespaceUri, local } = request, { maxOutputBytes, maxWorkUnits, maxElements, signal, partName } = options;
  const inspectionOptions: XmlInspectionOptions = { ...options, ...(options.limits ? { limits: { ...options.limits } } : {}) };
  if (![maxOutputBytes, maxWorkUnits, maxElements].every(value => Number.isSafeInteger(value) && value > 0)) throw new RangeError("Invalid XML sibling reorder limits.");
  if (!Number.isSafeInteger(parentElementIndex) || parentElementIndex <= 0 || typeof namespaceUri !== "string" || typeof local !== "string" || !local || !Array.isArray(request.elementIndexes) || !Array.isArray(request.requestedIndexes)) throw new TypeError("Invalid XML sibling reorder identity.");
  if (!request.elementIndexes.length || request.elementIndexes.length > maxElements || request.requestedIndexes.length !== request.elementIndexes.length) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML sibling reorder exceeds maxElements or has no elements.");
  if (request.elementIndexes.length > maxWorkUnits) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML sibling identities exceed maxWorkUnits.");
  const original = [...request.elementIndexes], requested = [...request.requestedIndexes];
  const identities = new Set(original);
  if (original.some(value => !Number.isSafeInteger(value) || value <= 0) || identities.size !== original.length || new Set(requested).size !== original.length || requested.some(value => !identities.has(value))) throw new TypeError("XML sibling reorder requires an exact unique identity permutation.");
  const check = () => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "XML sibling reorder was aborted.", partName); };
  check(); const view = byteViewInfo(input);
  if (view.shared) throw new TypeError("XML sibling reorder requires an ordinary byte buffer.");
  if (view.length > maxOutputBytes) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML sibling reorder exceeds maxOutputBytes.");
  const workUnits = view.length * 4 + original.length * 4;
  if (!Number.isSafeInteger(workUnits) || workUnits > maxWorkUnits) throw new XmlSourceEditError("xml-edit-limit-exceeded", "XML sibling reorder exceeds maxWorkUnits.");
  const bytes = copyByteView(input), source = await captureXmlTextSource(bytes, inspectionOptions);
  try {
    check();
    const selected = original.map(index => source.elements[index - 1]);
    if (selected.some((element, index) => !element || element.parentElementIndex !== parentElementIndex || element.namespaceUri !== namespaceUri || element.local !== local || index > 0 && selected[index - 1]!.outerStart >= element.outerStart)) throw new XmlSourceEditError("xml-edit-conflict", "XML sibling source identities or physical order differ from the request.");
    const encoding = source.inspection.encoding, text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
    const slices = selected.map(element => text.slice(element!.outerStart, element!.outerEnd)), positions = new Map(original.map((index, position) => [index, position]));
    for (let i = 1; i < selected.length; i++) if (!/^[\t\r\n ]*$/u.test(text.slice(selected[i - 1]!.outerEnd, selected[i]!.outerStart))) throw new XmlSourceEditError("xml-edit-unsupported", "XML sibling elements must be separated only by whitespace.");
    const pieces = [text.slice(0, selected[0]!.outerStart)];
    for (let i = 0; i < selected.length; i++) {
      pieces.push(slices[positions.get(requested[i]!)!]!, text.slice(selected[i]!.outerEnd, i + 1 < selected.length ? selected[i + 1]!.outerStart : text.length));
    }
    const candidate = pieces.join("");
    const bomLength = encoding === "utf-8" ? bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0 : 2;
    let output: Uint8Array;
    if (encoding === "utf-8") { const encoded = new TextEncoder().encode(candidate); output = new Uint8Array(bomLength + encoded.length); output.set(bytes.subarray(0, bomLength)); output.set(encoded, bomLength); }
    else { output = new Uint8Array(bomLength + candidate.length * 2); output.set(bytes.subarray(0, bomLength)); for (let i = 0; i < candidate.length; i++) { const unit = candidate.charCodeAt(i), offset = bomLength + i * 2; output[offset + (encoding === "utf-16le" ? 0 : 1)] = unit & 255; output[offset + (encoding === "utf-16le" ? 1 : 0)] = unit >>> 8; } }
    if (output.length !== bytes.length || output.length > maxOutputBytes) throw new XmlSourceEditError("xml-edit-conflict", "XML sibling reorder changed the encoded size.");
    check(); const rebound = await captureXmlTextSource(output, inspectionOptions), movedElementIndexes = new Map<number, number>();
    try {
      const children = rebound.elements.filter(element => element.parentElementIndex === parentElementIndex && element.namespaceUri === namespaceUri && element.local === local && element.outerStart >= selected[0]!.outerStart && element.outerStart < selected[selected.length - 1]!.outerEnd);
      if (children.length !== requested.length || children.some((element, i) => rebound.readElement({ elementIndex: element.elementIndex, namespaceUri, local }, { maxCharacters: maxOutputBytes }) !== slices[positions.get(requested[i]!)!])) throw new XmlSourceEditError("xml-edit-conflict", "Reordered XML siblings failed source verification.");
      for (const [index, child] of children.entries()) {
        const oldRoot = selected[positions.get(requested[index]!)!]!;
        let oldIndex = oldRoot.elementIndex, newIndex = child.elementIndex;
        while (oldIndex <= source.elements.length && source.elements[oldIndex - 1]!.outerStart < oldRoot.outerEnd) {
          const before = source.elements[oldIndex - 1]!, after = rebound.elements[newIndex - 1];
          if (!after || after.outerStart >= child.outerEnd || before.namespaceUri !== after.namespaceUri || before.local !== after.local) throw new XmlSourceEditError("xml-edit-conflict", "Reordered XML descendant identities differ from their source.");
          movedElementIndexes.set(oldIndex++, newIndex++);
        }
        if (rebound.elements[newIndex - 1] && rebound.elements[newIndex - 1]!.outerStart < child.outerEnd) throw new XmlSourceEditError("xml-edit-conflict", "Reordered XML descendant count differs from its source.");
      }
    } finally { rebound.dispose(); }
    check(); return { bytes: output, workUnits, movedElementIndexes };
  } finally { source.dispose(); }
}
