import { readOpcSource, snapshotOptions } from "../opc/read-package.js";
import { asciiKey } from "../opc/names.js";
import { inspectClosure } from "../preservation/inspect-closure.js";
import { appendImageContentType } from "../preservation/append-image-content-type.js";
import { appendImageRelationship } from "../preservation/append-image-relationship.js";
import { replaceXmlRelationshipId } from "../preservation/replace-xml-relationship-id.js";
import { replaceArchiveEntries } from "../preservation/replace-archive-entries.js";
import { structureContext } from "../presentation/structure-context.js";
import { readStructureFromContext } from "../presentation/read-structure.js";
import { attribute } from "../presentation/structure-xml.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { rasterMetadata } from "./image-metadata.js";
import { validatePngReplacement } from "./validate-png-replacement.js";
import { validateGifReplacement } from "./validate-gif-replacement.js";
import { validateJpegReplacement } from "./validate-jpeg-replacement.js";
import { projectStructure } from "./project-structure.js";
import { projectContent } from "./project-content.js";
import { snapshot } from "./state.js";
import { verifyPictureProjection } from "./verify-picture-projection.js";
import { PptxImportError, type ImportedPictureSource, type PptxSessionLimits } from "../types/import.js";
import type { PptxInspectionOptions } from "../types/structure.js";

export class PictureSourceUnsupportedError extends PptxImportError {
  constructor(message: string, source?: { part: string; elementIndex?: number }) { super("save-candidate-invalid", message, source); }
}

/** Private candidate builder. The session must still validate the complete mapped Core state. */
export async function preparePictureCandidate(source: Uint8Array, binding: ImportedPictureSource, image: Uint8Array, mimeType: "image/png" | "image/jpeg" | "image/gif", options: PptxInspectionOptions, limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>): Promise<{ bytes: Uint8Array; part: string; relationshipId: string }> {
  if (!(source instanceof Uint8Array) || !(image instanceof Uint8Array) || !binding || !["image/png", "image/jpeg", "image/gif"].includes(mimeType)) throw new TypeError("Invalid picture candidate input.");
  const sourceView = byteViewInfo(source), imageView = byteViewInfo(image);
  if (sourceView.shared || imageView.shared) throw new TypeError("Picture candidate requires ordinary byte buffers.");
  const captured = snapshotOptions(options), savedLimits = { ...limits }, targetBinding = structuredClone(binding), identity = { part: targetBinding.source.part, index: targetBinding.sourceElementIndex, id: targetBinding.relationshipId, media: targetBinding.mediaPart };
  if (captured.signal?.aborted) throw new PptxImportError("operation-aborted", "Picture candidate preparation was aborted.");
  if (sourceView.length > captured.zipLimits.maxInputBytes || imageView.length > savedLimits.maxImportedAssetBytes || imageView.length > savedLimits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Picture source or replacement exceeds its byte budget.");
  const originalBytes = copyByteView(source), replacement = copyByteView(image), signal = captured.signal;
  let analysisWork = 0;
  const check = async () => { if (signal?.aborted) throw new PptxImportError("operation-aborted", "Picture candidate preparation was aborted."); if (++analysisWork > savedLimits.maxSaveAnalysisWork) throw new PptxImportError("save-limit-exceeded", "Picture candidate preparation exceeds maxSaveAnalysisWork."); await checkpoint(); if (signal?.aborted) throw new PptxImportError("operation-aborted", "Picture candidate preparation was aborted."); };
  const optionsCopy = { limits: captured.limits, xmlLimits: captured.xmlLimits, zipLimits: captured.zipLimits, ...(signal ? { signal } : {}) };
  const validationUsage: { work: number; decodedBytes?: number } = { work: 0 };
  await check();
  const metadata = await rasterMetadata(replacement, check);
  if (!metadata || metadata.mimeType !== mimeType) throw new PictureSourceUnsupportedError("Picture replacement header and declared MIME type must agree.");
  if (mimeType === "image/png" && !await validatePngReplacement(replacement, savedLimits.maxSaveValidationWork, check, validationUsage, savedLimits.maxSaveValidationBytes, signal)) throw new PictureSourceUnsupportedError("PNG replacement has an invalid or unsupported container or compressed scanline stream.");
  if (mimeType === "image/gif" && !await validateGifReplacement(replacement, savedLimits.maxSaveValidationWork, savedLimits.maxSaveValidationBytes, validationUsage, check)) throw new PictureSourceUnsupportedError("GIF replacement has an invalid or unsupported container or LZW stream.");
  if (mimeType === "image/jpeg" && !await validateJpegReplacement(replacement, savedLimits.maxSaveValidationWork, validationUsage, check)) throw new PictureSourceUnsupportedError("JPEG replacement has invalid or unsupported markers, tables or scan framing.");
  const closure = await inspectClosure(originalBytes, { ...optionsCopy, roots: [identity.part], closureLimits: { maxFindings: savedLimits.maxSaveChanges } });
  if (closure.findings.length) throw new PictureSourceUnsupportedError("Picture source relationship closure is incomplete.", { part: identity.part });
  const original = await readOpcSource(originalBytes, optionsCopy, false, true), context = await structureContext(originalBytes, optionsCopy, true);
  const relation = original.inspection.relationships.find(item => item.source === identity.part && item.id === identity.id);
  if (!relation || relation.targetMode !== "Internal" || relation.type !== `${context.r}/image` || relation.targetPart !== identity.media) throw new PictureSourceUnsupportedError("Picture source relationship is stale or mismatched.", { part: identity.part });
  const root = context.rawXml.get(identity.part), nodes = root ? [root] : []; let index = 0, selected = undefined as typeof root;
  while (nodes.length) {
    const node = nodes.pop()!; if (++index === identity.index) selected = node;
    for (let i = node.children.length - 1; i >= 0; i--) { const child = node.children[i]!; if (child.kind === "element") nodes.push(child); }
    await check();
  }
  if (!selected || selected.namespaceUri !== context.a || selected.local !== "blip" || attribute(selected, "embed", context.r) !== identity.id || attribute(selected, "link", context.r) !== undefined) throw new PictureSourceUnsupportedError("Picture source blip is stale or has an external link.", { part: identity.part, elementIndex: identity.index });
  // Conservatively reject all opaque blip effects and attributes. An alphaModFix
  // with its sole direct scalar is retained but no other effect is interpreted.
  if (selected.attributes.some(item => item.namespaceUri !== "http://www.w3.org/2000/xmlns/" && !(item.namespaceUri === context.r && item.local === "embed")) || selected.children.filter(child => child.kind === "element").length > 1 || selected.children.some(child => child.kind === "element" ? child.namespaceUri !== context.a || child.local !== "alphaModFix" || child.children.length !== 0 || child.attributes.some(item => item.namespaceUri !== "http://www.w3.org/2000/xmlns/" && (item.namespaceUri !== "" || item.local !== "amt")) : (child.kind === "text" || child.kind === "cdata") && !/^[\t\r\n ]*$/.test(child.value))) throw new PictureSourceUnsupportedError("Picture source blip contains unverified attributes or effects.", { part: identity.part, elementIndex: identity.index });
  const partNames = new Set(original.inspection.parts.map(item => asciiKey(item.name))), ids = new Set(closure.relationships.filter(item => item.source === identity.part).map(item => item.id));
  const ext = mimeType === "image/png" ? "png" : mimeType === "image/jpeg" ? "jpg" : "gif";
  let serial = 1, part = "", relationshipId = "";
  do { await check(); part = `/ppt/media/pptkit-image-${serial}.${ext}`; relationshipId = `pptkitImage${serial++}`; } while (partNames.has(asciiKey(part)) || ids.has(relationshipId));
  const slidePart = original.inspection.parts.find(item => item.name === identity.part), relPart = original.inspection.parts.find(item => item.name === relation.relationshipPart);
  const payload = (entry: string) => { const value = original.payloads!.get(entry); if (!value) throw new PptxImportError("save-candidate-invalid", "Picture source payload is unavailable.", { part: entry }); return value; };
  if (!slidePart || !relPart) throw new PptxImportError("save-candidate-invalid", "Picture source parts are unavailable.", { part: identity.part });
  const xmlOptions = { limits: captured.xmlLimits, maxOutputBytes: savedLimits.maxPreparedOutputBytes, ...(signal ? { signal } : {}) };
  const slide = await replaceXmlRelationshipId(payload(slidePart.entryName), { elementIndex: identity.index, elementNamespace: context.a, elementLocal: "blip", attributeNamespace: context.r, attributeLocal: "embed", expected: identity.id, replacement: relationshipId }, xmlOptions);
  const target = `${"../".repeat(identity.part.split("/").length - 2)}${part.slice(1)}`;
  const relationships = await appendImageRelationship(payload(relPart.entryName), { id: relationshipId, target, dialect: context.entry.dialect }, xmlOptions);
  const types = await appendImageContentType(payload("[Content_Types].xml"), { part, mimeType }, xmlOptions);
  const changes = new Map([[slidePart.entryName, slide], [relPart.entryName, relationships], ["[Content_Types].xml", types]]);
  const output = await replaceArchiveEntries(originalBytes, [...changes].map(([name, bytes]) => ({ name, bytes })), { limits: captured.zipLimits, maxOutputBytes: savedLimits.maxPreparedOutputBytes, additions: [{ name: part.slice(1), bytes: replacement }], ...(signal ? { signal } : {}) });
  await check();
  const verified = await readOpcSource(output, { ...optionsCopy, zipLimits: { ...captured.zipLimits, maxInputBytes: Math.max(captured.zipLimits.maxInputBytes, output.length) } }, false, true);
  let validationBytes = validationUsage.decodedBytes ?? 0;
  const equal = async (a: Uint8Array, b: Uint8Array) => {
    validationBytes += b.length;
    if (validationBytes > savedLimits.maxSaveValidationBytes) throw new PptxImportError("save-limit-exceeded", "Picture candidate comparison exceeds maxSaveValidationBytes.");
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) { if (++validationUsage.work > savedLimits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Picture candidate comparison exceeds maxSaveValidationWork."); if (a[i] !== b[i]) return false; if (i % 65536 === 65535) await check(); }
    return true;
  };
  if (verified.payloads!.size !== original.payloads!.size + 1) throw new PptxImportError("save-candidate-invalid", "Picture candidate has unexpected archive entries.");
  for (const [entry, bytes] of original.payloads!) {
    validationBytes += bytes.length;
    if (validationBytes > savedLimits.maxSaveValidationBytes) throw new PptxImportError("save-limit-exceeded", "Picture candidate comparison exceeds maxSaveValidationBytes.");
    if (!await equal(changes.get(entry) ?? bytes, verified.payloads!.get(entry) ?? new Uint8Array())) throw new PptxImportError("save-candidate-invalid", "Picture candidate altered an unexpected source payload.", { part: entry });
  }
  if (!await equal(replacement, payloadFrom(verified.payloads!, part.slice(1)))) throw new PptxImportError("save-candidate-invalid", "Picture candidate media payload disagrees with its request.", { part });
  const next = verified.inspection.relationships.find(item => item.source === identity.part && item.id === relationshipId);
  const media = verified.inspection.parts.find(item => item.name === part);
  if (!next || next.type !== `${context.r}/image` || next.targetPart !== part || next.targetMode !== "Internal" || media?.contentType !== mimeType) throw new PptxImportError("save-candidate-invalid", "Picture candidate relationship or media type is incorrect.", { part });
  const project = async (current: Awaited<ReturnType<typeof structureContext>>, inspectionOptions: PptxInspectionOptions) => {
    const inspection = await readStructureFromContext(current), projection = await projectStructure(inspection, savedLimits.maxReportFindings, check), pictureSources: ImportedPictureSource[] = [];
    await projectContent(projection.document, projection.report, current, savedLimits, check, inspectionOptions, undefined, undefined, undefined, undefined, undefined, undefined, undefined, false, false, false, false, false, false, false, false, false, false, false, pictureSources);
    return { document: snapshot(projection.document, savedLimits.maxSnapshotValues), pictureSources };
  };
  const beforeProjection = await project(context, optionsCopy), afterOptions = { ...optionsCopy, zipLimits: { ...captured.zipLimits, maxInputBytes: Math.max(captured.zipLimits.maxInputBytes, output.length) } };
  const afterProjection = await project(await structureContext(output, afterOptions, true), afterOptions);
  verifyPictureProjection(beforeProjection.document, afterProjection.document, beforeProjection.pictureSources, afterProjection.pictureSources, targetBinding, part, relationshipId, savedLimits.maxSnapshotValues);
  await check(); return { bytes: output, part, relationshipId };
}

function payloadFrom(payloads: ReadonlyMap<string, Uint8Array>, name: string): Uint8Array { return payloads.get(name) ?? new Uint8Array(); }
