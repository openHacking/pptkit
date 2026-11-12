import type { NormalizedElement, NormalizedPresentation } from "@pptkit/core";
import { PptxImportError, type ImportedPictureSource, type PptxSaveChange, type PptxSessionLimits } from "../types/import.js";
import { rasterMetadata } from "./image-metadata.js";
import { equalState } from "./state.js";

/** Exactly one new byte-backed asset and one bound image retarget; all other state is equal. */
export async function analyzePictureEdit(baseline: NormalizedPresentation, candidate: NormalizedPresentation, sources: readonly ImportedPictureSource[], limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>): Promise<{ binding: ImportedPictureSource; bytes: Uint8Array; mimeType: "image/png" | "image/jpeg" | "image/gif"; changes: PptxSaveChange[] }> {
  const unsupported = (message: string): never => { throw new PptxImportError("save-candidate-invalid", message); };
  if (candidate.assets.length !== baseline.assets.length + 1 || !equalState(baseline.assets, candidate.assets.slice(0, baseline.assets.length), limits.maxSnapshotValues)) return unsupported("Picture replacement requires one new asset and retains existing assets unchanged.");
  const asset = candidate.assets.at(-1)!;
  if (baseline.assets.some(item => item.id === asset.id) || asset.kind !== "image" || asset.source.type !== "bytes" || asset.source.encoding !== "base64" || !["image/png", "image/jpeg", "image/gif"].includes(asset.mimeType ?? "")) return unsupported("Picture replacement requires a new byte-backed PNG, JPEG or GIF asset.");
  const encoded = asset.source.value;
  if (encoded.length > Math.ceil(limits.maxImportedAssetBytes / 3) * 4 || encoded.length > Math.ceil(limits.maxPreparedOutputBytes / 3) * 4) throw new PptxImportError("save-limit-exceeded", "Picture replacement exceeds the byte budget.");
  if (!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(encoded)) return unsupported("Picture asset bytes require canonical base64.");
  let decoded: string;
  try { decoded = atob(encoded); if (btoa(decoded) !== encoded) return unsupported("Picture asset bytes require canonical base64."); } catch { return unsupported("Picture asset bytes require canonical base64."); }
  const bytes = new Uint8Array(decoded.length); for (let i = 0; i < decoded.length; i++) bytes[i] = decoded.charCodeAt(i);
  if (bytes.length > limits.maxImportedAssetBytes || bytes.length > limits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Picture replacement exceeds the byte budget.");
  const metadata = await rasterMetadata(bytes, checkpoint), mimeType = asset.mimeType as "image/png" | "image/jpeg" | "image/gif";
  if (!metadata || metadata.mimeType !== mimeType || metadata.width !== asset.width || metadata.height !== asset.height || !equalState(asset.accessibility, { decorative: false }, limits.maxSnapshotValues) || asset.dedupeKey !== undefined) return unsupported("Picture asset MIME, dimensions and metadata must match the imported image representation.");
  const expected = structuredClone(baseline), sourceByKey = new Map(sources.map(item => [JSON.stringify([item.slideId, item.elementId]), item]));
  if (sourceByKey.size !== sources.length) return unsupported("Picture source identities are ambiguous.");
  let chosen: { binding: ImportedPictureSource; path: string } | undefined, work = 0;
  for (let i = 0; i < baseline.slides.length; i++) {
    const leftSlide = baseline.slides[i]!, rightSlide = candidate.slides[i], expectedSlide = expected.slides[i]!;
    if (!rightSlide) continue;
    const stack: { left: NormalizedElement; right: NormalizedElement | undefined; expected: NormalizedElement; path: string }[] = leftSlide.elements.map((left, j) => ({ left, right: rightSlide.elements[j], expected: expectedSlide.elements[j]!, path: `/slides/${i}/elements/${j}` }));
    while (stack.length) {
      if (++work > limits.maxSaveAnalysisWork) throw new PptxImportError("save-limit-exceeded", "Picture save analysis exceeds maxSaveAnalysisWork.");
      await checkpoint(); const { left, right, expected: writable, path } = stack.pop()!;
      if (left.type === "group" && writable.type === "group") for (let j = 0; j < left.children.length; j++) stack.push({ left: left.children[j]!, right: right?.type === "group" ? right.children[j] : undefined, expected: writable.children[j]!, path: `${path}/children/${j}` });
      if (left.type !== "image" || right?.type !== "image" || writable.type !== "image" || left.assetId === right.assetId) continue;
      const binding = sourceByKey.get(JSON.stringify([leftSlide.id, left.id]));
      if (chosen || !binding || binding.assetId !== left.assetId || right.assetId !== asset.id) return unsupported("Exactly one source-bound picture must reference the new asset.");
      if (binding.hasSvgExtension || binding.hasExternalLink) return unsupported("Pictures with an SVG representation or external link are not supported for replacement.");
      writable.assetId = asset.id; chosen = { binding, path: `${path}/assetId` };
    }
  }
  if (!chosen) return unsupported("Exactly one source-bound picture must reference the new asset.");
  expected.assets.push(asset);
  if (!equalState(expected, candidate, limits.maxSnapshotValues)) return unsupported("Picture replacement cannot be combined with other document changes.");
  const changes: PptxSaveChange[] = [
    { kind: "asset", documentPath: `/assets/${baseline.assets.length}` },
    { kind: "asset", documentPath: chosen.path, slideId: chosen.binding.slideId, elementId: chosen.binding.elementId, source: chosen.binding.source },
  ];
  if (changes.length > limits.maxSaveChanges) throw new PptxImportError("save-limit-exceeded", "Picture changes exceed maxSaveChanges.");
  return { binding: chosen.binding, bytes, mimeType, changes };
}
