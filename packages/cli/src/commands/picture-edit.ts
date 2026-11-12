import type { ImportedPresentation } from "@pptkit/pptx-parser";
import type { PresentationElement, PresentationElementInput, PresentationSlide } from "@pptkit/core";
import { CliError } from "./arguments.js";
import { checkAbort } from "./read-file.js";

export function applyPictureEdit(session: ImportedPresentation, bytes: Uint8Array, image: Uint8Array, sourceSha256: string, imageSha256: string, signal?: AbortSignal): void {
  let value: unknown;
  try { value = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)); }
  catch { throw new CliError("picture-edit-invalid", "Picture edit must be valid UTF-8 JSON.", 3); }
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new CliError("picture-edit-invalid", "Picture edit must be an object.", 3);
  const v = value as Record<string, unknown>, hash = (x: unknown): x is string => typeof x === "string" && /^[0-9a-f]{64}$/.test(x);
  if (v.version !== "pptkit-picture-edit-v1" || Object.keys(v).some(k => !["version", "sourceSha256", "replacementSha256", "slideId", "elementId", "expectedAssetId", "mimeType", "width", "height"].includes(k)) || !hash(v.sourceSha256) || !hash(v.replacementSha256) || typeof v.slideId !== "string" || !v.slideId || typeof v.elementId !== "string" || !v.elementId || typeof v.expectedAssetId !== "string" || !v.expectedAssetId || typeof v.mimeType !== "string" || !["image/png", "image/jpeg", "image/gif"].includes(v.mimeType) || !Number.isSafeInteger(v.width) || (v.width as number) <= 0 || !Number.isSafeInteger(v.height) || (v.height as number) <= 0) throw new CliError("picture-edit-invalid", "Expected picture identity, hashes, supported MIME and positive integer dimensions.", 3);
  if (v.sourceSha256 !== sourceSha256 || v.replacementSha256 !== imageSha256) throw new CliError("picture-edit-source-mismatch", "Picture manifest hashes do not match the captured source and replacement.", 3);
  if (!image.byteLength) throw new CliError("picture-edit-invalid", "Replacement image must not be empty.", 3);
  const bindings = session.report.pictureSources?.filter(b => b.slideId === v.slideId && b.elementId === v.elementId && b.assetId === v.expectedAssetId) ?? [];
  if (bindings.length !== 1) throw new CliError("picture-edit-target-unbound", "Picture source identity is missing or ambiguous.", 3);
  const owners: { slide: PresentationSlide; root: PresentationElement; path: number[] }[] = [];
  for (const slide of session.document.slides) if (slide.id === v.slideId) for (const root of slide.elements) {
    const pending: { element: PresentationElement; path: number[] }[] = [{ element: root, path: [] }];
    while (pending.length) {
      checkAbort(signal); const { element, path } = pending.pop()!;
      if (element.type === "image" && element.id === v.elementId && element.assetId === v.expectedAssetId) owners.push({ slide, root, path });
      else if (element.type === "group") for (let i = element.children.length - 1; i >= 0; i--) pending.push({ element: element.children[i]!, path: [...path, i] });
    }
  }
  if (owners.length !== 1) throw new CliError("picture-edit-target-unbound", "Picture owner is missing or ambiguous.", 3);
  const owner = owners[0]!, root = structuredClone(owner.root) as PresentationElementInput; let element = root;
  for (const index of owner.path) { if (element.type !== "group" || !element.children[index]) throw new CliError("picture-edit-target-unbound", "Picture ancestry is inconsistent.", 3); element = element.children[index]!; }
  if (element.type !== "image" || element.id !== v.elementId || element.assetId !== v.expectedAssetId) throw new CliError("picture-edit-target-unbound", "Picture ancestry is inconsistent.", 3);
  const ids = new Set(session.document.assets.map(asset => asset.id)); let suffix = 1; while (ids.has(`pptkit-cli-image-${suffix}`)) { checkAbort(signal); suffix++; }
  const id = `pptkit-cli-image-${suffix}`;
  checkAbort(signal);
  session.document.registerAsset({ id, kind: "image", source: { type: "bytes", value: image }, mimeType: v.mimeType as string, width: v.width as number, height: v.height as number });
  element.assetId = id; owner.slide.replaceElement(owner.root.id, root);
}
