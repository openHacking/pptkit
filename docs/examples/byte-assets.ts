import { createPresentation } from "@pptkit/core";
import { generatePptx, type PptxAssetResolver } from "@pptkit/pptx-exporter";
import { renderPresentationToSvg } from "@pptkit/svg-renderer";

export async function presentImage(bytes: Uint8Array, signal?: AbortSignal) {
  const document = createPresentation();
  const asset = document.registerAsset({
    kind: "image", source: { type: "bytes", value: bytes }, mimeType: "image/png",
    width: 800, height: 400,
  });
  document.addSlide().addElement({ type: "image", assetId: asset.id,
    box: { x: 40, y: 40, width: 400, height: 200 }, fit: "contain" });
  const output = await generatePptx(document, { maxAssetBytes: 16 * 1024 * 1024, ...(signal ? { signal } : {}) });
  const preview = await renderPresentationToSvg(document);
  return { output, preview };
}

export function assetMapResolver(images: ReadonlyMap<string, Uint8Array>): PptxAssetResolver {
  return (asset, context) => {
    context.signal?.throwIfAborted();
    const bytes = images.get(asset.id);
    // Missing entries are explicitly omitted; no implicit network/path fallback.
    return bytes ? { bytes } : undefined;
  };
}
