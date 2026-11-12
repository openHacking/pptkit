import { createPresentation, type PresentationAssetReplacement } from "@pptkit/core";

export function reuseAndReplaceImage(first: Uint8Array, next: Uint8Array) {
  const document = createPresentation();
  const asset = document.registerAsset({ kind: "image", source: { type: "bytes", value: first } }, { deduplicate: "content" });
  const slide = document.addSlide();
  const element = slide.addElement({ type: "image", assetId: asset.id, box: { x: 0, y: 0, width: 200, height: 100 } });
  const replacement: PresentationAssetReplacement = { kind: "image", source: { type: "bytes", value: next } };
  document.replaceAsset(asset.id, replacement); // All uses retain the same asset ID.
  const uses = document.getAssetReferences(asset.id);
  slide.removeElement(element.id);
  // A consumer can retain assets referenced outside Core's known image graph.
  document.removeUnusedAssets({ retainAssetIds: [asset.id] });
  const removed = document.removeAsset(asset.id);
  return { document, uses, removed };
}
