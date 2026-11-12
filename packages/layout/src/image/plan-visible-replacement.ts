import { planImageReplacement, type ImageReplacementPlan, type PresentationDocument } from "@pptkit/core";
import { resolveLayout } from "../engine/resolve-layout.js";
import type { LayoutComposition } from "../types/composition.js";

export type VisibleImageReplacementFit = "preserve" | "stretch" | "contain" | "cover";
export interface VisibleImageReplacementOptions {
  /** Optional composition is captured and evaluated before image fitting. */
  readonly composition?: readonly LayoutComposition[];
}

/** Plan a top-level slide image replacement using its current visible frame, optionally after composition. */
export function planVisibleImageReplacement(document: PresentationDocument, slideId: string, elementId: string, targetAssetId: string, fit: VisibleImageReplacementFit, options: VisibleImageReplacementOptions = {}): ImageReplacementPlan {
  if (![slideId, elementId, targetAssetId].every(value => typeof value === "string" && value.length > 0)) throw new TypeError("Slide, element and target asset IDs are required.");
  if (!["preserve", "stretch", "contain", "cover"].includes(fit)) throw new TypeError("An explicit visible-frame fit is required.");
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => key !== "composition")) throw new TypeError("Only composition is supported in visible-frame options.");
  const inputComposition = options.composition;
  const composition = inputComposition === undefined ? undefined : structuredClone(inputComposition);
  const slide = document.slides.find(value => value.id === slideId);
  if (!slide) throw new Error(`Unknown slide id "${slideId}".`);
  const image = slide.elements.find(value => value.id === elementId);
  if (!image || image.type !== "image") throw new Error("A top-level slide image is required.");
  const target = document.getAsset(targetAssetId);
  if (!target || target.kind !== "image") throw new Error(`Unknown image asset id "${targetAssetId}".`);
  const layout = resolveLayout(document, composition === undefined ? undefined : { composition });
  // Do not accept a partially applied composition or hide unrelated warnings.
  if (composition !== undefined && layout.analysis?.diagnostics.length) throw new Error("Visible-frame composition produced diagnostics and cannot be used as a replacement basis.");
  const resolved = layout.slides.find(value => value.id === slideId)?.elements.find(value => value.id === elementId);
  if (!resolved || resolved.type !== "image") throw new Error("Image could not be resolved on the selected slide.");
  if (resolved.fit === "contain" || resolved.fit === "cover") throw new Error("The original image fit is unresolved; intrinsic dimensions and nonzero frame are required.");
  if (resolved.box.width <= 0 || resolved.box.height <= 0) throw new Error("The original image has no positive visible frame.");
  // Keep authored properties and undo separate from the resolved geometry snapshot.
  const previous = planImageReplacement(image, target, { fit: "stretch" }).previous;
  const visibleImage = { ...image, box: resolved.box, fit: resolved.fit, crop: resolved.crop };
  const replacement = planImageReplacement(visibleImage, target, { fit }).replacement;
  return Object.freeze({ previous, replacement });
}
