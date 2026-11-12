import type { PresentationAsset } from "../types/asset.js";
import type { ImageElementInput } from "../types/element.js";
import { deepClone, deepFreeze } from "../utils/clone.js";

/** Explicit policies act on an authored frame, never on an already-resolved visual box. */
export type ImageReplacementPolicy =
  | { readonly fit: "preserve" | "stretch" | "contain" | "cover" }
  | { readonly fit: "size"; readonly width: number; readonly height: number };
export interface ImageReplacementPlan {
  /** Detached snapshots suitable for replaceElement and undo respectively. */
  readonly replacement: Readonly<ImageElementInput>;
  readonly previous: Readonly<ImageElementInput>;
}

/** Plan one image retarget without changing either asset or the presentation. */
export function planImageReplacement(image: ImageElementInput, asset: PresentationAsset, policy: ImageReplacementPolicy): ImageReplacementPlan {
  // Clone before validation and return: subsequent caller mutation cannot alter the plan.
  const original = deepClone(image), choice = deepClone(policy);
  const target = asset && { id: asset.id, kind: asset.kind, width: asset.width, height: asset.height };
  if (!original || original.type !== "image" || typeof original.assetId !== "string" || !original.assetId) throw new TypeError("A valid image element is required.");
  if (!target || target.kind !== "image" || typeof target.id !== "string" || !target.id) throw new TypeError("A target image asset with an ID is required.");
  if (!choice || typeof choice !== "object" || !["preserve", "stretch", "contain", "cover", "size"].includes(choice.fit)) throw new TypeError("An explicit image replacement policy is required.");
  const box = original.box;
  if (!box || ![box.x, box.y, box.width, box.height].every(Number.isFinite) || box.width < 0 || box.height < 0) throw new RangeError("Image frame must have finite coordinates and nonnegative dimensions.");
  if (original.fit !== undefined && !["stretch", "contain", "cover", "crop"].includes(original.fit)) throw new TypeError("Invalid existing image fit.");
  if (original.crop !== undefined) {
    const crop = original.crop;
    if (!crop || typeof crop !== "object") throw new TypeError("Invalid image crop.");
    const { left = 0, top = 0, right = 0, bottom = 0 } = crop;
    if (![left, top, right, bottom].every(value => typeof value === "number" && Number.isFinite(value) && value >= 0 && value <= 1) || left + right >= 1 || top + bottom >= 1) throw new RangeError("Image crop must leave a positive visible area.");
  }
  const fit = choice.fit === "preserve" ? original.fit : choice.fit === "size" ? "stretch" : choice.fit;
  if (fit === "contain" || fit === "cover") {
    const aspect = target.width! / target.height!, frameAspect = box.width / box.height;
    if (!Number.isFinite(target.width) || !Number.isFinite(target.height) || target.width! <= 0 || target.height! <= 0 || !Number.isFinite(aspect) || aspect <= 0) throw new RangeError("Target dimensions and aspect must be positive and finite for contain or cover.");
    if (box.width <= 0 || box.height <= 0 || !Number.isFinite(frameAspect) || frameAspect <= 0) throw new RangeError("Image frame dimensions and aspect must be positive and finite for contain or cover.");
    const visible = Math.min(aspect, frameAspect) / Math.max(aspect, frameAspect);
    if (!Number.isFinite(visible) || visible <= 0 || (fit === "cover" && 1 - visible === 1)) throw new RangeError("Image aspect ratios cannot be resolved safely.");
    if (fit === "contain") {
      const fittedWidth = aspect > frameAspect ? box.width : box.height * aspect;
      const fittedHeight = aspect > frameAspect ? box.width / aspect : box.height;
      const x = box.x + (box.width - fittedWidth) / 2, y = box.y + (box.height - fittedHeight) / 2;
      if (![fittedWidth, fittedHeight, x, y].every(Number.isFinite) || fittedWidth <= 0 || fittedHeight <= 0) throw new RangeError("Contained image frame cannot be resolved safely.");
    }
  }
  if (choice.fit === "size" && (![choice.width, choice.height].every(Number.isFinite) || choice.width <= 0 || choice.height <= 0)) throw new RangeError("Replacement size must be positive and finite.");
  const replacement = { ...original, assetId: target.id };
  if (choice.fit !== "preserve") {
    replacement.fit = choice.fit === "size" ? "stretch" : choice.fit;
    delete replacement.crop;
    if (choice.fit === "size") replacement.box = { ...box, width: choice.width, height: choice.height };
  }
  return deepFreeze({ previous: original, replacement });
}
