import { planImageReplacement, type GroupElementInput, type PresentationDocument, type PresentationElement } from "@pptkit/core";
import { resolveLayout } from "../engine/resolve-layout.js";
import type { LayoutElement } from "../types/layout.js";
import type { VisibleImageReplacementFit, VisibleImageReplacementOptions } from "./plan-visible-replacement.js";

export interface NestedVisibleImageReplacementPlan {
  readonly rootElementId: string;
  readonly imageId: string;
  /** Entire top-level group snapshots; pass either to the slide's replaceElement. */
  readonly replacement: Readonly<GroupElementInput>;
  readonly previous: Readonly<GroupElementInput>;
  /** Changed ancestors whose composition geometry is committed, outermost first. */
  readonly committedAncestorIds?: readonly string[];
}
export interface NestedVisibleImageReplacementOptions extends VisibleImageReplacementOptions { readonly commitAncestorGeometry?: boolean }

function freezeTree<T>(value: T): T {
  const stack: unknown[] = [value], seen = new WeakSet<object>();
  while (stack.length) {
    const current = stack.pop();
    if (current === null || typeof current !== "object" || seen.has(current)) continue;
    seen.add(current); for (const child of Object.values(current)) stack.push(child); Object.freeze(current);
  }
  return value;
}

/** Plan a nested image in its immediate parent coordinates, retaining the top-level group for atomic commit. */
export function planNestedVisibleImageReplacement(document: PresentationDocument, slideId: string, imageId: string, targetAssetId: string, fit: VisibleImageReplacementFit, options: NestedVisibleImageReplacementOptions = {}): NestedVisibleImageReplacementPlan {
  if (![slideId, imageId, targetAssetId].every(value => typeof value === "string" && value.length > 0)) throw new TypeError("Slide, image and target asset IDs are required.");
  if (!["preserve", "stretch", "contain", "cover"].includes(fit)) throw new TypeError("An explicit visible-frame fit is required.");
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => key !== "composition" && key !== "commitAncestorGeometry")) throw new TypeError("Invalid nested visible-frame options.");
  const { composition: inputComposition, commitAncestorGeometry = false } = options;
  if (typeof commitAncestorGeometry !== "boolean") throw new TypeError("commitAncestorGeometry must be a boolean.");
  const composition = inputComposition === undefined ? undefined : structuredClone(inputComposition);
  const slide = document.slides.find(value => value.id === slideId); if (!slide) throw new Error(`Unknown slide id "${slideId}".`);
  const target = document.getAsset(targetAssetId); if (!target || target.kind !== "image") throw new Error(`Unknown image asset id "${targetAssetId}".`);
  const stack: { node: PresentationElement; path: number[] }[] = slide.elements.map((node, index) => ({ node, path: [index] })).reverse();
  let match: { node: PresentationElement; path: number[] } | undefined;
  while (stack.length) {
    const item = stack.pop()!;
    if (item.node.id === imageId) { match = item; break; }
    if (item.node.type === "group") for (let i = item.node.children.length - 1; i >= 0; i--) stack.push({ node: item.node.children[i]!, path: [...item.path, i] });
  }
  if (!match || match.node.type !== "image" || match.path.length < 2) throw new Error("An image nested in a top-level slide group is required.");
  const root = slide.elements[match.path[0]!]!; if (root.type !== "group") throw new Error("Nested image group root is unavailable.");
  const base = resolveLayout(document), result = composition === undefined ? base : resolveLayout(document, { composition });
  if (composition !== undefined && result.analysis?.diagnostics.length) throw new Error("Visible-frame composition produced diagnostics and cannot be used as a replacement basis.");
  const baseElements = base.slides.find(value => value.id === slideId)?.elements;
  const elements = result.slides.find(value => value.id === slideId)?.elements;
  if (!baseElements || !elements) throw new Error("Resolved slide is unavailable.");
  let current: LayoutElement | undefined = elements[match.path[0]!], baseline: LayoutElement | undefined = baseElements[match.path[0]!];
  const changed: { depth: number; id: string; box: Extract<LayoutElement, { type: "group" }>["box"]; coordinateSize: Extract<LayoutElement, { type: "group" }>["coordinateSize"]; transform: Extract<LayoutElement, { type: "group" }>["transform"] }[] = [];
  for (let depth = 1; depth < match.path.length; depth++) {
    if (!current || !baseline || current.type !== "group" || baseline.type !== "group" || current.id !== baseline.id) throw new Error("Resolved group path changed.");
    // A group frame change would change the parent coordinate transform; this plan cannot commit it.
    if (JSON.stringify([current.box, current.coordinateSize, current.transform]) !== JSON.stringify([baseline.box, baseline.coordinateSize, baseline.transform])) {
      if (!commitAncestorGeometry) throw new Error("Composition changed an ancestor group frame; the nested visible frame cannot be committed alone.");
      changed.push({ depth: depth - 1, id: current.id, box: current.box, coordinateSize: current.coordinateSize, transform: current.transform });
    }
    current = current.children[match.path[depth]!]; baseline = baseline.children[match.path[depth]!];
  }
  if (!current || current.type !== "image" || current.id !== imageId) throw new Error("Resolved nested image path changed.");
  if (current.fit === "contain" || current.fit === "cover") throw new Error("The original image fit is unresolved; intrinsic dimensions and nonzero frame are required.");
  if (current.box.width <= 0 || current.box.height <= 0) throw new Error("The original image has no positive visible frame.");
  const imagePlan = planImageReplacement({ ...match.node, box: current.box, fit: current.fit, crop: current.crop }, target, { fit });
  const previous = structuredClone(root) as GroupElementInput, replacement = structuredClone(root) as GroupElementInput;
  let parent = replacement;
  const applyGeometry = (depth: number, group: GroupElementInput) => {
    const geometry = changed.find(item => item.depth === depth);
    if (geometry) { group.box = structuredClone(geometry.box); group.coordinateSize = structuredClone(geometry.coordinateSize); group.transform = structuredClone(geometry.transform); }
  };
  applyGeometry(0, parent);
  for (let depth = 1; depth < match.path.length - 1; depth++) {
    const next = parent.children[match.path[depth]!]!;
    if (next.type !== "group") throw new Error("Nested image group path changed."); parent = next; applyGeometry(depth, parent);
  }
  parent.children[match.path[match.path.length - 1]!] = imagePlan.replacement;
  return Object.freeze({ rootElementId: root.id, imageId, previous: freezeTree(previous), replacement: freezeTree(replacement), ...(commitAncestorGeometry ? { committedAncestorIds: Object.freeze(changed.map(item => item.id)) } : {}) });
}
