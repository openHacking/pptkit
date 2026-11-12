import { planImageReplacement, type PresentationDocument, type PresentationElement, type SlideMasterInput, type SlideLayoutInput } from "@pptkit/core";
import { resolveLayout } from "../engine/resolve-layout.js";
import type { LayoutElement } from "../types/layout.js";
import type { VisibleImageReplacementFit, VisibleImageReplacementOptions } from "./plan-visible-replacement.js";

export type SharedImageOwner = { readonly kind: "master" | "layout"; readonly id: string };
export type SharedVisibleImageReplacementPlan =
  | { readonly kind: "master"; readonly ownerId: string; readonly imageId: string; readonly previous: Readonly<SlideMasterInput>; readonly replacement: Readonly<SlideMasterInput>; readonly committedAncestorIds?: readonly string[] }
  | { readonly kind: "layout"; readonly ownerId: string; readonly imageId: string; readonly previous: Readonly<SlideLayoutInput>; readonly replacement: Readonly<SlideLayoutInput>; readonly committedAncestorIds?: readonly string[] };
export interface SharedVisibleImageReplacementOptions extends VisibleImageReplacementOptions { readonly commitAncestorGeometry?: boolean }

function freezeTree<T>(value: T): T {
  const stack: unknown[] = [value], seen = new WeakSet<object>();
  while (stack.length) {
    const current = stack.pop(); if (current === null || typeof current !== "object" || seen.has(current)) continue;
    seen.add(current); for (const child of Object.values(current)) stack.push(child); Object.freeze(current);
  }
  return value;
}

/** Shared definition replacement affects all pages using that master or layout. */
export function planSharedVisibleImageReplacement(document: PresentationDocument, owner: SharedImageOwner, imageId: string, targetAssetId: string, fit: VisibleImageReplacementFit, options: SharedVisibleImageReplacementOptions = {}): SharedVisibleImageReplacementPlan {
  if (!owner || typeof owner !== "object") throw new TypeError("A shared owner is required.");
  const { kind, id: ownerId } = owner;
  if (!["master", "layout"].includes(kind) || typeof ownerId !== "string" || !ownerId || typeof imageId !== "string" || !imageId || typeof targetAssetId !== "string" || !targetAssetId) throw new TypeError("A shared owner, image and target asset IDs are required.");
  if (!["preserve", "stretch", "contain", "cover"].includes(fit)) throw new TypeError("An explicit visible-frame fit is required.");
  if (!options || typeof options !== "object" || Array.isArray(options) || Object.keys(options).some(key => key !== "composition" && key !== "commitAncestorGeometry")) throw new TypeError("Invalid shared visible-frame options.");
  const { composition: inputComposition, commitAncestorGeometry = false } = options;
  if (typeof commitAncestorGeometry !== "boolean") throw new TypeError("commitAncestorGeometry must be a boolean.");
  const composition = inputComposition === undefined ? undefined : structuredClone(inputComposition);
  const definition = kind === "master" ? document.masters?.find(value => value.id === ownerId) : document.layouts.find(value => value.id === ownerId);
  if (!definition) throw new Error(`Unknown ${kind} id "${ownerId}".`);
  const target = document.getAsset(targetAssetId); if (!target || target.kind !== "image") throw new Error(`Unknown image asset id "${targetAssetId}".`);
  const stack: { node: PresentationElement; path: number[] }[] = (definition.elements ?? []).map((node, index) => ({ node, path: [index] })).reverse();
  let match: { node: PresentationElement; path: number[] } | undefined;
  while (stack.length) {
    const item = stack.pop()!; if (item.node.id === imageId) { match = item; break; }
    if (item.node.type === "group") for (let i = item.node.children.length - 1; i >= 0; i--) stack.push({ node: item.node.children[i]!, path: [...item.path, i] });
  }
  if (!match || match.node.type !== "image") throw new Error("An image owned by the selected definition is required.");
  const base = resolveLayout(document), result = composition === undefined ? base : resolveLayout(document, { composition });
  if (composition !== undefined && result.analysis?.diagnostics.length) throw new Error("Visible-frame composition produced diagnostics and cannot be used as a replacement basis.");
  const ownerElements = (layout: typeof base): LayoutElement[] | undefined => kind === "master" ? layout.masters?.find(value => value.id === ownerId)?.elements : layout.layouts.find(value => value.id === ownerId)?.elements;
  const baseElements = ownerElements(base), elements = ownerElements(result); if (!baseElements || !elements) throw new Error("Resolved shared definition is unavailable.");
  let current: LayoutElement | undefined = elements[match.path[0]!], baseline: LayoutElement | undefined = baseElements[match.path[0]!];
  const changed: { depth: number; id: string; box: Extract<LayoutElement, { type: "group" }>["box"]; coordinateSize: Extract<LayoutElement, { type: "group" }>["coordinateSize"]; transform: Extract<LayoutElement, { type: "group" }>["transform"] }[] = [];
  for (let depth = 1; depth < match.path.length; depth++) {
    if (!current || !baseline || current.type !== "group" || baseline.type !== "group" || current.id !== baseline.id) throw new Error("Resolved group path changed.");
    if (JSON.stringify([current.box, current.coordinateSize, current.transform]) !== JSON.stringify([baseline.box, baseline.coordinateSize, baseline.transform])) {
      if (!commitAncestorGeometry) throw new Error("Composition changed an ancestor group frame; the shared visible frame cannot be committed alone.");
      changed.push({ depth: depth - 1, id: current.id, box: current.box, coordinateSize: current.coordinateSize, transform: current.transform });
    }
    current = current.children[match.path[depth]!]; baseline = baseline.children[match.path[depth]!];
  }
  if (!current || current.type !== "image" || current.id !== imageId) throw new Error("Resolved shared image path changed.");
  if (current.fit === "contain" || current.fit === "cover") throw new Error("The original image fit is unresolved; intrinsic dimensions and nonzero frame are required.");
  if (current.box.width <= 0 || current.box.height <= 0) throw new Error("The original image has no positive visible frame.");
  const image = planImageReplacement({ ...match.node, box: current.box, fit: current.fit, crop: current.crop }, target, { fit }).replacement;
  const previous = structuredClone(definition), replacement = structuredClone(definition);
  const elementsToEdit = replacement.elements! as PresentationElement[];
  if (match.path.length === 1) elementsToEdit[match.path[0]!] = image as PresentationElement;
  else {
    let parent = elementsToEdit[match.path[0]!]!;
    const applyGeometry = (depth: number, group: Extract<PresentationElement, { type: "group" }>) => {
      const geometry = changed.find(item => item.depth === depth);
      if (geometry) { group.box = structuredClone(geometry.box); group.coordinateSize = structuredClone(geometry.coordinateSize); group.transform = structuredClone(geometry.transform); }
    };
    if (parent.type !== "group") throw new Error("Shared image group path changed."); applyGeometry(0, parent);
    for (let depth = 1; depth < match.path.length - 1; depth++) {
      if (parent.type !== "group") throw new Error("Shared image group path changed."); parent = parent.children[match.path[depth]!]!;
      if (parent.type !== "group") throw new Error("Shared image group path changed."); applyGeometry(depth, parent);
    }
    if (parent.type !== "group") throw new Error("Shared image group path changed.");
    (parent.children as PresentationElement[])[match.path[match.path.length - 1]!] = image as PresentationElement;
  }
  const committed = commitAncestorGeometry ? { committedAncestorIds: Object.freeze(changed.map(item => item.id)) } : {};
  if (kind === "master") return Object.freeze({ kind, ownerId, imageId, previous: freezeTree(previous as SlideMasterInput), replacement: freezeTree(replacement as SlideMasterInput), ...committed });
  return Object.freeze({ kind: "layout", ownerId, imageId, previous: freezeTree(previous as SlideLayoutInput), replacement: freezeTree(replacement as SlideLayoutInput), ...committed });
}
