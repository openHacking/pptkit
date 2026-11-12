import type { LayoutResult } from "@pptkit/layout";
import type { SvgSlideReadingOrder } from "../types/public.js";
import { captureReadingOrders } from "./reading-order.js";

/** Caller-owned revision must change whenever external rendering inputs change. */
export interface SvgPreviewSnapshotOptions {
  readonly environmentRevision: string;
  readonly idNamespace?: string;
  readonly readingOrder?: readonly SvgSlideReadingOrder[];
}

export interface SvgPreviewSnapshot {
  readonly schemaVersion: "pptkit-svg-preview-snapshot-v1";
  readonly shared: string;
  readonly slides: readonly { readonly slideId: string; readonly index: number; readonly dependencies: string }[];
}

export interface SvgPreviewInvalidation {
  readonly reusable: readonly { readonly slideId: string; readonly index: number }[];
  readonly render: readonly { readonly slideId: string; readonly index: number }[];
  readonly removedSlideIds: readonly string[];
}

function serialize(value: unknown): string {
  return JSON.stringify(value, (_key, item: unknown) => {
    if ((typeof item === "number" && !Number.isFinite(item)) || typeof item === "function" || typeof item === "symbol" || typeof item === "bigint") throw new TypeError("Preview snapshot dependencies must be finite serializable Layout data.");
    return item;
  });
}

/** Capture conservative SVG string dependencies, detached from the mutable Layout. */
export function captureSvgPreviewSnapshot(layout: LayoutResult, options: SvgPreviewSnapshotOptions): SvgPreviewSnapshot {
  if (!options || typeof options.environmentRevision !== "string" || !options.environmentRevision) throw new TypeError("environmentRevision must be a nonempty string.");
  if (options.idNamespace !== undefined && (typeof options.idNamespace !== "string" || !options.idNamespace)) throw new TypeError("idNamespace must be a nonempty string when provided.");
  const orders = captureReadingOrders(layout, options.readingOrder);
  const seen = new Set<string>();
  const layouts = new Map(layout.layouts.map(item => [item.id, item]));
  const masters = new Map((layout.masters ?? []).map(item => [item.id, item]));
  // These fields may affect any page or asset resolver. Extra invalidation is intentional.
  const shared = serialize([layout.size, layout.metadata.title, layout.theme, layout.assets, options.idNamespace, options.environmentRevision]);
  const slides = layout.slides.map((slide, index) => {
    if (typeof slide.id !== "string" || !slide.id || seen.has(slide.id)) throw new TypeError("Preview snapshots require distinct nonempty slide IDs.");
    seen.add(slide.id);
    const owner = layouts.get(slide.layoutId);
    const master = owner?.masterId === undefined ? undefined : masters.get(owner.masterId);
    const order = orders.get(slide.id);
    return Object.freeze({ slideId: slide.id, index, dependencies: serialize(order === undefined ? [slide, owner, master] : [slide, owner, master, order]) });
  });
  return Object.freeze({ schemaVersion: "pptkit-svg-preview-snapshot-v1" as const, shared, slides: Object.freeze(slides) });
}

/** Compare two captured snapshots; a moved page is invalidated because SVG IDs and title use its index. */
export function planSvgPreviewInvalidation(previous: SvgPreviewSnapshot, next: SvgPreviewSnapshot): SvgPreviewInvalidation {
  if (previous.schemaVersion !== "pptkit-svg-preview-snapshot-v1" || next.schemaVersion !== "pptkit-svg-preview-snapshot-v1") throw new TypeError("Unsupported SVG preview snapshot.");
  const old = new Map(previous.slides.map(item => [item.slideId, item]));
  const nextIds = new Set(next.slides.map(item => item.slideId));
  const reusable: { slideId: string; index: number }[] = [], render: { slideId: string; index: number }[] = [];
  for (const item of next.slides) {
    const prior = old.get(item.slideId);
    const target = previous.shared === next.shared && prior?.index === item.index && prior.dependencies === item.dependencies ? reusable : render;
    target.push(Object.freeze({ slideId: item.slideId, index: item.index }));
  }
  return Object.freeze({ reusable: Object.freeze(reusable), render: Object.freeze(render), removedSlideIds: Object.freeze(previous.slides.filter(item => !nextIds.has(item.slideId)).map(item => item.slideId)) });
}
