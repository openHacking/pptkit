import { inspectMp4MovieDuration } from "@pptkit/media-inspection";
import type { NormalizedAsset, NormalizedElement, NormalizedPresentation } from "@pptkit/core";
import { decodeAssetBytes } from "../assets/load-bytes.js";

export interface Mp4VideoResource { readonly data: Uint8Array; readonly durationMs: number }

/** Experimental profile: bounded visible top-level videos, no authored timing. */
export function planMp4Video(document: NormalizedPresentation, mode: "iso-bmff-v1" | undefined, maxBytes: number | undefined, signal?: AbortSignal, external?: Map<string, NormalizedAsset>, resolved?: ReadonlyMap<string, Uint8Array>): ReadonlyMap<string, Mp4VideoResource> {
  const assets = new Map(document.assets.map(asset => [asset.id, asset])), resources = new Map<string, Mp4VideoResource>();
  const reject = (reason: string): never => { throw new TypeError(`PPTX video export: ${reason}`); };
  const check = (element: NormalizedElement, slideId?: string, nested = false) => {
    if (element.type === "group") { for (const child of element.children) check(child, slideId, true); return; }
    if (element.type !== "video") return;
    if (element.playbackStart !== undefined && element.playbackStart !== "click-sequence" && element.playbackStart !== "automatic") return reject("playbackStart must be click-sequence or automatic.");
    if (mode !== "iso-bmff-v1") return reject("requires videoMode iso-bmff-v1; video cannot be silently omitted.");
    if (!slideId || nested || element.hidden || element.action !== undefined || element.placeholderKey !== undefined || element.opacity !== 1 || !element.box.width || !element.box.height) return reject("only visible, nontransparent slide-local top-level video with nonzero geometry and no action or placeholder binding is supported.");
    const asset = assets.get(element.videoAssetId);
    if (!asset || asset.kind !== "video") return reject("video requires a video asset.");
    if (asset.mimeType?.trim().toLowerCase() !== "video/mp4") return reject("a declared video/mp4 MIME type is required.");
    if (resources.has(asset.id)) return;
    if (asset.source.type !== "bytes" && !resolved?.has(asset.id)) {
      if (!external) return reject("video requires a byte-backed asset; paths and links are not loaded.");
      external.set(asset.id, asset); return;
    }
    const ceiling = Math.min(maxBytes ?? 64 * 1024 * 1024, 64 * 1024 * 1024);
    const length = asset.source.type === "bytes" ? asset.source.value.length / 4 * 3 - (asset.source.value.endsWith("==") ? 2 : asset.source.value.endsWith("=") ? 1 : 0) : resolved!.get(asset.id)!.byteLength;
    if (length > ceiling) throw new RangeError("Video exceeds the MP4 inspection or maxAssetBytes limit.");
    signal?.throwIfAborted();
    const data = asset.source.type === "bytes" ? decodeAssetBytes(asset.source.value) : resolved!.get(asset.id)!, durationMs = inspectMp4MovieDuration(data, { maxBytes: ceiling, ...(signal ? { signal } : {}) });
    resources.set(asset.id, { data, durationMs });
  };
  for (const layout of document.layouts) for (const element of layout.elements) check(element);
  for (const master of document.masters ?? []) for (const element of master.elements ?? []) check(element);
  for (const slide of document.slides) {
    const videos = slide.elements.filter(element => element.type === "video").length;
    if (videos > 4096) reject("at most 4096 video objects per slide are supported.");
    if (videos && slide.eventTimeline !== undefined) reject("video cannot share a slide with authored event timing.");
    for (const element of slide.elements) check(element, slide.id);
  }
  return resources;
}
