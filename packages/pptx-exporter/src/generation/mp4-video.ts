import { inspectIsoBmff } from "@pptkit/media-inspection";
import type { NormalizedElement, NormalizedPresentation } from "@pptkit/core";
import { decodeAssetBytes } from "../assets/load-bytes.js";

export interface Mp4VideoResource { readonly data: Uint8Array; readonly durationMs: number }

/** Read the v0 movie-header clock only; this does not validate tracks or codecs. */
function movieDuration(data: Uint8Array, maxBytes: number, signal?: AbortSignal): number {
  const framing = inspectIsoBmff(data, { maxBytes, ...(signal ? { signal } : {}) });
  const movies = framing.boxes.filter(box => box.type === "moov"), media = framing.boxes.filter(box => box.type === "mdat");
  if (movies.length !== 1 || media.length !== 1 || media[0]!.size <= media[0]!.headerSize || framing.boxes.some(box => box.type === "moof")) throw new TypeError("MP4 requires one movie box and one nonempty media-data box without fragments.");
  const movie = movies[0]!, view = new DataView(data.buffer, data.byteOffset, data.byteLength), end = movie.offset + movie.size;
  let at = movie.offset + movie.headerSize, count = 0, durationMs: number | undefined;
  while (at < end) {
    signal?.throwIfAborted();
    if (++count > 4096) throw new RangeError("MP4 movie box child limit exceeded.");
    if (end - at < 8) throw new TypeError("MP4 movie box has a truncated child header.");
    const shortSize = view.getUint32(at), type = String.fromCharCode(...data.subarray(at + 4, at + 8));
    let size = shortSize, header = 8;
    if (shortSize === 0) size = end - at;
    else if (shortSize === 1) {
      if (end - at < 16 || view.getUint32(at + 8) !== 0) throw new TypeError("MP4 movie box has an invalid extended child size.");
      header = 16; size = view.getUint32(at + 12);
    }
    if (size < header || size > end - at) throw new TypeError("MP4 movie box has an invalid child size.");
    if (type === "mvhd") {
      if (durationMs !== undefined || size - header < 100) throw new TypeError("MP4 needs one complete version-zero movie header.");
      const start = at + header;
      if (view.getUint32(start) !== 0) throw new TypeError("Only version-zero MP4 movie headers with zero flags are supported.");
      const scale = view.getUint32(start + 12), duration = view.getUint32(start + 16);
      if (!scale || !duration || duration === 0xffffffff) throw new TypeError("MP4 movie duration must be known and nonzero.");
      durationMs = Math.round(duration * 1000 / scale);
      if (!Number.isSafeInteger(durationMs) || durationMs < 1 || durationMs > 2147483647) throw new TypeError("MP4 movie duration must round to 1–2147483647 milliseconds.");
    }
    at += size;
  }
  if (durationMs === undefined) throw new TypeError("MP4 movie header is missing.");
  return durationMs;
}

/** Experimental profile: one visible top-level video per slide, no authored timing. */
export function planMp4Video(document: NormalizedPresentation, mode: "iso-bmff-v1" | undefined, maxBytes: number | undefined, signal?: AbortSignal): ReadonlyMap<string, Mp4VideoResource> {
  const assets = new Map(document.assets.map(asset => [asset.id, asset])), resources = new Map<string, Mp4VideoResource>();
  const reject = (reason: string): never => { throw new TypeError(`PPTX video export: ${reason}`); };
  const check = (element: NormalizedElement, slideId?: string, nested = false) => {
    if (element.type === "group") { for (const child of element.children) check(child, slideId, true); return; }
    if (element.type !== "video") return;
    if (mode !== "iso-bmff-v1") return reject("requires videoMode iso-bmff-v1; video cannot be silently omitted.");
    if (!slideId || nested || element.hidden || element.action !== undefined || element.placeholderKey !== undefined || element.opacity !== 1 || !element.box.width || !element.box.height) return reject("only visible, nontransparent slide-local top-level video with nonzero geometry and no action or placeholder binding is supported.");
    const asset = assets.get(element.videoAssetId);
    if (!asset || asset.kind !== "video" || asset.source.type !== "bytes") return reject("video requires a byte-backed asset; paths and links are not loaded.");
    if (asset.mimeType?.trim().toLowerCase() !== "video/mp4") return reject("a declared video/mp4 MIME type is required.");
    if (resources.has(asset.id)) return;
    const ceiling = Math.min(maxBytes ?? 64 * 1024 * 1024, 64 * 1024 * 1024);
    const length = asset.source.value.length / 4 * 3 - (asset.source.value.endsWith("==") ? 2 : asset.source.value.endsWith("=") ? 1 : 0);
    if (length > ceiling) throw new RangeError("Video exceeds the MP4 inspection or maxAssetBytes limit.");
    signal?.throwIfAborted();
    const data = decodeAssetBytes(asset.source.value), durationMs = movieDuration(data, ceiling, signal);
    resources.set(asset.id, { data, durationMs });
  };
  for (const layout of document.layouts) for (const element of layout.elements) check(element);
  for (const master of document.masters ?? []) for (const element of master.elements ?? []) check(element);
  for (const slide of document.slides) {
    const videos = slide.elements.filter(element => element.type === "video").length;
    if (videos > 1 || videos && (slide.eventTimeline !== undefined || slide.elements.some(element => element.type === "audio"))) reject("at most one video per slide is supported, without audio or authored event timing on that slide.");
    for (const element of slide.elements) check(element, slide.id);
  }
  return resources;
}
