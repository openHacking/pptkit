import { inspectWavePcm } from "@pptkit/media-inspection";
import type { NormalizedElement, NormalizedPresentation } from "@pptkit/core";
import { decodeAssetBytes } from "../assets/load-bytes.js";

export interface WaveAudioResource { readonly data: Uint8Array; readonly durationMs: number }

/** Validate the complete bounded profile before invoking any external asset resolver. */
export function planWaveAudio(document: NormalizedPresentation, mode: "wave-pcm-v1" | undefined, maxBytes: number | undefined, signal?: AbortSignal): ReadonlyMap<string, WaveAudioResource> {
  const assets = new Map(document.assets.map(asset => [asset.id, asset]));
  const resources = new Map<string, WaveAudioResource>();
  const reject = (reason: string): never => { throw new TypeError(`PPTX audio export: ${reason}`); };
  const check = (element: NormalizedElement, slideId?: string, nested = false) => {
    if (element.type === "group") { for (const child of element.children) check(child, slideId, true); return; }
    if (element.type !== "audio") return;
    if (mode !== "wave-pcm-v1") return reject("requires audioMode wave-pcm-v1; audio cannot be silently omitted.");
    if (!slideId || nested || element.hidden || element.action !== undefined || element.placeholderKey !== undefined || element.opacity !== 1 || !element.box.width || !element.box.height) return reject("only visible, nontransparent slide-local top-level audio with nonzero geometry and no action or placeholder binding is supported.");
    const asset = assets.get(element.audioAssetId);
    if (!asset || asset.kind !== "audio" || asset.source.type !== "bytes") return reject("audio requires a byte-backed audio asset; paths and links are not loaded.");
    const mime = asset.mimeType?.trim().toLowerCase();
    if (mime !== "audio/wav" && mime !== "audio/x-wav") return reject("a declared audio/wav or audio/x-wav MIME type is required.");
    if (resources.has(asset.id)) return;
    const length = asset.source.value.length / 4 * 3 - (asset.source.value.endsWith("==") ? 2 : asset.source.value.endsWith("=") ? 1 : 0);
    if (length > Math.min(maxBytes ?? 64 * 1024 * 1024, 64 * 1024 * 1024)) throw new RangeError("Audio exceeds the WAV inspection or maxAssetBytes limit.");
    signal?.throwIfAborted();
    const data = decodeAssetBytes(asset.source.value), info = inspectWavePcm(data, { ...(signal ? { signal } : {}), maxBytes: Math.min(maxBytes ?? 64 * 1024 * 1024, 64 * 1024 * 1024) });
    const durationMs = Math.round(info.durationSeconds * 1000);
    if (!Number.isSafeInteger(durationMs) || durationMs <= 0 || durationMs > 2147483647) return reject("audio duration must round to 1–2147483647 milliseconds.");
    resources.set(asset.id, { data, durationMs });
  };
  for (const layout of document.layouts) for (const element of layout.elements) check(element);
  for (const master of document.masters ?? []) for (const element of master.elements ?? []) check(element);
  for (const slide of document.slides) {
    const count = slide.elements.filter(element => element.type === "audio").length;
    if (count > 1) reject("at most one audio object per slide is currently supported.");
    if (count && slide.eventTimeline !== undefined) reject("audio and authored event timelines cannot be combined.");
    for (const element of slide.elements) check(element, slide.id);
  }
  return resources;
}
