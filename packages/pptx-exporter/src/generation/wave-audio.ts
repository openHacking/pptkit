import { inspectWavePcm } from "@pptkit/media-inspection";
import type { NormalizedAsset, NormalizedElement, NormalizedPresentation } from "@pptkit/core";
import { decodeAssetBytes } from "../assets/load-bytes.js";

export interface WaveAudioResource { readonly data: Uint8Array; readonly durationMs: number }

/** Validate the complete bounded profile before invoking any external asset resolver. */
export function planWaveAudio(document: NormalizedPresentation, mode: "wave-pcm-v1" | undefined, maxBytes: number | undefined, signal?: AbortSignal, external?: Map<string, NormalizedAsset>, resolved?: ReadonlyMap<string, Uint8Array>): ReadonlyMap<string, WaveAudioResource> {
  const assets = new Map(document.assets.map(asset => [asset.id, asset]));
  const resources = new Map<string, WaveAudioResource>();
  const reject = (reason: string): never => { throw new TypeError(`PPTX audio export: ${reason}`); };
  const check = (element: NormalizedElement, slideId?: string, nested = false) => {
    if (element.type === "group") { for (const child of element.children) check(child, slideId, true); return; }
    if (element.type !== "audio") return;
    if (mode !== "wave-pcm-v1") return reject("requires audioMode wave-pcm-v1; audio cannot be silently omitted.");
    if (element.playbackStart !== undefined && element.playbackStart !== "click-sequence" && element.playbackStart !== "object-click" && element.playbackStart !== "automatic") return reject("playbackStart must be click-sequence, object-click, or automatic.");
    if (element.playbackVolume !== undefined && element.playbackVolume !== "low" && element.playbackVolume !== "medium" && element.playbackVolume !== "high" && element.playbackVolume !== "mute") return reject("playbackVolume must be low, medium, high, or mute.");
    if (element.loop !== undefined && typeof element.loop !== "boolean") return reject("loop must be a boolean.");
    if (element.rewindAfterPlayback !== undefined && typeof element.rewindAfterPlayback !== "boolean") return reject("rewindAfterPlayback must be a boolean.");
    for (const key of ["trimStartMs", "trimEndMs"] as const) if (element[key] !== undefined && (!Number.isSafeInteger(element[key]) || element[key]! < 0 || element[key]! > 2147483647)) return reject(`${key} must be a nonnegative integer no greater than 2147483647 milliseconds.`);
    if (!slideId || nested || element.hidden || element.action !== undefined || element.placeholderKey !== undefined || element.opacity !== 1 || !element.box.width || !element.box.height) return reject("only visible, nontransparent slide-local top-level audio with nonzero geometry and no action or placeholder binding is supported.");
    const asset = assets.get(element.audioAssetId);
    if (!asset || asset.kind !== "audio") return reject("audio requires an audio asset.");
    const mime = asset.mimeType?.trim().toLowerCase();
    if (mime !== "audio/wav" && mime !== "audio/x-wav") return reject("a declared audio/wav or audio/x-wav MIME type is required.");
    const validateDuration = (durationMs: number) => { if ((element.trimStartMs ?? 0) + (element.trimEndMs ?? 0) >= durationMs) return reject("audio trim must leave at least one millisecond of the rounded source duration."); };
    const existing = resources.get(asset.id);
    if (existing) { validateDuration(existing.durationMs); return; }
    if (asset.source.type !== "bytes" && !resolved?.has(asset.id)) {
      if (!external) return reject("audio requires a byte-backed audio asset; paths and links are not loaded.");
      external.set(asset.id, asset); return;
    }
    const length = asset.source.type === "bytes" ? asset.source.value.length / 4 * 3 - (asset.source.value.endsWith("==") ? 2 : asset.source.value.endsWith("=") ? 1 : 0) : resolved!.get(asset.id)!.byteLength;
    if (length > Math.min(maxBytes ?? 64 * 1024 * 1024, 64 * 1024 * 1024)) throw new RangeError("Audio exceeds the WAV inspection or maxAssetBytes limit.");
    signal?.throwIfAborted();
    const data = asset.source.type === "bytes" ? decodeAssetBytes(asset.source.value) : resolved!.get(asset.id)!, info = inspectWavePcm(data, { ...(signal ? { signal } : {}), maxBytes: Math.min(maxBytes ?? 64 * 1024 * 1024, 64 * 1024 * 1024) });
    const durationMs = Math.round(info.durationSeconds * 1000);
    if (!Number.isSafeInteger(durationMs) || durationMs <= 0 || durationMs > 2147483647) return reject("audio duration must round to 1–2147483647 milliseconds.");
    validateDuration(durationMs);
    resources.set(asset.id, { data, durationMs });
  };
  for (const layout of document.layouts) for (const element of layout.elements) check(element);
  for (const master of document.masters ?? []) for (const element of master.elements ?? []) check(element);
  for (const slide of document.slides) {
    const count = slide.elements.filter(element => element.type === "audio").length;
    if (count > 4096) reject("at most 4096 audio objects per slide are supported.");
    if (count && slide.eventTimeline !== undefined) reject("audio and authored event timelines cannot be combined.");
    for (const element of slide.elements) check(element, slide.id);
  }
  return resources;
}
