import type { NormalizedAsset } from '@pptkit/core';
import type { LayoutElement, LayoutResult } from '@pptkit/layout';
import { resolveAndMountBrowserMedia, type BrowserMediaController, type BrowserMediaResolveState } from './browser-media.js';

export type BrowserMediaAssetRequest = Readonly<Omit<NormalizedAsset, 'source' | 'accessibility'>> & {
  readonly source: Readonly<NormalizedAsset['source']>;
  readonly accessibility: Readonly<NormalizedAsset['accessibility']>;
};
export interface BrowserLayoutMediaOptions {
  readonly slideId: string;
  readonly elementId: string;
  readonly maxBytes: number;
  /** Optionally supplies a codec-qualified type; its base MIME must match a declared asset type. */
  readonly contentType?: string;
  readonly signal?: AbortSignal;
  readonly controls?: boolean;
  readonly onResolveState?: (state: BrowserMediaResolveState) => void;
  /** Explicitly owns authorization, streaming limits and loading for URL/path assets. */
  readonly resolveExternal?: (asset: BrowserMediaAssetRequest, context: Readonly<{ signal: AbortSignal; maxBytes: number }>) => Blob | Promise<Blob>;
}
export interface BrowserLayoutMedia {
  readonly slideId: string;
  readonly elementId: string;
  readonly assetId: string;
  readonly kind: 'audio' | 'video';
  /** Native start intent is reported but never executed by this adapter. */
  readonly playbackStart: 'click-sequence' | 'object-click' | 'automatic';
  readonly controller: BrowserMediaController;
}

function assetBytes(source: string, maxBytes: number, view: Window & typeof globalThis): BlobPart {
  if (!source || source.length % 4 || source.length > Math.ceil(maxBytes / 3) * 4) throw new RangeError('Byte asset is empty, malformed or exceeds maxBytes.');
  if (!/^[A-Za-z0-9+/]*={0,2}$/u.test(source)) throw new TypeError('Byte asset must use canonical Base64.');
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  if ((source.endsWith('==') && (alphabet.indexOf(source[source.length - 3]!) & 15) !== 0) || (source.endsWith('=') && !source.endsWith('==') && (alphabet.indexOf(source[source.length - 2]!) & 3) !== 0)) throw new TypeError('Byte asset must use canonical Base64.');
  const size = source.length / 4 * 3 - (source.endsWith('==') ? 2 : source.endsWith('=') ? 1 : 0);
  if (!size || size > maxBytes) throw new RangeError('Byte asset is empty or exceeds maxBytes.');
  let binary: string; try { binary = view.atob(source); } catch { throw new TypeError('Byte asset must use canonical Base64.'); }
  const bytes = new Uint8Array(binary.length); for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/** Mount one explicit slide-local Layout media object; caller owns placement and playback. */
export async function mountLayoutMedia(container: HTMLElement, layout: LayoutResult, options: BrowserLayoutMediaOptions): Promise<BrowserLayoutMedia> {
  if (!options || typeof options.slideId !== 'string' || !options.slideId || typeof options.elementId !== 'string' || !options.elementId) throw new TypeError('slideId and elementId are required.');
  const { slideId, elementId } = options;
  if (!Number.isSafeInteger(options.maxBytes) || options.maxBytes < 1) throw new TypeError('maxBytes must be a positive safe integer.');
  if (options.resolveExternal !== undefined && typeof options.resolveExternal !== 'function') throw new TypeError('resolveExternal must be a function.');
  if (!layout || layout.status !== 'resolved' || !Array.isArray(layout.slides) || !Array.isArray(layout.assets)) throw new TypeError('A resolved Layout result is required.');
  if (layout.slides.length > 100000 || layout.assets.length > 100000) throw new RangeError('Layout media search exceeds its work limit.');
  const slides = layout.slides.filter(slide => slide.id === slideId);
  if (slides.length !== 1) throw new TypeError('slideId must identify exactly one slide.');
  const slide = slides[0]!;
  if (!Array.isArray(slide.elements) || slide.elements.length > 100000) throw new RangeError('Slide media search exceeds its work limit.');
  const stack = slide.elements.map(element => ({ element, hidden: false })), seen = new Set<LayoutElement>();
  let target: LayoutElement | undefined, hidden = false, work = 0;
  while (stack.length) {
    if (++work > 100000) throw new RangeError('Slide media search exceeds its work limit.');
    const item = stack.pop()!, element = item.element;
    if (!element || typeof element !== 'object' || seen.has(element)) throw new TypeError('Slide element structure is invalid or cyclic.');
    seen.add(element);
    const invisible = item.hidden || element.hidden;
    if (element.id === elementId) { if (target) throw new TypeError('elementId is ambiguous on this slide.'); target = element; hidden = invisible; }
    if (element.type === 'group') { if (!Array.isArray(element.children)) throw new TypeError('Group children are invalid.'); if (work + stack.length + element.children.length > 100000) throw new RangeError('Slide media search exceeds its work limit.'); for (const child of element.children) stack.push({ element: child, hidden: invisible }); }
  }
  if (!target || (target.type !== 'audio' && target.type !== 'video')) throw new TypeError('elementId must identify an audio or video object on the slide.');
  if (hidden) throw new TypeError('Hidden media or media in a hidden group cannot be mounted.');
  const kind = target.type, assetId = kind === 'audio' ? target.audioAssetId : target.videoAssetId;
  const assets = layout.assets.filter(asset => asset.id === assetId);
  if (assets.length !== 1 || assets[0]!.kind !== kind) throw new TypeError('Media asset is missing, ambiguous or has the wrong kind.');
  const asset = assets[0]!;
  if (!asset.source || !['bytes', 'url', 'path'].includes(asset.source.type) || typeof asset.source.value !== 'string' || !asset.source.value || (asset.source.type === 'bytes' && asset.source.encoding !== 'base64')) throw new TypeError('Media asset source is invalid.');
  const baseMime = (value: string) => value.split(';', 1)[0]!.trim().toLowerCase();
  const contentType = options.contentType ?? asset.mimeType;
  if (typeof contentType !== 'string' || !new RegExp(`^${kind}/[a-z0-9!#$%&'*+.^_\x60|~-]+$`, 'u').test(baseMime(contentType)) || (asset.mimeType !== undefined && baseMime(asset.mimeType) !== baseMime(contentType))) throw new TypeError('An appropriate matching media content type is required.');
  const playbackStart = target.playbackStart === undefined ? 'click-sequence' : target.playbackStart;
  if (!['click-sequence', 'object-click', 'automatic'].includes(playbackStart)) throw new TypeError('Invalid media playbackStart.');
  const level = target.playbackVolume === undefined ? 'high' : target.playbackVolume;
  if (!['low', 'medium', 'high', ...(kind === 'audio' ? ['mute'] : [])].includes(level)) throw new TypeError('Invalid media playbackVolume.');
  const volume = level === 'low' ? 0.2 : level === 'medium' ? 0.5 : 0.8;
  if (target.type === 'video' && target.muted !== undefined && typeof target.muted !== 'boolean') throw new TypeError('Invalid video mute setting.');
  const muted = target.type === 'audio' ? level === 'mute' : target.muted ?? false;
  const request: BrowserMediaAssetRequest = Object.freeze({ ...asset, source: Object.freeze({ ...asset.source }), accessibility: Object.freeze({ ...asset.accessibility }) });
  const { maxBytes, resolveExternal, signal, controls, onResolveState } = options;
  if (request.source.type !== 'bytes' && !resolveExternal) throw new TypeError('External media requires an explicit resolver.');
  const owner = container?.ownerDocument?.defaultView;
  if (!owner) throw new TypeError('A live browser container is required.');
  const controller = await resolveAndMountBrowserMedia(container, {
    kind, contentType, maxBytes, volume, muted,
    ...(target.loop === undefined ? {} : { loop: target.loop }),
    ...(target.rewindAfterPlayback === undefined ? {} : { rewindAfterPlayback: target.rewindAfterPlayback }),
    ...(target.trimStartMs === undefined ? {} : { trimStartMs: target.trimStartMs }),
    ...(target.trimEndMs === undefined ? {} : { trimEndMs: target.trimEndMs }),
    ...(controls === undefined ? {} : { controls }), ...(signal === undefined ? {} : { signal }), ...(onResolveState === undefined ? {} : { onResolveState }),
    resolve: ({ signal: loadSignal }) => {
      if (request.source.type === 'bytes') return new owner.Blob([assetBytes(request.source.value, maxBytes, owner)], { type: baseMime(contentType) });
      return resolveExternal!(request, Object.freeze({ signal: loadSignal, maxBytes }));
    },
  });
  return Object.freeze({ slideId, elementId, assetId, kind, playbackStart, controller });
}
