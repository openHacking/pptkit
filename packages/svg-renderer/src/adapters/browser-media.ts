export type BrowserMediaCapability = 'probably' | 'maybe' | 'unsupported';
export type BrowserMediaFailure = 'permission' | 'unsupported' | 'network' | 'decode' | 'invalid-duration' | 'playback';
export interface BrowserMediaState {
  readonly status: 'loading' | 'ready' | 'playing' | 'paused' | 'ended' | 'error' | 'disposed';
  readonly capability: BrowserMediaCapability;
  /** Source-relative milliseconds, not the trimmed interval's local clock. */
  readonly currentTimeMs: number;
  readonly durationMs?: number;
  readonly failure?: BrowserMediaFailure;
}
export interface BrowserMediaOptions {
  readonly kind: 'audio' | 'video';
  /** Caller-owned, already acquired bytes; no URLs are fetched by the adapter. */
  readonly source: Blob;
  /** MIME type, optionally codec-qualified, used as a capability hint rather than a guarantee. */
  readonly contentType: string;
  readonly trimStartMs?: number;
  readonly trimEndMs?: number;
  readonly volume?: number;
  readonly muted?: boolean;
  readonly loop?: boolean;
  readonly rewindAfterPlayback?: boolean;
  readonly controls?: boolean;
  /** Aborting disposes a mounted instance; the caller owns the signal. */
  readonly signal?: AbortSignal;
}
export type BrowserMediaResolveState = Readonly<{ status: 'loading' | 'mounted' | 'cancelled' | 'failed' }>;
export interface BrowserMediaResolveOptions extends Omit<BrowserMediaOptions, 'source'> {
  /** Explicit callback owns authorization, redirects, transport and actual stream budgets. */
  readonly resolve: (context: Readonly<{ signal: AbortSignal }>) => Blob | Promise<Blob>;
  /** Maximum resolved Blob size in bytes. */
  readonly maxBytes: number;
  /** Optional observation of resolution; callback errors do not alter the operation. */
  readonly onResolveState?: (state: BrowserMediaResolveState) => void;
}
export interface BrowserMediaController {
  readonly element: HTMLMediaElement;
  readonly state: BrowserMediaState;
  subscribe(listener: (state: BrowserMediaState) => void): () => void;
  /** Must be called by the consumer, normally from a user gesture. */
  play(): Promise<void>;
  pause(): void;
  /** Seek in source-relative milliseconds within the retained interval. */
  seek(sourceTimeMs: number): void;
  dispose(): void;
}

function checkConfiguration(container: HTMLElement, options: Omit<BrowserMediaOptions, 'source'>) {
  const doc = container?.ownerDocument, view = doc?.defaultView;
  if (!doc || !view || !container.isConnected || !view.URL?.createObjectURL || !view.URL.revokeObjectURL) throw new TypeError('A connected browser container with object URL support is required.');
  if (!options || (options.kind !== 'audio' && options.kind !== 'video') || typeof options.contentType !== 'string' || !options.contentType.trim()) throw new TypeError('A media kind and contentType are required.');
  for (const key of ['trimStartMs', 'trimEndMs'] as const) if (options[key] !== undefined && (!Number.isSafeInteger(options[key]) || options[key]! < 0)) throw new TypeError(`${key} must be a nonnegative integer.`);
  for (const key of ['muted', 'loop', 'rewindAfterPlayback', 'controls'] as const) if (options[key] !== undefined && typeof options[key] !== 'boolean') throw new TypeError(`${key} must be a boolean.`);
  if (options.volume !== undefined && (typeof options.volume !== 'number' || !Number.isFinite(options.volume) || options.volume < 0 || options.volume > 1)) throw new TypeError('volume must be between 0 and 1.');
  if (options.signal !== undefined && !(options.signal instanceof view.AbortSignal)) throw new TypeError('signal must be an AbortSignal from the container’s window.');
  return { doc, view };
}
function abortReason(signal: AbortSignal): unknown {
  return signal.reason ?? new DOMException('Media operation aborted.', 'AbortError');
}

/** Mount one explicitly supplied blob without starting playback or fetching a URL. */
export function mountBrowserMedia(container: HTMLElement, options: BrowserMediaOptions): BrowserMediaController {
  const { doc, view } = checkConfiguration(container, options);
  if (!(options.source instanceof view.Blob) || !options.source.size) throw new TypeError('A nonempty Blob is required.');
  const signal = options.signal;
  if (signal?.aborted) throw abortReason(signal);
  const element = doc.createElement(options.kind), start = options.trimStartMs ?? 0, trimEnd = options.trimEndMs ?? 0, loop = options.loop ?? false, rewind = options.rewindAfterPlayback ?? false;
  const verdict = element.canPlayType(options.contentType), capability: BrowserMediaCapability = verdict === 'probably' || verdict === 'maybe' ? verdict : 'unsupported';
  const listeners = new Set<(state: BrowserMediaState) => void>();
  let disposed = false, epoch = 0, duration: number | undefined, boundary = false, wantsPlayback = false;
  let current: BrowserMediaState = Object.freeze({ status: 'loading', capability, currentTimeMs: 0 });
  const publish = (status: BrowserMediaState['status'], failure?: BrowserMediaFailure) => {
    if (disposed && status !== 'disposed') return;
    current = Object.freeze({ status, capability, currentTimeMs: Number.isFinite(element.currentTime) ? element.currentTime * 1000 : 0, ...(duration === undefined ? {} : { durationMs: duration }), ...(failure ? { failure } : {}) });
    for (const listener of [...listeners]) { try { listener(current); } catch { /* A subscriber cannot interrupt media cleanup or other observers. */ } }
  };
  const ensure = () => { if (disposed) throw new TypeError('Media controller is disposed.'); };
  const end = () => duration === undefined ? undefined : duration - trimEnd;
  const classify = (error: unknown): BrowserMediaFailure => {
    const name = typeof error === 'object' && error !== null && 'name' in error ? String(error.name) : '';
    return name === 'NotAllowedError' ? 'permission' : name === 'NotSupportedError' ? 'unsupported' : 'playback';
  };
  const finish = () => {
    if (disposed || boundary || duration === undefined || current.status === 'ended' || current.status === 'error') return;
    boundary = true;
    if (loop) {
      element.currentTime = start / 1000;
      if (element.paused && wantsPlayback) { const token = epoch; void element.play().catch(error => { if (!disposed && token === epoch) { wantsPlayback = false; publish('error', classify(error)); } }); }
      else publish(element.paused ? 'paused' : 'playing');
    } else {
      wantsPlayback = false; element.pause();
      if (rewind) element.currentTime = start / 1000;
      publish('ended');
    }
    boundary = false;
  };
  const handlers: [keyof HTMLMediaElementEventMap, EventListener][] = [
    ['loadedmetadata', () => {
      if (current.status === 'error' && current.failure !== 'permission' && current.failure !== 'playback') return;
      if (!Number.isFinite(element.duration) || element.duration <= 0 || element.duration * 1000 > Number.MAX_SAFE_INTEGER || start + trimEnd >= element.duration * 1000) { ++epoch; wantsPlayback = false; element.pause(); publish('error', 'invalid-duration'); return; }
      duration = element.duration * 1000;
      if (element.currentTime < start / 1000 || element.currentTime > (duration - trimEnd) / 1000) element.currentTime = start / 1000;
      if (current.status === 'error') publish('error', current.failure); else publish('ready');
    }],
    ['playing', () => { if (!element.paused) { wantsPlayback = true; if (current.status !== 'error') publish('playing'); } }],
    ['pause', () => { if (!element.paused) return; if (!boundary && !element.ended) wantsPlayback = false; if (duration !== undefined && !boundary && current.status !== 'error' && current.status !== 'ended') publish('paused'); }],
    ['seeking', () => { if (duration !== undefined && element.currentTime * 1000 < start) element.currentTime = start / 1000; }],
    ['timeupdate', () => { if (end() !== undefined && element.currentTime * 1000 >= end()!) finish(); else if (current.status !== 'error' && current.status !== 'ended') publish(current.status); }],
    ['ended', finish],
    ['error', () => { ++epoch; wantsPlayback = false; const code = element.error?.code; publish('error', code === 2 ? 'network' : code === 3 ? 'decode' : code === 4 ? 'unsupported' : 'playback'); }],
  ];
  element.preload = 'metadata'; element.controls = options.controls ?? false; element.muted = options.muted ?? false; element.volume = options.volume ?? 0.8;
  for (const [name, handler] of handlers) element.addEventListener(name, handler);
  const url = view.URL.createObjectURL(options.source);
  try { element.src = url; container.appendChild(element); }
  catch (error) { for (const [name, handler] of handlers) element.removeEventListener(name, handler); element.removeAttribute('src'); view.URL.revokeObjectURL(url); throw error; }
  const controller: BrowserMediaController = {
    element,
    get state() { return current; },
    subscribe(listener) { ensure(); if (typeof listener !== 'function') throw new TypeError('listener must be a function.'); listeners.add(listener); try { listener(current); } catch { /* Observer errors are isolated. */ } return () => { listeners.delete(listener); }; },
    async play() {
      ensure();
      if (current.status === 'error' && current.failure !== 'permission' && current.failure !== 'playback') throw new TypeError('Media is in an error state.');
      if (current.status === 'error') { publish(duration === undefined ? 'loading' : 'paused'); ensure(); }
      const token = ++epoch; wantsPlayback = true;
      if (duration !== undefined && (element.currentTime * 1000 >= end()! || element.currentTime * 1000 < start)) element.currentTime = start / 1000;
      try { await element.play(); }
      catch (error) { if (!disposed && token === epoch) { wantsPlayback = false; publish('error', classify(error)); } throw error; }
    },
    pause() { ensure(); ++epoch; wantsPlayback = false; element.pause(); },
    seek(ms) { ensure(); if (duration === undefined || !Number.isFinite(ms) || ms < start || ms > end()!) throw new RangeError('seek must be within the retained source interval after metadata loads.'); element.currentTime = ms / 1000; if (current.status === 'ended') publish('paused'); },
    dispose() {
      if (disposed) return;
      disposed = true; signal?.removeEventListener('abort', abort); ++epoch; wantsPlayback = false; element.pause(); for (const [name, handler] of handlers) element.removeEventListener(name, handler);
      element.removeAttribute('src'); element.load(); element.remove(); view.URL.revokeObjectURL(url); publish('disposed'); listeners.clear();
    },
  };
  const abort = () => controller.dispose();
  signal?.addEventListener('abort', abort, { once: true });
  if (signal?.aborted) { controller.dispose(); throw abortReason(signal); }
  return controller;
}

/** Resolve explicitly supplied bytes with bounded result size and cancellation. */
export async function resolveAndMountBrowserMedia(container: HTMLElement, options: BrowserMediaResolveOptions): Promise<BrowserMediaController> {
  const { view } = checkConfiguration(container, options);
  if (typeof options.resolve !== 'function') throw new TypeError('resolve must be a function.');
  if (!Number.isSafeInteger(options.maxBytes) || options.maxBytes < 1) throw new TypeError('maxBytes must be a positive safe integer.');
  if (options.onResolveState !== undefined && typeof options.onResolveState !== 'function') throw new TypeError('onResolveState must be a function.');
  const { kind, contentType, trimStartMs, trimEndMs, volume, muted, loop, rewindAfterPlayback, controls, signal, resolve, maxBytes, onResolveState } = options;
  const internal = signal ?? new view.AbortController().signal;
  const notify = (status: BrowserMediaResolveState['status']) => { try { onResolveState?.(Object.freeze({ status })); } catch { /* Observers do not control resource ownership. */ } };
  try {
    if (internal.aborted) throw abortReason(internal);
    notify('loading');
    if (internal.aborted) throw abortReason(internal);
    const context = Object.freeze({ signal: internal });
    let onAbort: (() => void) | undefined;
    const aborted = new Promise<never>((_, reject) => {
      onAbort = () => reject(abortReason(internal));
      internal.addEventListener('abort', onAbort, { once: true });
      if (internal.aborted) onAbort();
    });
    let source: Blob;
    try { source = await Promise.race([Promise.resolve().then(() => { if (internal.aborted) throw abortReason(internal); return resolve(context); }), aborted]); }
    finally { if (onAbort) internal.removeEventListener('abort', onAbort); }
    if (internal.aborted) throw abortReason(internal);
    if (!(source instanceof view.Blob) || !source.size) throw new TypeError('The media resolver must return a nonempty Blob.');
    if (source.size > maxBytes) throw new RangeError('Resolved media exceeds maxBytes.');
    const controller = mountBrowserMedia(container, { kind, contentType, source, ...(trimStartMs === undefined ? {} : { trimStartMs }), ...(trimEndMs === undefined ? {} : { trimEndMs }), ...(volume === undefined ? {} : { volume }), ...(muted === undefined ? {} : { muted }), ...(loop === undefined ? {} : { loop }), ...(rewindAfterPlayback === undefined ? {} : { rewindAfterPlayback }), ...(controls === undefined ? {} : { controls }), ...(signal === undefined ? {} : { signal }) });
    notify('mounted');
    return controller;
  } catch (error) {
    notify(internal.aborted ? 'cancelled' : 'failed');
    throw error;
  }
}
