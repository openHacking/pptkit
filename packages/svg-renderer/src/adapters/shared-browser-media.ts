import { resolveAndMountBrowserMedia, type BrowserMediaController, type BrowserMediaResolveOptions } from './browser-media.js';

export interface SharedBrowserMediaPoolOptions {
  /** Bound the resolved Blob and provide the same budget to the caller-owned loader. */
  readonly maxBytes: number;
  /** The caller owns authorization, transport, MIME validation and streaming limits. */
  readonly resolve: (key: string, context: Readonly<{ signal: AbortSignal; maxBytes: number }>) => Blob | Promise<Blob>;
}
export type SharedBrowserMediaMountOptions = Omit<BrowserMediaResolveOptions, 'resolve' | 'maxBytes'>;
export interface SharedBrowserMediaPool {
  /** Instances with the same key share bytes while at least one mount is pending or active. */
  mount(container: HTMLElement, key: string, options: SharedBrowserMediaMountOptions): Promise<BrowserMediaController>;
}

interface Entry { readonly abort: AbortController; readonly result: Promise<Blob>; refs: number }

/** Create an explicit, window-bound pool; it does not fetch or retain unused resources. */
export function createSharedBrowserMediaPool(view: Window & typeof globalThis, options: SharedBrowserMediaPoolOptions): SharedBrowserMediaPool {
  if (!view || typeof view.AbortController !== 'function' || typeof view.Blob !== 'function') throw new TypeError('A browser window is required.');
  if (!options || !Number.isSafeInteger(options.maxBytes) || options.maxBytes < 1 || typeof options.resolve !== 'function') throw new TypeError('A positive maxBytes and resolver are required.');
  const { maxBytes, resolve } = options, entries = new Map<string, Entry>();
  return Object.freeze({
    async mount(container: HTMLElement, key: string, mountOptions: SharedBrowserMediaMountOptions): Promise<BrowserMediaController> {
      if (!container || container.ownerDocument?.defaultView !== view) throw new TypeError('Container must belong to the pool window.');
      if (typeof key !== 'string' || !key) throw new TypeError('A nonempty media cache key is required.');
      let entry: Entry | undefined, released = false;
      const release = () => {
        if (released || !entry) return;
        released = true;
        if (--entry.refs === 0) { if (entries.get(key) === entry) entries.delete(key); entry.abort.abort(); }
      };
      try {
        const controller = await resolveAndMountBrowserMedia(container, {
          ...mountOptions, maxBytes,
          resolve: () => {
            entry = entries.get(key);
            if (!entry) {
              const abort = new view.AbortController();
              let created: Entry;
              const result = Promise.resolve().then(() => resolve(key, Object.freeze({ signal: abort.signal, maxBytes }))).then(source => {
                if (!(source instanceof view.Blob) || !source.size) throw new TypeError('The media resolver must return a nonempty Blob.');
                if (source.size > maxBytes) throw new RangeError('Resolved media exceeds maxBytes.');
                return source;
              }).catch(error => { if (entries.get(key) === created) entries.delete(key); throw error; });
              created = { abort, result, refs: 0 }; entry = created; entries.set(key, created);
            }
            entry.refs++;
            return entry.result;
          },
        });
        if (controller.state.status === 'disposed') release();
        else controller.subscribe(state => { if (state.status === 'disposed') release(); });
        return controller;
      } catch (error) { release(); throw error; }
    },
  });
}
