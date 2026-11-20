import type { PptxMediaResolver } from "../types/export.js";
import { withAbort } from "../assets/limits.js";

export interface HttpMediaResolverOptions {
  /** Called with each normalized URL before any request to it, including redirects. */
  authorizeUrl: (url: string) => boolean | Promise<boolean>;
  /** Maximum redirect responses to follow; defaults to zero, at most ten. */
  maxRedirects?: number;
}

/** Optional HTTP adapter; the caller owns destination authorization and network policy. */
export function createHttpMediaResolver(options: HttpMediaResolverOptions): PptxMediaResolver {
  if (!options || typeof options.authorizeUrl !== "function") throw new TypeError("authorizeUrl must be a function.");
  const { authorizeUrl, maxRedirects = 0 } = options;
  if (!Number.isSafeInteger(maxRedirects) || maxRedirects < 0 || maxRedirects > 10) throw new RangeError("maxRedirects must be an integer from 0 to 10.");
  return async (asset, { signal, maxBytes }) => {
    if (asset.source.type !== "url") throw new TypeError("HTTP media resolver requires a URL source.");
    if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0) throw new RangeError("maxBytes must be a positive safe integer.");
    let current = asset.source.value;
    for (let redirects = 0; ; redirects++) {
      signal?.throwIfAborted();
      const url = new URL(current);
      if ((url.protocol !== "http:" && url.protocol !== "https:") || url.username || url.password || url.hash) throw new TypeError("Media URL must be HTTP(S), without credentials or a fragment.");
      current = url.href;
      if (await withAbort(Promise.resolve().then(() => authorizeUrl(current)), signal) !== true) throw new Error("Media URL was not authorized.");
      signal?.throwIfAborted();
      const response = await fetch(current, { redirect: "manual", credentials: "omit", referrerPolicy: "no-referrer", ...(signal ? { signal } : {}) });
      if ([301, 302, 303, 307, 308].includes(response.status)) {
        const location = response.headers.get("location");
        await response.body?.cancel();
        if (!location || redirects >= maxRedirects) throw new Error("Media redirect is missing a location or exceeds maxRedirects.");
        current = new URL(location, current).href;
        continue;
      }
      if (response.status !== 200 || response.type === "opaqueredirect") {
        await response.body?.cancel();
        throw new Error(`Media HTTP response is not an accessible 200 (status ${response.status}).`);
      }
      if (!response.body) throw new Error("Media HTTP response has no readable body.");
      const reader = response.body.getReader(), chunks: Uint8Array[] = []; let size = 0;
      try {
        while (true) {
          signal?.throwIfAborted();
          const result = await withAbort(reader.read(), signal);
          if (result.done) break;
          size += result.value.byteLength;
          if (size > maxBytes) throw new RangeError(`Media exceeds maxBytes (${maxBytes}).`);
          chunks.push(new Uint8Array(result.value));
        }
      } catch (error) { void reader.cancel().catch(() => {}); throw error; }
      finally { reader.releaseLock(); }
      signal?.throwIfAborted();
      const bytes = new Uint8Array(size); let offset = 0;
      for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
      const mimeType = response.headers.get("content-type")?.split(";", 1)[0]?.trim().toLowerCase();
      return { bytes, ...(mimeType && mimeType !== "application/octet-stream" ? { mimeType } : {}) };
    }
  };
}
