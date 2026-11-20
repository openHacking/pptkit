import { constants } from "node:fs";
import { open } from "node:fs/promises";
import { isAbsolute, resolve } from "node:path";
import { withAbort } from "../assets/limits.js";
import type { PptxMediaResolver } from "../types/export.js";

export interface FileMediaResolverOptions {
  /** Authorize each normalized absolute path before opening it. */
  authorizePath: (path: string) => boolean | Promise<boolean>;
}

/** Optional Node adapter. Filesystem isolation and ancestor symlinks remain caller-owned. */
export function createFileMediaResolver(options: FileMediaResolverOptions): PptxMediaResolver {
  if (!options || typeof options.authorizePath !== "function") throw new TypeError("authorizePath must be a function.");
  const { authorizePath } = options;
  return async (asset, { signal, maxBytes }) => {
    if (asset.source.type !== "path" || !isAbsolute(asset.source.value)) throw new TypeError("File media resolver requires an absolute path source.");
    if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0) throw new RangeError("maxBytes must be a positive safe integer.");
    if (!constants.O_NOFOLLOW || !constants.O_NONBLOCK) throw new Error("File media resolver requires O_NOFOLLOW and O_NONBLOCK support.");
    signal?.throwIfAborted();
    const path = resolve(asset.source.value);
    if (await withAbort(Promise.resolve().then(() => authorizePath(path)), signal) !== true) throw new Error("Media path was not authorized.");
    signal?.throwIfAborted();
    const file = await open(path, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK);
    const chunks: Uint8Array[] = []; let total = 0;
    try {
      signal?.throwIfAborted();
      const stat = await file.stat();
      if (!stat.isFile()) throw new TypeError("Media path must identify a regular file.");
      if (stat.size > maxBytes) throw new RangeError(`Media exceeds maxBytes (${maxBytes}).`);
      while (true) {
        signal?.throwIfAborted();
        const chunk = new Uint8Array(Math.min(65536, maxBytes - total + 1));
        const { bytesRead } = await file.read(chunk, 0, chunk.length, null);
        signal?.throwIfAborted();
        if (bytesRead === 0) break;
        total += bytesRead;
        if (total > maxBytes) throw new RangeError(`Media exceeds maxBytes (${maxBytes}).`);
        chunks.push(chunk.subarray(0, bytesRead));
      }
    } finally { await file.close(); }
    signal?.throwIfAborted();
    const bytes = new Uint8Array(total); let offset = 0;
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
    return { bytes };
  };
}
