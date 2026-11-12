import type { NormalizedAsset } from "@pptkit/core";
import type { AssetLoadContext } from "../types/internal.js";
import { checkAssetSize, withAbort } from "./limits.js";
import { classifyOwnedImageBytes, decodeAssetBytes, loadImageBytes, packageImageMime } from "./load-bytes.js";
import type { LoadedAsset } from "../types/internal.js";

export async function loadAsset(
  asset: NormalizedAsset,
  context: AssetLoadContext,
): Promise<LoadedAsset> {
  const { source, mimeType: declaredMime } = asset;
  context.signal?.throwIfAborted();
  if (source.type === "bytes") {
    checkAssetSize(source.value.length / 4 * 3 - (source.value.endsWith("==") ? 2 : source.value.endsWith("=") ? 1 : 0), context.maxAssetBytes);
    return loadImageBytes(decodeAssetBytes(source.value), declaredMime);
  }
  if (source.type === "path") {
    throw new Error("Path assets are only supported by @pptkit/pptx-exporter/node.");
  }

  return loadUrlAsset(source.value, declaredMime, context);
}

export async function loadUrlAsset(value: string, declaredMime: string | undefined, context: AssetLoadContext): Promise<LoadedAsset> {
  const response = await fetch(value, context.signal ? { signal: context.signal } : undefined);
  if (!response.ok) { await response.body?.cancel(); throw new Error(`HTTP ${response.status} ${response.statusText}`); }
  const chunks: Uint8Array[] = []; let total = 0;
  const reader = response.body?.getReader();
  if (reader) {
    try {
      while (true) {
        context.signal?.throwIfAborted();
        const { value: chunk, done } = await withAbort(reader.read(), context.signal);
        if (done) break;
        total += chunk.length; checkAssetSize(total, context.maxAssetBytes); chunks.push(new Uint8Array(chunk));
      }
    } catch (error) { await reader.cancel().catch(() => {}); throw error; }
    finally { reader.releaseLock(); }
  }
  context.signal?.throwIfAborted();
  const data = new Uint8Array(total); let offset = 0;
  for (const chunk of chunks) { data.set(chunk, offset); offset += chunk.length; }
  const responseMime = response.headers.get("content-type")?.split(";", 1)[0];
  if (context.externalImageValidation === "signature") {
    const loaded = classifyOwnedImageBytes(data, declaredMime);
    const headerMime = responseMime?.trim().toLowerCase();
    if (headerMime && headerMime !== "application/octet-stream" && packageImageMime(headerMime) !== loaded.mimeType) {
      throw new Error(`HTTP Content-Type ${responseMime} does not match ${loaded.mimeType} signature.`);
    }
    return loaded;
  }
  const mimeType = packageImageMime(declaredMime ?? responseMime ?? mimeFromExtension(value));
  return { data, mimeType, extension: extensionForMime(mimeType) };
}

function mimeFromExtension(value: string): string {
  const pathname = urlPathname(value);
  const filename = pathname.slice(pathname.lastIndexOf("/") + 1);
  const dot = filename.lastIndexOf(".");
  const extension = dot === -1 ? "" : filename.slice(dot).toLowerCase();
  return extension === ".jpg" || extension === ".jpeg" ? "image/jpeg" : extension === ".gif" ? "image/gif" : "image/png";
}

function urlPathname(value: string): string {
  try {
    return new URL(value).pathname;
  } catch {
    return value.split(/[?#]/, 1)[0] ?? value;
  }
}

function extensionForMime(mimeType: string): string {
  if (mimeType === "image/jpeg") return ".jpg";
  if (mimeType === "image/gif") return ".gif";
  if (mimeType === "image/svg+xml") return ".svg";
  return ".png";
}
