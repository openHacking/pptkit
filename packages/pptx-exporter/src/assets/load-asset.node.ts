import type { NormalizedAsset } from "@pptkit/core";
import type { AssetLoadContext } from "../types/internal.js";
import { loadAsset } from "./load-asset.js";
import { classifyOwnedImageBytes, packageImageMime } from "./load-bytes.js";
import { checkAssetSize } from "./limits.js";
import { open } from "node:fs/promises";
import type { LoadedAsset } from "../types/internal.js";

export async function loadNodeAsset(asset: NormalizedAsset, context: AssetLoadContext): Promise<LoadedAsset> {
  const { source, mimeType: declaredMime } = asset;
  if (source.type !== "path") return loadAsset(asset, context);
  context.signal?.throwIfAborted();
  const file = await open(source.value, "r");
  const chunks: Uint8Array[] = []; let total = 0;
  try {
    const stat = await file.stat();
    if (!stat.isFile()) throw new Error("Path asset must be a regular file.");
    checkAssetSize(stat.size, context.maxAssetBytes);
    while (true) {
      context.signal?.throwIfAborted();
      const buffer = new Uint8Array(65536);
      const { bytesRead } = await file.read(buffer);
      if (bytesRead === 0) break;
      total += bytesRead; checkAssetSize(total, context.maxAssetBytes); chunks.push(buffer.subarray(0, bytesRead));
    }
  } finally { await file.close(); }
  context.signal?.throwIfAborted();
  const data = new Uint8Array(total); let offset = 0;
  for (const chunk of chunks) { data.set(chunk, offset); offset += chunk.length; }
  if (context.externalImageValidation === "signature") return classifyOwnedImageBytes(data, declaredMime);
  const mimeType = packageImageMime(declaredMime ?? mimeFromPath(source.value));
  return { data, mimeType, extension: extensionForMime(mimeType) };
}

function mimeFromPath(value: string): string {
  const dot = value.lastIndexOf(".");
  const extension = dot === -1 ? "" : value.slice(dot).toLowerCase();
  return extension === ".jpg" || extension === ".jpeg" ? "image/jpeg" : extension === ".gif" ? "image/gif" : "image/png";
}

function extensionForMime(mimeType: string): string {
  if (mimeType === "image/jpeg") return ".jpg";
  if (mimeType === "image/gif") return ".gif";
  if (mimeType === "image/svg+xml") return ".svg";
  return ".png";
}
