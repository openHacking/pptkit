import { readFile } from "node:fs/promises";
import type { LoadedAsset } from "../types/internal.js";
import { loadUrlAsset } from "./load-asset.js";

export async function loadNodeAsset(
  source: { type: "path" | "url"; value: string },
  declaredMime?: string,
): Promise<LoadedAsset> {
  if (source.type === "url") {
    return loadUrlAsset(source.value, declaredMime);
  }

  const data = new Uint8Array(await readFile(source.value));
  const mimeType = declaredMime ?? mimeFromPath(source.value);
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
