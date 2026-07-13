import { readFile } from "node:fs/promises";
import { extname } from "node:path";
import type { LoadedAsset } from "../types/internal.js";

export async function loadAsset(
  source: { type: "path" | "url"; value: string },
  declaredMime?: string,
): Promise<LoadedAsset> {
  let data: Buffer;
  let mimeType = declaredMime;
  if (source.type === "path") {
    data = await readFile(source.value);
  } else {
    const response = await fetch(source.value);
    if (!response.ok) throw new Error(`HTTP ${response.status} ${response.statusText}`);
    data = Buffer.from(await response.arrayBuffer());
    mimeType ??= response.headers.get("content-type")?.split(";", 1)[0];
  }
  mimeType ??= mimeFromExtension(source.value);
  return { data, mimeType, extension: extensionForMime(mimeType) };
}

function mimeFromExtension(value: string): string {
  const extension = extname(value).toLowerCase();
  return extension === ".jpg" || extension === ".jpeg" ? "image/jpeg" : extension === ".gif" ? "image/gif" : "image/png";
}

function extensionForMime(mimeType: string): string {
  if (mimeType === "image/jpeg") return ".jpg";
  if (mimeType === "image/gif") return ".gif";
  if (mimeType === "image/svg+xml") return ".svg";
  return ".png";
}
