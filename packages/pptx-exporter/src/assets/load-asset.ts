import type { LoadedAsset } from "../types/internal.js";

export async function loadAsset(
  source: { type: "path" | "url"; value: string },
  declaredMime?: string,
): Promise<LoadedAsset> {
  if (source.type === "path") {
    throw new Error("Path assets are only supported by @pptkit/pptx-exporter/node.");
  }

  return loadUrlAsset(source.value, declaredMime);
}

export async function loadUrlAsset(value: string, declaredMime?: string): Promise<LoadedAsset> {
  const response = await fetch(value);
  if (!response.ok) throw new Error(`HTTP ${response.status} ${response.statusText}`);
  const data = new Uint8Array(await response.arrayBuffer());
  const mimeType = declaredMime ?? response.headers.get("content-type")?.split(";", 1)[0] ?? mimeFromExtension(value);
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
