import type { PresentationAssetSource, PresentationAssetSourceInput } from "../types/asset.js";

const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
const typedArrayBuffer = Object.getOwnPropertyDescriptor(Object.getPrototypeOf(Uint8Array.prototype), "buffer")!.get!;
const sharedBufferLength = typeof SharedArrayBuffer === "undefined" ? undefined : Object.getOwnPropertyDescriptor(SharedArrayBuffer.prototype, "byteLength")!.get!;

function sharedBuffer(value: unknown): boolean {
  if (!sharedBufferLength) return false;
  try { sharedBufferLength.call(value); return true; } catch { return false; }
}

export function isCanonicalBase64(value: unknown): value is string {
  if (typeof value !== "string" || value.length === 0 || value.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(value)) return false;
  if (value.endsWith("==")) return (alphabet.indexOf(value[value.length - 3]!) & 15) === 0;
  if (value.endsWith("=")) return (alphabet.indexOf(value[value.length - 2]!) & 3) === 0;
  return true;
}

/** Pure conversion: no filesystem, fetching, image decoding or format-package state. */
export function captureAssetSource(source: PresentationAssetSourceInput): PresentationAssetSource {
  if (source.type !== "bytes") return { ...source };
  const value = source.value;
  if (typeof value === "string") {
    if (source.encoding !== "base64" || !isCanonicalBase64(value)) throw new TypeError("Byte asset sources require nonempty canonical Base64 and encoding: base64.");
    return { type: "bytes", encoding: "base64", value };
  }
  if (!(value instanceof Uint8Array) || (source.encoding !== undefined && source.encoding !== "raw")) throw new TypeError("Raw byte asset sources require Uint8Array.");
  if (sharedBuffer(typedArrayBuffer.call(value))) throw new TypeError("Byte assets require an ordinary buffer, not SharedArrayBuffer.");
  const bytes = new Uint8Array(value);
  if (bytes.length === 0) throw new TypeError("Byte asset sources cannot be empty.");
  // Every non-final chunk is divisible by three, so concatenated Base64 is canonical.
  const chunks: string[] = [];
  for (let offset = 0; offset < bytes.length; offset += 24576) chunks.push(btoa(String.fromCharCode(...bytes.subarray(offset, offset + 24576))));
  return { type: "bytes", encoding: "base64", value: chunks.join("") };
}
