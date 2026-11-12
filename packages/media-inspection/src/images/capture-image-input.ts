const typedArrayPrototype = Object.getPrototypeOf(Uint8Array.prototype);
const byteLengthGetter = Object.getOwnPropertyDescriptor(typedArrayPrototype, "byteLength")!.get!;
const bufferGetter = Object.getOwnPropertyDescriptor(typedArrayPrototype, "buffer")!.get!;
const sharedBufferLength = typeof SharedArrayBuffer === "undefined" ? undefined : Object.getOwnPropertyDescriptor(SharedArrayBuffer.prototype, "byteLength")!.get!;

function sharedBuffer(value: unknown): boolean {
  if (!sharedBufferLength) return false;
  try { sharedBufferLength.call(value); return true; } catch { return false; }
}

/** Snapshot the actual typed-array view without invoking subclass overrides. */
export function captureImageInput(input: Uint8Array, format: string, maxBytes: number, signal?: AbortSignal, limitName = "maxBytes"): Uint8Array<ArrayBuffer> {
  if (typeof format !== "string" || !format || typeof limitName !== "string" || !limitName) throw new TypeError("Image capture labels must be nonempty strings.");
  if (!Number.isSafeInteger(maxBytes) || maxBytes < 0) throw new RangeError("Image capture maxBytes must be a nonnegative safe integer.");
  if (!(input instanceof Uint8Array)) throw new TypeError(`${format} input must be Uint8Array.`);
  const buffer = bufferGetter.call(input), length = byteLengthGetter.call(input) as number;
  if (sharedBuffer(buffer)) throw new TypeError(`${format} input must not use SharedArrayBuffer.`);
  if (length > maxBytes) throw new RangeError(format === "Asset" ? `Asset exceeds maxAssetBytes (${maxBytes}).` : `${format} ${limitName} exceeded.`);
  if (signal?.aborted) throw signal.reason ?? new Error(`${format} inspection aborted.`);
  return new Uint8Array(input);
}
