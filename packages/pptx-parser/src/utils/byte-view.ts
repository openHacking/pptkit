const typedArrayPrototype = Object.getPrototypeOf(Uint8Array.prototype);
const lengthGetter = Object.getOwnPropertyDescriptor(typedArrayPrototype, "byteLength")!.get!;
const bufferGetter = Object.getOwnPropertyDescriptor(typedArrayPrototype, "buffer")!.get!;
const sharedLength = typeof SharedArrayBuffer === "undefined" ? undefined : Object.getOwnPropertyDescriptor(SharedArrayBuffer.prototype, "byteLength")!.get!;

/** Use intrinsic view properties; subclass getters and iterators are untrusted. */
export function byteViewInfo(input: Uint8Array): { length: number; shared: boolean } {
  const length = lengthGetter.call(input) as number, buffer = bufferGetter.call(input);
  if (!sharedLength) return { length, shared: false };
  try { sharedLength.call(buffer); return { length, shared: true }; } catch { return { length, shared: false }; }
}

export function copyByteView(input: Uint8Array): Uint8Array<ArrayBuffer> { return new Uint8Array(input); }
