const textEncoder = new TextEncoder();

export function encodeUtf8(value: string): Uint8Array {
  return textEncoder.encode(value);
}
