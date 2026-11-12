export class ImageInspectionLimitError extends RangeError {}

/** Capture a data URL's bytes within a caller-provided budget; not an image validator. */
export function captureImageDataUrl(url: string, maxBytes: number): Uint8Array {
  const comma = url.indexOf(",");
  if (comma < 0 || !/^data:image\//i.test(url)) throw new TypeError("Invalid image data URL.");
  const metadata = url.slice(5, comma), payload = url.slice(comma + 1);
  if (/;base64$/i.test(metadata)) {
    if (payload.length % 4 !== 0) throw new TypeError("Image Base64 must be canonical.");
    const size = payload.length / 4 * 3 - (payload.endsWith("==") ? 2 : payload.endsWith("=") ? 1 : 0);
    if (size > maxBytes) throw new ImageInspectionLimitError("Image inspection byte budget exceeded.");
    if (!/^[A-Za-z0-9+/]*={0,2}$/.test(payload)) throw new TypeError("Image Base64 must be canonical.");
    const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    if ((payload.endsWith("==") && (alphabet.indexOf(payload[payload.length - 3]!) & 15) !== 0) || (payload.endsWith("=") && !payload.endsWith("==") && (alphabet.indexOf(payload[payload.length - 2]!) & 3) !== 0)) throw new TypeError("Image Base64 must be canonical.");
    return new Uint8Array(Buffer.from(payload, "base64"));
  }
  const output = new Uint8Array(Math.min(payload.length, maxBytes + 1));
  let length = 0;
  for (let i = 0; i < payload.length; i++) {
    const code = payload.charCodeAt(i); let byte: number;
    if (code === 37) {
      const digits = payload.slice(i + 1, i + 3);
      if (!/^[0-9a-f]{2}$/i.test(digits)) throw new TypeError("Invalid image URL escape.");
      byte = Number.parseInt(digits, 16); i += 2;
    } else {
      if (code > 127 || code === 35) throw new TypeError("Unsupported raw image URL character.");
      byte = code;
    }
    if (length >= maxBytes) throw new ImageInspectionLimitError("Image inspection byte budget exceeded.");
    output[length++] = byte;
  }
  return output.slice(0, length);
}
