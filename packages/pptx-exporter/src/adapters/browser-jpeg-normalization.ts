import { inspectJpegExifOrientation } from '../assets/inspect-jpeg-exif.js';
import { withAbort } from '../assets/limits.js';
import { captureImageInput } from '../assets/capture-image-input.js';

export interface BrowserJpegNormalizationOptions {
  maxInputBytes?: number;
  maxPixels?: number;
  maxOutputBytes?: number;
  signal?: AbortSignal;
}
export interface BrowserJpegNormalizationResult {
  bytes: Uint8Array;
  mimeType: 'image/png';
  orientationNormalization: 'pixels';
  width: number;
  height: number;
  sourceOrientation: 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8;
}

/** Explicit browser decoder adapter; color/metadata conversion is browser-owned. */
export async function normalizeJpegOrientationInBrowser(input: Uint8Array, options: BrowserJpegNormalizationOptions = {}): Promise<BrowserJpegNormalizationResult> {
  const limit = (value: number | undefined, fallback: number, name: string) => {
    const result = value === undefined ? fallback : value;
    if (!Number.isSafeInteger(result) || result <= 0) throw new RangeError(`${name} must be a positive safe integer.`);
    return result;
  };
  const { signal } = options;
  const maxInputBytes = limit(options.maxInputBytes, 64 * 1024 * 1024, 'maxInputBytes');
  const maxPixels = limit(options.maxPixels, 40_000_000, 'maxPixels');
  const maxOutputBytes = limit(options.maxOutputBytes, 128 * 1024 * 1024, 'maxOutputBytes');
  signal?.throwIfAborted();
  const bytes = captureImageInput(input, 'JPEG', maxInputBytes, signal, 'maxInputBytes');
  const exif = inspectJpegExifOrientation(bytes, { maxBytes: maxInputBytes, ...(signal ? { signal } : {}) });
  if (exif.orientation === undefined || exif.stoppedAt !== 'scan') throw new Error('A pre-scan EXIF orientation and JPEG scan are required.');
  // Read the first frame header before allocating an ImageBitmap; no pixel decoding here.
  let offset = 2, width = 0, height = 0;
  while (offset < bytes.length) {
    signal?.throwIfAborted();
    if (bytes[offset++] !== 0xff) throw new Error('Invalid JPEG marker.');
    while (offset < bytes.length && bytes[offset] === 0xff) offset++;
    const marker = bytes[offset++]; if (marker === undefined || marker === 0xda || marker === 0xd9) break;
    if (offset + 2 > bytes.length) throw new Error('Truncated JPEG segment.');
    const length = bytes[offset]! * 256 + bytes[offset + 1]!;
    if (length < 2 || offset + length > bytes.length) throw new Error('Invalid JPEG segment length.');
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      if (length < 8) throw new Error('Invalid JPEG frame header.');
      height = bytes[offset + 3]! * 256 + bytes[offset + 4]!; width = bytes[offset + 5]! * 256 + bytes[offset + 6]!;
      const components = bytes[offset + 7]!;
      if (!width || !height || !components || length !== 8 + 3 * components) throw new Error('Unsupported JPEG frame dimensions.');
      break;
    }
    offset += length;
  }
  if (!width || !height) throw new Error('JPEG frame dimensions are unavailable.');
  if (width * height > maxPixels) throw new RangeError('JPEG maxPixels exceeded.');
  if (typeof document === 'undefined' || typeof createImageBitmap !== 'function') throw new Error('A browser document and createImageBitmap are required.');
  const pending = createImageBitmap(new Blob([bytes], { type: 'image/jpeg' }), { imageOrientation: 'from-image' }).then(bitmap => {
    if (signal?.aborted) { bitmap.close(); signal.throwIfAborted(); }
    return bitmap;
  });
  let bitmap: ImageBitmap;
  try { bitmap = await withAbort(pending, signal); }
  catch (error) {
    // The decoder may win the race immediately before the abort is observed.
    void pending.then(value => value.close(), () => {});
    throw error;
  }
  let canvas: HTMLCanvasElement | undefined;
  try {
    signal?.throwIfAborted();
    const swap = exif.orientation >= 5;
    if (bitmap.width !== (swap ? height : width) || bitmap.height !== (swap ? width : height)) throw new Error('Browser decoded dimensions do not match EXIF orientation.');
    canvas = document.createElement('canvas'); canvas.width = bitmap.width; canvas.height = bitmap.height;
    const context = canvas.getContext('2d'); if (!context) throw new Error('Canvas 2D context is unavailable.');
    context.drawImage(bitmap, 0, 0);
    const blob = await withAbort(new Promise<Blob>((resolve, reject) => canvas!.toBlob(value => value ? resolve(value) : reject(new Error('PNG encoding failed.')), 'image/png')), signal);
    signal?.throwIfAborted();
    if (blob.type !== 'image/png') throw new Error('Canvas did not produce PNG.');
    if (blob.size > maxOutputBytes) throw new RangeError('PNG maxOutputBytes exceeded.');
    const output = new Uint8Array(await withAbort(blob.arrayBuffer(), signal)); signal?.throwIfAborted();
    return { bytes: output, mimeType: 'image/png', orientationNormalization: 'pixels', width: bitmap.width, height: bitmap.height, sourceOrientation: exif.orientation };
  } finally { bitmap.close(); if (canvas) { canvas.width = 0; canvas.height = 0; } }
}
