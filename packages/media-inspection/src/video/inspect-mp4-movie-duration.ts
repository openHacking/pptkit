import { captureImageInput } from "../images/capture-image-input.js";
import { inspectIsoBmff, type IsoBmffInspectionOptions } from "./inspect-iso-bmff.js";

/** Read a bounded version-zero movie clock; this does not validate tracks or codecs. */
export function inspectMp4MovieDuration(input: Uint8Array, options: IsoBmffInspectionOptions = {}): number {
  const { maxBytes = 64 * 1024 * 1024, maxBoxes = 4096, signal } = options;
  if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0 || maxBytes > 64 * 1024 * 1024) throw new RangeError("maxBytes must be between 1 and 67108864.");
  if (!Number.isSafeInteger(maxBoxes) || maxBoxes <= 0 || maxBoxes > 4096) throw new RangeError("maxBoxes must be between 1 and 4096.");
  const data = captureImageInput(input, "MP4", maxBytes, signal);
  const framing = inspectIsoBmff(data, { maxBytes, maxBoxes, ...(signal ? { signal } : {}) });
  const movies = framing.boxes.filter(box => box.type === "moov"), media = framing.boxes.filter(box => box.type === "mdat");
  if (movies.length !== 1 || media.length !== 1 || media[0]!.size <= media[0]!.headerSize || framing.boxes.some(box => box.type === "moof")) throw new TypeError("MP4 requires one movie box and one nonempty media-data box without fragments.");
  const movie = movies[0]!, view = new DataView(data.buffer, data.byteOffset, data.byteLength), end = movie.offset + movie.size;
  let at = movie.offset + movie.headerSize, count = 0, durationMs: number | undefined;
  while (at < end) {
    if (signal?.aborted) throw signal.reason ?? new Error("MP4 inspection aborted.");
    if (++count > 4096) throw new RangeError("MP4 movie box child limit exceeded.");
    if (end - at < 8) throw new TypeError("MP4 movie box has a truncated child header.");
    const shortSize = view.getUint32(at), type = String.fromCharCode(...data.subarray(at + 4, at + 8));
    let size = shortSize, header = 8;
    if (shortSize === 0) size = end - at;
    else if (shortSize === 1) {
      if (end - at < 16 || view.getUint32(at + 8) !== 0) throw new TypeError("MP4 movie box has an invalid extended child size.");
      header = 16; size = view.getUint32(at + 12);
    }
    if (size < header || size > end - at) throw new TypeError("MP4 movie box has an invalid child size.");
    if (type === "mvhd") {
      if (durationMs !== undefined || size - header < 100) throw new TypeError("MP4 needs one complete version-zero movie header.");
      const start = at + header;
      if (view.getUint32(start) !== 0) throw new TypeError("Only version-zero MP4 movie headers with zero flags are supported.");
      const scale = view.getUint32(start + 12), duration = view.getUint32(start + 16);
      if (!scale || !duration || duration === 0xffffffff) throw new TypeError("MP4 movie duration must be known and nonzero.");
      durationMs = Math.round(duration * 1000 / scale);
      if (!Number.isSafeInteger(durationMs) || durationMs < 1 || durationMs > 2147483647) throw new TypeError("MP4 movie duration must round to 1–2147483647 milliseconds.");
    }
    at += size;
  }
  if (durationMs === undefined) throw new TypeError("MP4 movie header is missing.");
  return durationMs;
}
