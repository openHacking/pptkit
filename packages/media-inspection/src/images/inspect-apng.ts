import { capturePngInput, inspectCapturedPng, type PngInspection, type PngInspectionOptions } from "./inspect-png.js";

export interface ApngFrameInfo { width: number; height: number; x: number; y: number; delayNumerator: number; delayDenominator: number; disposal: number; blend: number; source: "IDAT" | "fdAT" }
export interface ApngInspection { png: PngInspection; animation?: { plays: number; defaultImageIsFrame: boolean; frames: ApngFrameInfo[] } }
export interface ApngInspectionOptions extends PngInspectionOptions { maxFrames?: number }

/** Checks selected APNG control structure, not compressed data or animation playback. */
export function inspectApng(input: Uint8Array, options: ApngInspectionOptions = {}): ApngInspection {
  const maxFrames = options.maxFrames === undefined ? 10000 : options.maxFrames;
  if (!Number.isSafeInteger(maxFrames) || maxFrames <= 0) throw new RangeError("maxFrames must be a positive safe integer.");
  const bytes = capturePngInput(input, options), png = inspectCapturedPng(bytes, options), fail = (reason: string): never => { throw new Error(`Invalid APNG: ${reason}`); };
  const abort = () => { if (options.signal?.aborted) throw options.signal.reason ?? new Error("APNG inspection aborted."); };
  const u32 = (i: number) => ((bytes[i]! * 0x1000000) + (bytes[i+1]! << 16) + (bytes[i+2]! << 8) + bytes[i+3]!) >>> 0;
  const u16 = (i: number) => (bytes[i]! << 8) | bytes[i+1]!;
  let at = 8, declared: number | undefined, plays = 0, sequence = 0, seenIdat = false, defaultImageIsFrame = false, frameHasData = false;
  const frames: ApngFrameInfo[] = [];
  while (at < bytes.length) {
    abort(); const length = u32(at), type = String.fromCharCode(...bytes.subarray(at + 4, at + 8)), data = at + 8;
    if (type === "acTL") {
      if (declared !== undefined || seenIdat || length !== 8) fail("duplicate, late or malformed acTL");
      declared = u32(data); plays = u32(data+4); if (!declared) fail("zero frame count"); if (declared > maxFrames) throw new RangeError("APNG maxFrames exceeded.");
    } else if (type === "fcTL") {
      if (declared === undefined || length !== 26 || u32(data) !== sequence++) fail("frame control or sequence");
      if (frames.length && !frameHasData) fail("previous frame has no data");
      if (frames.length >= maxFrames) throw new RangeError("APNG maxFrames exceeded.");
      const width=u32(data+4), height=u32(data+8), x=u32(data+12), y=u32(data+16), disposal=bytes[data+24]!, blend=bytes[data+25]!;
      if (!width || !height || x + width > png.width || y + height > png.height || disposal > 2 || blend > 1) fail("frame bounds or operations");
      const source = seenIdat ? "fdAT" : "IDAT";
      if (source === "IDAT") { if (frames.length || x || y || width !== png.width || height !== png.height) fail("default frame geometry or order"); defaultImageIsFrame = true; }
      frames.push({ width, height, x, y, delayNumerator:u16(data+20), delayDenominator:u16(data+22), disposal, blend, source }); frameHasData = false;
    } else if (type === "IDAT") {
      seenIdat = true; if (frames.length && frames[0]!.source === "IDAT") frameHasData = true;
    } else if (type === "fdAT") {
      if (declared === undefined || !seenIdat || !frames.length || frames[frames.length-1]!.source !== "fdAT" || length < 4 || u32(data) !== sequence++) fail("frame data or sequence");
      frameHasData = true;
    } else if (type === "IEND") break;
    at = data + length + 4;
  }
  if (declared === undefined) { if (png.animationChunks) fail("animation chunks without acTL"); return { png }; }
  if (frames.length !== declared || !frameHasData) fail("frame count or missing frame data");
  return { png, animation: { plays, defaultImageIsFrame, frames } };
}
