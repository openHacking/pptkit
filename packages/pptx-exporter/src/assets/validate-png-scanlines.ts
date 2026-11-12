import type { ApngInspection } from "./inspect-apng.js";
import { withAbort } from "./limits.js";

type Span = { start: number; end: number };
type Frame = { width: number; height: number; spans: Span[] };
const MAX_SCANLINE_BYTES = 128 * 1024 * 1024;
const adam7 = [[0, 0, 8, 8], [4, 0, 8, 8], [0, 4, 4, 8], [2, 0, 4, 4], [0, 2, 2, 4], [1, 0, 2, 2], [0, 1, 1, 2]] as const;

/** Internal: inspectApng has already verified chunk framing, CRCs, and frame ordering. */
export async function validatePngScanlines(bytes: Uint8Array, info: ApngInspection, signal?: AbortSignal): Promise<void> {
  const fail = (message: string): never => { throw new Error(`Invalid PNG scanlines: ${message}`); };
  if (typeof DecompressionStream !== "function") throw new Error("DecompressionStream is required to validate PNG scanlines.");
  const { width, height, colorType, bitDepth, interlaced } = info.png;
  const channels = [1, 0, 3, 1, 2, 0, 4][colorType]!;
  const base: Frame = { width, height, spans: [] };
  const frames = [base];
  let active = base, controlIndex = 0, paletteEntries = 0;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let at = 8; at < bytes.length;) {
    signal?.throwIfAborted();
    const length = view.getUint32(at);
    const code = String.fromCharCode(...bytes.subarray(at + 4, at + 8));
    const start = at + 8, end = start + length;
    if (code === "IDAT" && length) base.spans.push({ start, end });
    if (code === "PLTE") paletteEntries = length / 3;
    if (code === "fcTL") {
      const frame = info.animation!.frames[controlIndex++]!;
      if (frame.source === "fdAT") {
        active = { width: frame.width, height: frame.height, spans: [] };
        frames.push(active);
      }
    }
    if (code === "fdAT" && length > 4) active.spans.push({ start: start + 4, end });
    at = end + 4;
  }

  let remaining = MAX_SCANLINE_BYTES;
  for (const frame of frames) {
    const passes: { width: number; rows: number; bytes: number }[] = [];
    let expected = 0;
    for (const [x, y, dx, dy] of interlaced ? adam7 : [[0, 0, 1, 1] as const]) {
      const w = Math.ceil((frame.width - x) / dx), h = Math.ceil((frame.height - y) / dy);
      if (w <= 0 || h <= 0) continue;
      const row = 1 + Math.ceil(w * channels * bitDepth / 8);
      if (row > Math.floor((remaining - expected) / h)) throw new RangeError("PNG decoded scanlines exceed the 128 MiB limit.");
      expected += row * h;
      passes.push({ width: w, rows: h, bytes: row });
    }
    remaining -= expected;

    let spanIndex = 0, at = frame.spans[0]?.start ?? 0;
    const compressed = new ReadableStream<BufferSource>({
      pull(controller) {
        signal?.throwIfAborted();
        const span = frame.spans[spanIndex];
        if (!span) { controller.close(); return; }
        // Limit expansion per decompressor write, even if the declared image is tiny.
        const end = Math.min(at + 1024, span.end);
        controller.enqueue(new Uint8Array(bytes.subarray(at, end)));
        if (end === span.end) at = frame.spans[++spanIndex]?.start ?? 0;
        else at = end;
      },
    }, { highWaterMark: 0 });
    const reader = compressed.pipeThrough(new DecompressionStream("deflate")).getReader();
    let count = 0, pass = 0, row = 0, column = 0, filter = 0, upperLeft = 0;
    let prior = colorType === 3 ? new Uint8Array(passes[0]!.bytes - 1) : undefined;
    try {
      while (true) {
        const { value, done } = await withAbort(reader.read(), signal);
        if (done) break;
        if (value.length > expected - count) fail("decoded data exceeds the declared dimensions");
        count += value.length;
        for (const byte of value) {
          const current = passes[pass]!;
          if (column === 0) {
            if (byte > 4) fail("invalid row filter");
            filter = byte;
            upperLeft = 0;
          } else if (prior) {
            const index = column - 1, above = prior[index]!, left = index ? prior[index - 1]! : 0;
            let predictor = 0;
            if (filter === 1) predictor = left;
            else if (filter === 2) predictor = above;
            else if (filter === 3) predictor = (left + above) >>> 1;
            else if (filter === 4) {
              const p = left + above - upperLeft, a = Math.abs(p - left), b = Math.abs(p - above), c = Math.abs(p - upperLeft);
              predictor = a <= b && a <= c ? left : b <= c ? above : upperLeft;
            }
            const decoded = (byte + predictor) & 255;
            prior[index] = decoded;
            upperLeft = above;
            for (let bit = 0; bit < 8 && index * 8 + bit < current.width * bitDepth; bit += bitDepth) {
              if (((decoded >>> (8 - bitDepth - bit)) & ((1 << bitDepth) - 1)) >= paletteEntries) fail("palette index is out of range");
            }
          }
          if (++column === current.bytes) {
            column = 0;
            if (++row === current.rows) {
              row = 0;
              pass++;
              if (prior && pass < passes.length) prior = new Uint8Array(passes[pass]!.bytes - 1);
            }
          }
        }
      }
      if (count !== expected) fail("decoded data is shorter than the declared dimensions");
    } catch (error) {
      await reader.cancel().catch(() => {});
      signal?.throwIfAborted();
      throw new Error(`PNG image data validation failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally { reader.releaseLock(); }
  }
}
