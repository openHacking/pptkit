import { PptxImportError } from "../types/import.js";

/** Check marker/scan framing only; entropy coding and decoded pixels are not verified. */
export async function validateJpegReplacement(bytes: Uint8Array, maxWork: number, usage: { work: number }, checkpoint: () => Promise<void>): Promise<boolean> {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return false;
  let position = 2, pending = 0, frame = false, scans = 0, inScan = false, scanData = false;
  let checkQuantization = false, huffmanFrame = 0;
  const frameComponents = new Map<number, number>();
  const sequentialComponents = new Set<number>();
  // Previous low approximation bit for each coefficient; -1 means not yet sent.
  const progressiveCoefficients = new Map<number, Int8Array>();
  const quantizationTables = new Set<number>();
  const huffmanTables = new Set<number>();
  const consume = async (count: number) => {
    usage.work += count; pending += count;
    if (usage.work > maxWork) throw new PptxImportError("save-limit-exceeded", "JPEG replacement inspection exceeds maxSaveValidationWork.");
    if (pending >= 32768) { pending = 0; await checkpoint(); }
  };
  await checkpoint();
  while (position < bytes.length) {
    if (inScan && bytes[position] !== 0xff) {
      const start = position;
      while (position < bytes.length && bytes[position] !== 0xff && position - start < 32768) position++;
      if (position > start) scanData = true;
      await consume(position - start); continue;
    }
    if (bytes[position++] !== 0xff) return false;
    let fill = 1;
    while (position < bytes.length && bytes[position] === 0xff) { position++; fill++; if (fill === 32768) { await consume(fill); fill = 0; } }
    if (position >= bytes.length) return false;
    const marker = bytes[position++]!; await consume(fill + 1);
    if (inScan && marker === 0x00) { scanData = true; continue; }
    if (inScan && marker >= 0xd0 && marker <= 0xd7) continue;
    if (inScan && !scanData) return false;
    inScan = false;
    if (marker === 0xd9) return frame && scans > 0 && position === bytes.length &&
      (huffmanFrame !== 0xc0 && huffmanFrame !== 0xc1 || sequentialComponents.size === frameComponents.size) &&
      (huffmanFrame !== 0xc2 || [...frameComponents.keys()].every(id => (progressiveCoefficients.get(id)?.[0] ?? -1) >= 0));
    if (marker === 0x00 || marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7) || position + 2 > bytes.length) return false;
    const length = (bytes[position]! << 8) | bytes[position + 1]!;
    if (length < 2 || length > bytes.length - position) return false;
    const end = position + length;
    if (marker === 0xdb) {
      let offset = position + 2;
      if (offset === end) return false;
      while (offset < end) {
        const spec = bytes[offset++]!, precision = spec >>> 4;
        if (precision > 1 || (spec & 15) > 3 || end - offset < 64 * (precision + 1)) return false;
        for (let i = 0; i < 64; i++) {
          const value = precision ? (bytes[offset]! << 8) | bytes[offset + 1]! : bytes[offset]!;
          if (!value) return false;
          offset += precision + 1;
        }
        quantizationTables.add(spec & 15);
      }
    }
    if (marker === 0xc4) {
      let offset = position + 2;
      if (offset === end) return false;
      while (offset < end) {
        const spec = bytes[offset++]!;
        if ((spec >>> 4) > 1 || (spec & 15) > 3 || end - offset < 16) return false;
        let slots = 1, symbols = 0;
        for (let i = 0; i < 16; i++) { const count = bytes[offset++]!; slots = slots * 2 - count; symbols += count; if (slots < 0) return false; }
        if (!symbols || symbols > 256 || symbols > end - offset) return false;
        offset += symbols;
        huffmanTables.add(spec);
      }
    }
    if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
      if (frame || length < 8 || !bytes[position + 7] || length !== 8 + 3 * bytes[position + 7]!) return false;
      const precision = bytes[position + 2]!;
      if (marker === 0xc0 && precision !== 8 || (marker === 0xc1 || marker === 0xc2) && precision !== 8 && precision !== 12) return false;
      for (let i = 0; i < bytes[position + 7]!; i++) {
        const id = bytes[position + 8 + 3 * i]!, sampling = bytes[position + 9 + 3 * i]!, table = bytes[position + 10 + 3 * i]!;
        if (frameComponents.has(id) || (sampling >>> 4) < 1 || (sampling >>> 4) > 4 || (sampling & 15) < 1 || (sampling & 15) > 4 || table > 3) return false;
        frameComponents.set(id, table);
      }
      checkQuantization = marker === 0xc0 || marker === 0xc1 || marker === 0xc2;
      huffmanFrame = marker >= 0xc0 && marker <= 0xc2 ? marker : 0;
      frame = true;
    }
    if (marker === 0xda) {
      if (!frame || length < 6 || !bytes[position + 2] || length !== 6 + 2 * bytes[position + 2]!) return false;
      const scanComponents = new Set<number>();
      const spectralStart = bytes[position + 3 + 2 * bytes[position + 2]!]!;
      const spectralEnd = bytes[position + 4 + 2 * bytes[position + 2]!]!;
      const approximation = bytes[position + 5 + 2 * bytes[position + 2]!]!;
      const approximationHigh = approximation >>> 4, approximationLow = approximation & 15;
      if (huffmanFrame === 0xc2 && ((spectralStart === 0 && spectralEnd !== 0) || spectralStart > spectralEnd || spectralEnd > 63 || (spectralStart !== 0 && bytes[position + 2] !== 1) || (approximationHigh !== 0 && approximationHigh !== approximationLow + 1) || approximationHigh > 13 || approximationLow > 13)) return false;
      if ((huffmanFrame === 0xc0 || huffmanFrame === 0xc1) && (spectralStart !== 0 || spectralEnd !== 63 || approximation !== 0)) return false;
      for (let i = 0; i < bytes[position + 2]!; i++) {
        const id = bytes[position + 3 + 2 * i]!;
        if (!frameComponents.has(id) || scanComponents.has(id)) return false;
        if ((huffmanFrame === 0xc0 || huffmanFrame === 0xc1) && sequentialComponents.has(id)) return false;
        if (checkQuantization && !quantizationTables.has(frameComponents.get(id)!)) return false;
        if (huffmanFrame) {
          const selectors = bytes[position + 4 + 2 * i]!, dc = selectors >>> 4, ac = selectors & 15;
          const needsDc = huffmanFrame !== 0xc2 || (spectralStart === 0 && approximationHigh === 0);
          const needsAc = huffmanFrame !== 0xc2 || spectralStart !== 0;
          const limit = huffmanFrame === 0xc0 ? 1 : 3;
          if ((needsDc && (dc > limit || !huffmanTables.has(dc))) || (needsAc && (ac > limit || !huffmanTables.has(0x10 | ac)))) return false;
        }
        scanComponents.add(id);
      }
      if (huffmanFrame === 0xc0 || huffmanFrame === 0xc1) for (const id of scanComponents) sequentialComponents.add(id);
      if (huffmanFrame === 0xc2) {
        for (const id of scanComponents) {
          let coefficients = progressiveCoefficients.get(id);
          if (!coefficients) { coefficients = new Int8Array(64).fill(-1); progressiveCoefficients.set(id, coefficients); }
          if (spectralStart !== 0 && coefficients[0] === -1) return false;
          await consume(spectralEnd - spectralStart + 1);
          for (let coefficient = spectralStart; coefficient <= spectralEnd; coefficient++) {
            const previous = coefficients[coefficient]!;
            if (previous === -1 ? approximationHigh !== 0 : approximationHigh !== previous || approximationLow !== previous - 1) return false;
            coefficients[coefficient] = approximationLow;
          }
        }
      }
      scans++; inScan = true; scanData = false;
    }
    position += length; await consume(length);
  }
  return false;
}
