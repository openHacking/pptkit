import { PackageInspectionError } from "../types/archive.js";

// RFC 1951 structural scan: no output allocation and no dependency on the
// inflater's private cursor. CRC and actual decoded bytes are checked separately.
const LENGTH_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LENGTH_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DISTANCE_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const DISTANCE_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const CODE_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
interface Tree { counts: number[]; symbols: number[]; max: number }

export async function validateDeflate(bytes: Uint8Array, entry: string, acceptSize: (size: number) => void, signal: AbortSignal | undefined, maxDistance = 32768): Promise<void> {
  let bit = 0, output = 0, operations = 0;
  const invalid = (message: string): never => { throw new PackageInspectionError("archive-invalid", message, entry); };
  const abort = (): void => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "Archive inspection was aborted.", entry); };
  const read = (width: number): number => {
    if (bit + width > bytes.length * 8) invalid("Truncated DEFLATE stream.");
    let value = 0;
    for (let i = 0; i < width; i++, bit++) value |= ((bytes[bit >>> 3]! >>> (bit & 7)) & 1) << i;
    return value;
  };
  const tree = (lengths: number[], kind: "codes" | "literal" | "distance"): Tree => {
    const counts = Array<number>(16).fill(0);
    let max = 0;
    for (const length of lengths) { counts[length] = counts[length]! + 1; max = Math.max(max, length); }
    if (max === 0) {
      if (kind !== "distance") invalid("Empty DEFLATE Huffman tree.");
      return { counts, symbols: [], max };
    }
    let available = 1;
    for (let length = 1; length <= 15; length++) {
      available = available * 2 - counts[length]!;
      if (available < 0) invalid("Oversubscribed DEFLATE Huffman tree.");
    }
    if (available && (kind === "codes" || max !== 1)) invalid("Incomplete DEFLATE Huffman tree.");
    const symbols: number[] = [];
    for (let length = 1; length <= max; length++) for (let symbol = 0; symbol < lengths.length; symbol++) if (lengths[symbol] === length) symbols.push(symbol);
    return { counts, symbols, max };
  };
  const decode = (codes: Tree): number => {
    let code = 0, first = 0, offset = 0;
    for (let length = 1; length <= codes.max; length++) {
      code = code * 2 + read(1);
      const count = codes.counts[length]!;
      if (code >= first && code < first + count) return codes.symbols[offset + code - first]!;
      offset += count; first = (first + count) * 2;
    }
    return invalid("Invalid DEFLATE Huffman symbol.");
  };
  const addOutput = (length: number): void => { output += length; acceptSize(output); };
  const checkpoint = async (): Promise<void> => { await new Promise<void>((resolve) => setTimeout(resolve, 0)); abort(); };
  const fixedLiteral = tree(Array.from({ length: 288 }, (_, i) => i < 144 ? 8 : i < 256 ? 9 : i < 280 ? 7 : 8), "literal");
  const fixedDistance = tree(Array<number>(32).fill(5), "distance");
  let final = 0;
  abort();
  do {
    final = read(1);
    const kind = read(2);
    if (kind === 3) invalid("Reserved DEFLATE block type.");
    if (kind === 0) {
      bit = Math.ceil(bit / 8) * 8;
      const length = read(16), complement = read(16);
      if ((length ^ complement) !== 0xffff) invalid("DEFLATE stored block length mismatch.");
      if (bit + length * 8 > bytes.length * 8) invalid("Truncated DEFLATE stored block.");
      bit += length * 8; addOutput(length);
    } else {
      let literal = fixedLiteral, distance = fixedDistance;
      if (kind === 2) {
        const literals = read(5) + 257, distances = read(5) + 1, codeCount = read(4) + 4;
        if (literals > 286) invalid("Reserved DEFLATE literal count.");
        const codeLengths = Array<number>(19).fill(0);
        for (let i = 0; i < codeCount; i++) codeLengths[CODE_ORDER[i]!] = read(3);
        const codes = tree(codeLengths, "codes"), lengths: number[] = [];
        while (lengths.length < literals + distances) {
          const symbol = decode(codes);
          if (symbol < 16) lengths.push(symbol);
          else {
            if (symbol === 16 && !lengths.length) invalid("DEFLATE repeat has no preceding code length.");
            const repeat = symbol === 16 ? read(2) + 3 : symbol === 17 ? read(3) + 3 : read(7) + 11;
            const value = symbol === 16 ? lengths[lengths.length - 1]! : 0;
            if (lengths.length + repeat > literals + distances) invalid("DEFLATE code-length repeat exceeds its alphabet.");
            for (let i = 0; i < repeat; i++) lengths.push(value);
          }
        }
        if (!lengths[256]) invalid("DEFLATE literal tree has no end-of-block symbol.");
        literal = tree(lengths.slice(0, literals), "literal");
        distance = tree(lengths.slice(literals), "distance");
      }
      while (true) {
        const symbol = decode(literal);
        if (symbol === 256) break;
        if (symbol < 256) addOutput(1);
        else {
          if (symbol > 285) invalid("Reserved DEFLATE length symbol.");
          const index = symbol - 257, length = LENGTH_BASE[index]! + read(LENGTH_EXTRA[index]!);
          const distanceSymbol = decode(distance);
          if (distanceSymbol > 29) invalid("Reserved DEFLATE distance symbol.");
          const back = DISTANCE_BASE[distanceSymbol]! + read(DISTANCE_EXTRA[distanceSymbol]!);
          if (back > output) invalid("DEFLATE reference precedes available output.");
          if (back > maxDistance) invalid("DEFLATE reference exceeds the advertised window.");
          addOutput(length);
        }
        if (++operations % 4096 === 0) await checkpoint();
      }
    }
    // Empty blocks must yield too: they can otherwise consume the entire input
    // without visiting any output-producing symbol.
    if (++operations % 4096 === 0) await checkpoint();
  } while (!final);
  if (Math.ceil(bit / 8) !== bytes.length) invalid("Trailing bytes after the final DEFLATE block.");
  // Unused bits in the last byte are padding, not a second byte/stream.
  acceptSize(output); abort();
}
