import { zipSync } from "fflate";
import type { ZipPart } from "../types/internal.js";

const ZIP_EPOCH = new Date(1980, 0, 1);

export function createZip(parts: ZipPart[]): Uint8Array {
  return zipSync(Object.fromEntries(parts.map((part) => [part.name, part.data])), {
    level: 6,
    mtime: ZIP_EPOCH,
  });
}
