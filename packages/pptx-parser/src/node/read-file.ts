import { constants, type BigIntStats } from "node:fs";
import { createHash } from "node:crypto";
import path from "node:path";
import type { FileIo } from "./file-io.js";
import { PptxFileError } from "../types/file.js";

export const fileHash = (bytes: Uint8Array): string => createHash("sha256").update(bytes).digest("hex");
export const errorCode = (error: unknown): string | undefined => error && typeof error === "object" && "code" in error && typeof error.code === "string" ? error.code : undefined;
export const checkAbort = (signal: AbortSignal | undefined, target: string): void => {
  if (signal?.aborted) throw new PptxFileError("operation-aborted", "PPTX file operation was aborted before commit.", target);
};
export function checkPlatform(target: string): void {
  if (!constants.O_NOFOLLOW || !constants.O_NONBLOCK) throw new PptxFileError("file-platform-unsupported", "This file profile requires no-follow and nonblocking filesystem open flags.", target);
}
export const sameFile = (a: BigIntStats, b: BigIntStats, includeChangeTime = false): boolean =>
  a.dev === b.dev && a.ino === b.ino && a.size === b.size && a.mode === b.mode && a.mtimeNs === b.mtimeNs && (!includeChangeTime || a.ctimeNs === b.ctimeNs);
export async function canonicalFile(io: FileIo, file: string): Promise<string> {
  if (typeof file !== "string" || !file.length || file.includes("\0")) throw new TypeError("PPTX file path must be a nonempty string without NUL.");
  const absolute = path.resolve(file);
  return path.join(await io.realpath(path.dirname(absolute)), path.basename(absolute));
}
/** Bounded regular-file capture; rejects name/identity/content changes during read. */
export async function readStableFile(io: FileIo, target: string, maxBytes: number, signal?: AbortSignal): Promise<{ bytes: Uint8Array; stat: BigIntStats; sha256: string }> {
  checkAbort(signal, target);
  const before = await io.lstat(target, { bigint: true }); checkAbort(signal, target);
  if (!before.isFile()) throw new PptxFileError("file-not-regular", "PPTX input/replace targets must be regular files, not symlinks or special files.", target);
  if (before.size > BigInt(maxBytes)) throw new PptxFileError("file-limit-exceeded", "PPTX file exceeds its captured byte limit.", target);
  const file = await io.open(target, constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_NONBLOCK);
  try {
    checkAbort(signal, target);
    const start = await file.stat({ bigint: true });
    if (!start.isFile() || !sameFile(before, start, true)) throw new PptxFileError("file-source-changed", "PPTX file changed before its bounded read.", target);
    const bytes = new Uint8Array(Number(start.size));
    for (let offset = 0; offset < bytes.length;) {
      checkAbort(signal, target);
      const result = await file.read(bytes, offset, Math.min(64 * 1024, bytes.length - offset), offset);
      if (!result.bytesRead) throw new PptxFileError("file-source-changed", "PPTX file shrank during its bounded read.", target);
      offset += result.bytesRead;
    }
    checkAbort(signal, target);
    const extra = await file.read(new Uint8Array(1), 0, 1, bytes.length);
    const end = await file.stat({ bigint: true }), named = await io.lstat(target, { bigint: true });
    checkAbort(signal, target);
    if (extra.bytesRead || !named.isFile() || !sameFile(start, end, true) || !sameFile(end, named, true)) throw new PptxFileError("file-source-changed", "PPTX file or its name changed during read.", target);
    return { bytes, stat: end, sha256: fileHash(bytes) };
  } finally {
    try { await file.close(); }
    catch (error) { try { await file.close(); } catch { /* original close error remains authoritative */ } throw error; }
  }
}
