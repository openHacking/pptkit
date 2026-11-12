import { open, stat } from "node:fs/promises";
import type { FileHandle } from "node:fs/promises";
import { constants, type BigIntStats } from "node:fs";
import { createHash } from "node:crypto";
import path from "node:path";
import { CliError } from "./arguments.js";
export function checkAbort(signal?: AbortSignal): void { if (signal?.aborted) throw new CliError("operation-aborted", "Command cancelled.", 130); }
type CaptureIo = { openFile: (path: string, flags: number) => Promise<FileHandle>; statTarget: (path: string) => Promise<BigIntStats> };
const defaultIo: CaptureIo = { openFile: open, statTarget: target => stat(target, { bigint: true }) };
function sameFile(a: BigIntStats, b: BigIntStats): boolean {
  return a.isFile() && b.isFile() && a.dev === b.dev && a.ino === b.ino && a.size === b.size && a.mtimeNs === b.mtimeNs && a.ctimeNs === b.ctimeNs;
}
/** Regular-file capture only. No stdin, directories, devices, sockets or network. */
export async function readInput(file: string, cwd: string, limit: number, signal?: AbortSignal, io: CaptureIo = defaultIo) {
  checkAbort(signal); const target = path.resolve(cwd, file);
  let handle;
  try {
    handle = await io.openFile(target, constants.O_RDONLY | constants.O_NONBLOCK); checkAbort(signal);
    const before = await handle.stat({ bigint: true });
    if (!before.isFile()) throw new CliError("file-not-regular", "Input must be a regular local file.", 3);
    if (before.size > BigInt(limit)) throw new CliError("input-limit-exceeded", "Input exceeds max-input-bytes.", 4);
    const targetStat = async () => {
      try { return await io.statTarget(target); }
      catch (error) {
        if (["ENOENT", "ENOTDIR", "ELOOP"].includes((error as NodeJS.ErrnoException).code ?? "")) throw new CliError("file-changed", "Input changed while reading.", 3);
        throw error;
      }
    };
    if (!sameFile(before, await targetStat())) throw new CliError("file-changed", "Input changed while reading.", 3);
    const bytes = new Uint8Array(Number(before.size)); let offset = 0;
    while (offset < bytes.length) {
      checkAbort(signal); const read = await handle.read(bytes, offset, Math.min(64 * 1024, bytes.length - offset), offset);
      if (!read.bytesRead) throw new CliError("file-changed", "Input changed while reading.", 3); offset += read.bytesRead;
    }
    const extra = await handle.read(new Uint8Array(1), 0, 1, offset), after = await handle.stat({ bigint: true }); checkAbort(signal);
    if (extra.bytesRead || !sameFile(before, after) || !sameFile(after, await targetStat())) throw new CliError("file-changed", "Input changed while reading.", 3);
    checkAbort(signal);
    const input = { path: path.relative(cwd, target).split(path.sep).join("/"), byteLength: bytes.length, sha256: createHash("sha256").update(bytes).digest("hex") };
    return { bytes, input };
  } catch (error) { if (error instanceof CliError) throw error; throw new CliError("file-io", "Cannot read input file.", 3); }
  finally { await handle?.close(); }
}
