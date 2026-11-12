import { randomUUID } from "node:crypto";
import { link, open, unlink } from "node:fs/promises";
import path from "node:path";
import { CliError } from "./arguments.js";
import { checkAbort } from "./read-file.js";

/** Create-only publication of fully rendered preview text; commit is the hard link. */
export async function publishPreview(destination: string, content: string, signal?: AbortSignal): Promise<{ temporaryPath?: string }> {
  const temporary = path.join(path.dirname(destination), `.pptkit-svg-${randomUUID()}.tmp`);
  let handle: Awaited<ReturnType<typeof open>> | undefined, exists = false, committed = false;
  try {
    checkAbort(signal);
    handle = await open(temporary, "wx", 0o600); exists = true;
    await handle.writeFile(content, "utf8"); await handle.sync(); await handle.close(); handle = undefined;
    checkAbort(signal);
    await link(temporary, destination); committed = true;
    try { await unlink(temporary); exists = false; } catch { return { temporaryPath: temporary }; }
    return {};
  } catch (error) {
    if (error instanceof CliError) throw error;
    if (exists && !committed && (error as NodeJS.ErrnoException)?.code === "EEXIST") throw new CliError("file-target-exists", "Preview destination already exists.", 3);
    throw new CliError("file-io", "Cannot publish preview.", 3);
  } finally {
    if (handle) await handle.close().catch(() => {});
    if (exists && !committed) {
      try { await unlink(temporary); }
      catch { throw new CliError("file-cleanup-failed", "Preview temporary file could not be removed.", 3, { temporaryPath: temporary }); }
    }
  }
}
