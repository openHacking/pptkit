import { constants, type BigIntStats } from "node:fs";
import { randomUUID } from "node:crypto";
import path from "node:path";
import type { FileHandle } from "node:fs/promises";
import type { FileIo } from "./file-io.js";
import { canonicalFile, checkAbort, checkPlatform, errorCode, fileHash, readStableFile, sameFile } from "./read-file.js";
import { captureImportOptions, importPresentation } from "../import/import-pptx.js";
import { PptxImportError, type ImportedPresentation, type PptxImportOptions, type PptxSavePlan } from "../types/import.js";
import { PptxFileError, type ImportedPptxFile, type PptxFileSnapshot, type PptxFileWriteOptions, type PptxFileWriteResult, type PptxFileWriteWarning } from "../types/file.js";

/** Private IO seam for deterministic filesystem-failure tests; not a public API. */
export function createPptxFileAdapter(io: FileIo) {
  const snapshots = new WeakMap<PptxFileSnapshot, BigIntStats>(), busy = new Set<string>();
  const snapshot = (target: string, bytes: number, sha256: string, stat: BigIntStats): PptxFileSnapshot => {
    const file = Object.freeze({ path: target, byteLength: bytes, sha256, toJSON() { throw new TypeError("PPTX file snapshots are process-local and not serializable."); } }) as unknown as PptxFileSnapshot;
    snapshots.set(file, stat); return file;
  };
  const ioError = (error: unknown, target: string): never => {
    if (error instanceof PptxFileError) throw error;
    throw new PptxFileError("file-io", "PPTX filesystem operation failed.", target, { cause: error });
  };
  return {
    async importFile(filePath: string, options: PptxImportOptions): Promise<ImportedPptxFile> {
      // Capture every caller option before the first filesystem await.
      const { mode: projection, profile: editSaveProfile, descriptionEditSaveProfile, drawingNameSaveProfile, decorativeSaveProfile, slideVisibilitySaveProfile, metadataSaveProfile, customPropertySaveProfile, customPropertySourceProfile, firstSlideNumberSaveProfile, masterArtwork, pictureSourceProfile, audioSourceProfile, audioProjection, pictureEditSaveProfile, fieldTextProfile, textFontProfile, notesTextProfile, textTargetSaveProfile, chartEditSaveProfile, diagramPreview, mediaPreview, diagramTextComparison, diagramLimits, captured, limits } = captureImportOptions(options);
      const config: PptxImportOptions = { projection, ...(editSaveProfile !== undefined ? { editSaveProfile } : {}), ...(descriptionEditSaveProfile ? { descriptionEditSaveProfile } : {}), ...(drawingNameSaveProfile ? { drawingNameSaveProfile } : {}), ...(decorativeSaveProfile ? { decorativeSaveProfile } : {}), ...(slideVisibilitySaveProfile ? { slideVisibilitySaveProfile } : {}), ...(metadataSaveProfile ? { metadataSaveProfile } : {}), ...(customPropertySourceProfile ? { customPropertySourceProfile } : {}), ...(customPropertySaveProfile ? { customPropertySaveProfile } : {}), ...(firstSlideNumberSaveProfile ? { firstSlideNumberSaveProfile } : {}), ...(masterArtwork ? { masterArtwork } : {}), ...(pictureSourceProfile ? { pictureSourceProfile } : {}), ...(audioSourceProfile ? { audioSourceProfile } : {}), ...(audioProjection ? { audioProjection } : {}), ...(pictureEditSaveProfile ? { pictureEditSaveProfile } : {}), ...(fieldTextProfile ? { fieldTextProfile } : {}), ...(textFontProfile ? { textFontProfile } : {}), ...(notesTextProfile ? { notesTextProfile } : {}), ...(textTargetSaveProfile ? { textTargetSaveProfile } : {}), ...(chartEditSaveProfile ? { chartEditSaveProfile } : {}), ...(diagramPreview ? { diagramPreview } : {}), ...(mediaPreview ? { mediaPreview } : {}), ...(diagramTextComparison ? { diagramTextComparison } : {}), diagramLimits, limits: captured.limits, xmlLimits: captured.xmlLimits, zipLimits: captured.zipLimits, ...(captured.signal ? { signal: captured.signal } : {}), sessionLimits: limits };
      checkPlatform(filePath); checkAbort(captured.signal, filePath);
      let target = filePath;
      let content: Awaited<ReturnType<typeof readStableFile>>;
      try {
        target = await canonicalFile(io, filePath);
        content = await readStableFile(io, target, captured.zipLimits.maxInputBytes, captured.signal);
      } catch (error) { return ioError(error, target); }
      const session = await importPresentation(content.bytes, config);
      return Object.freeze({ session, file: snapshot(target, content.bytes.length, content.sha256, content.stat) });
    },
    async writeFile(session: ImportedPresentation, plan: PptxSavePlan, filePath: string, options: PptxFileWriteOptions): Promise<PptxFileWriteResult> {
      if (!options || typeof options !== "object" || Array.isArray(options)) throw new TypeError("File write requires explicit create or replace options.");
      const mode = options.mode, signal = options.signal;
      if (mode !== "create" && mode !== "replace") throw new TypeError("File write requires explicit create or replace mode.");
      const allowed = mode === "create" ? ["mode", "signal", "acceptLossIds"] : ["mode", "signal", "acceptLossIds", "expected"];
      if (Object.keys(options).some(key => !allowed.includes(key))) throw new TypeError("Unknown PPTX file write option.");
      const accepted = options.acceptLossIds;
      const consent = accepted === undefined ? {} : { acceptLossIds: Array.isArray(accepted) ? [...accepted] : accepted };
      const expected = mode === "replace" ? options.expected : undefined, expectedStat = expected && snapshots.get(expected);
      if (mode === "replace" && !expectedStat) throw new PptxFileError("file-snapshot-invalid", "Replace requires an authentic snapshot from this adapter instance.", filePath);
      checkPlatform(filePath); checkAbort(signal, filePath);
      const saved = await session.save(plan, consent);
      const sha256 = fileHash(saved.bytes);
      checkAbort(signal, filePath);
      let target: string;
      try { target = await canonicalFile(io, filePath); } catch (error) { return ioError(error, filePath); }
      if (expected && expected.path !== target) throw new PptxFileError("file-snapshot-invalid", "Expected snapshot belongs to a different canonical file path.", target);
      const keys = [target, ...(expectedStat ? [`inode:${expectedStat.dev}:${expectedStat.ino}`] : [])];
      if (keys.some(key => busy.has(key))) throw new PptxFileError("file-write-busy", "A write to this path or source inode is already in progress.", target);
      for (const key of keys) busy.add(key);
      const temporary = path.join(path.dirname(target), `.pptkit-${randomUUID()}.tmp`), warnings: PptxFileWriteWarning[] = [];
      let temp: FileHandle | undefined, parent: FileHandle | undefined, tempExists = false, committed = false;
      const warn = (code: PptxFileWriteWarning["code"], at: string, error: unknown) => warnings.push(Object.freeze({ code, path: at, ...(errorCode(error) ? { errorCode: errorCode(error)! } : {}), message: error instanceof Error ? error.message : String(error) }));
      const guardTarget = async () => {
        if (!expected || !expectedStat) {
          try { await io.lstat(target); } catch (error) { if (errorCode(error) === "ENOENT") return; throw error; }
          throw new PptxFileError("file-target-exists", "Create mode never replaces an existing name.", target);
        }
        try {
          const named = await io.lstat(target, { bigint: true });
          if (!named.isFile() || !sameFile(named, expectedStat)) throw new PptxFileError("file-target-changed", "Replace target identity/metadata differs from the expected snapshot.", target);
          const current = await readStableFile(io, target, expected.byteLength, signal);
          if (!sameFile(current.stat, expectedStat) || current.sha256 !== expected.sha256) throw new PptxFileError("file-target-changed", "Replace target bytes differ from the expected snapshot.", target);
        } catch (error) {
          if (error instanceof PptxFileError && error.code === "operation-aborted") throw error;
          if (errorCode(error) === "ENOENT" || errorCode(error) === "ELOOP" || error instanceof PptxFileError) throw new PptxFileError("file-target-changed", "Replace target disappeared or changed during verification.", target, { cause: error });
          throw error;
        }
      };
      try {
        checkAbort(signal, target);
        parent = await io.open(path.dirname(target), constants.O_RDONLY | constants.O_NOFOLLOW | constants.O_DIRECTORY);
        const parentStat = await parent.stat({ bigint: true });
        if (!parentStat.isDirectory()) throw new PptxFileError("file-not-regular", "Output parent must be an existing directory.", target);
        await guardTarget(); checkAbort(signal, target);
        temp = await io.open(temporary, constants.O_WRONLY | constants.O_CREAT | constants.O_EXCL | constants.O_NOFOLLOW, 0o600); tempExists = true;
        await temp.chmod(expectedStat ? Number(expectedStat.mode & 0o777n) : 0o600);
        for (let offset = 0; offset < saved.bytes.length;) {
          checkAbort(signal, target);
          const written = await temp.write(saved.bytes, offset, Math.min(64 * 1024, saved.bytes.length - offset), offset);
          if (!written.bytesWritten) throw new PptxFileError("file-io", "Temporary PPTX write made no progress.", target);
          offset += written.bytesWritten;
        }
        checkAbort(signal, target); await temp.sync(); checkAbort(signal, target);
        const tempStat = await temp.stat({ bigint: true });
        if (!tempStat.isFile() || tempStat.size !== BigInt(saved.bytes.length)) throw new PptxFileError("file-io", "Temporary file size differs from the prepared bytes.", target);
        await temp.close(); temp = undefined;
        const verified = await readStableFile(io, temporary, saved.bytes.length, signal);
        if (!sameFile(tempStat, verified.stat, true) || verified.sha256 !== sha256) throw new PptxFileError("file-io", "Flushed temporary file differs from the prepared candidate.", target);
        const namedParent = await io.lstat(path.dirname(target), { bigint: true });
        if (!namedParent.isDirectory() || namedParent.dev !== parentStat.dev || namedParent.ino !== parentStat.ino) throw new PptxFileError("file-target-changed", "Output parent directory identity changed.", target);
        await guardTarget(); checkAbort(signal, target);
        const namedTemporary = await io.lstat(temporary, { bigint: true });
        checkAbort(signal, target);
        if (!namedTemporary.isFile() || !sameFile(tempStat, namedTemporary, true)) throw new PptxFileError("file-target-changed", "Temporary candidate name changed before commit.", target);
        const receipt = snapshot(target, saved.bytes.length, sha256, tempStat);
        // This guard is synchronous: no JS await separates freshness from syscall dispatch.
        session.assertSavePlanCurrent(plan, consent);
        if (mode === "create") {
          try { await io.link(temporary, target); }
          catch (error) { if (errorCode(error) === "EEXIST") throw new PptxFileError("file-target-exists", "Create lost a race to an existing destination; it was not replaced.", target, { cause: error }); throw error; }
        } else await io.rename(temporary, target);
        committed = true;
        if (mode === "replace") tempExists = false;
        if (tempExists) {
          try { await io.unlink(temporary); tempExists = false; } catch (error) { warn("temporary-cleanup-failed", temporary, error); }
        }
        let directorySynced = false;
        try { await parent.sync(); directorySynced = true; } catch (error) { warn("directory-sync-failed", path.dirname(target), error); }
        try { await parent.close(); } catch (error) {
          warn("directory-close-failed", path.dirname(target), error);
          try { await parent.close(); } catch (failure) { warn("directory-close-failed", path.dirname(target), failure); }
        }
        parent = undefined;
        return Object.freeze({ status: "committed", file: receipt, report: saved.report, directorySynced, warnings: Object.freeze(warnings) });
      } catch (error) {
        if (committed) throw error;
        let cleanupFailure: unknown;
        if (temp) { try { await temp.close(); } catch (failure) { cleanupFailure = failure; } }
        if (tempExists) { try { await io.unlink(temporary); tempExists = false; } catch (failure) { cleanupFailure = failure; } }
        if (parent) { try { await parent.close(); } catch (failure) { cleanupFailure = failure; } parent = undefined; }
        if (cleanupFailure) throw new PptxFileError("file-cleanup-failed", "No output was committed; temporary cleanup/handle close failed.", target, { cause: new AggregateError([error, cleanupFailure]), ...(tempExists ? { temporaryPath: temporary } : {}) });
        // Preserve parser plan/lifecycle errors; wrap filesystem errors only.
        if (error instanceof PptxImportError) throw error;
        return ioError(error, target);
      } finally { for (const key of keys) busy.delete(key); }
    },
  };
}
