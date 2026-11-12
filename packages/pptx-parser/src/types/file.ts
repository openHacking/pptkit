import type { ImportedPresentation, PptxSaveReport } from "./import.js";

declare const fileSnapshotIdentity: unique symbol;
export interface PptxFileSnapshot {
  readonly [fileSnapshotIdentity]: true;
  readonly path: string;
  readonly byteLength: number;
  readonly sha256: string;
}
export interface ImportedPptxFile {
  readonly session: ImportedPresentation;
  readonly file: PptxFileSnapshot;
}
export type PptxFileWriteOptions = {
  readonly signal?: AbortSignal;
  readonly acceptLossIds?: readonly string[];
} & ({ readonly mode: "create" } | { readonly mode: "replace"; readonly expected: PptxFileSnapshot });
export interface PptxFileWriteWarning {
  readonly code: "temporary-cleanup-failed" | "directory-sync-failed" | "directory-close-failed";
  readonly path: string;
  readonly errorCode?: string;
  readonly message: string;
}
export interface PptxFileWriteResult {
  readonly status: "committed";
  readonly file: PptxFileSnapshot;
  readonly report: PptxSaveReport;
  /** File data was synced before publication; this reports the directory sync. */
  readonly directorySynced: boolean;
  readonly warnings: readonly PptxFileWriteWarning[];
}
export type PptxFileErrorCode = "file-io" | "file-not-regular" | "file-limit-exceeded" | "file-source-changed" | "file-target-changed" | "file-target-exists" | "file-write-busy" | "file-snapshot-invalid" | "file-platform-unsupported" | "file-cleanup-failed" | "operation-aborted";
export class PptxFileError extends Error {
  readonly committed = false;
  constructor(readonly code: PptxFileErrorCode, message: string, readonly path: string, options?: ErrorOptions & { temporaryPath?: string }) {
    super(message, options); this.name = "PptxFileError";
    if (options?.temporaryPath !== undefined) this.temporaryPath = options.temporaryPath;
  }
  readonly temporaryPath?: string;
}
