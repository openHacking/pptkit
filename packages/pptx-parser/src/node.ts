import { createPptxFileAdapter } from "./node/file-adapter.js";
import { nodeFileIo } from "./node/file-io.js";
import type { ImportedPresentation, PptxImportOptions, PptxSavePlan } from "./types/import.js";
import type { ImportedPptxFile, PptxFileWriteOptions, PptxFileWriteResult } from "./types/file.js";

export { PptxFileError } from "./types/file.js";
export type { PptxFileSnapshot, ImportedPptxFile, PptxFileWriteOptions, PptxFileWriteResult, PptxFileWriteWarning, PptxFileErrorCode } from "./types/file.js";
const adapter = createPptxFileAdapter(nodeFileIo);
export async function importPptxFile(filePath: string, options: PptxImportOptions): Promise<ImportedPptxFile> {
  return adapter.importFile(filePath, options);
}
export async function writePptxFile(session: ImportedPresentation, plan: PptxSavePlan, filePath: string, options: PptxFileWriteOptions): Promise<PptxFileWriteResult> {
  return adapter.writeFile(session, plan, filePath, options);
}
