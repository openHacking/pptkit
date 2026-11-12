import { inspectZipArchive, PackageInspectionError } from "@pptkit/pptx-parser";

export async function inspectArchive(bytes: Uint8Array, signal?: AbortSignal) {
  try {
    return await inspectZipArchive(bytes, {
      limits: { maxInputBytes: 32 * 1024 * 1024, maxEntries: 10000 },
      ...(signal ? { signal } : {}),
    });
  } catch (error) {
    if (error instanceof PackageInspectionError) return { code: error.code, entry: error.entry, message: error.message };
    throw error;
  }
}
