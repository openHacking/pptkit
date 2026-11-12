import {
  importPptxFile,
  writePptxFile,
  type PptxFileSnapshot,
} from "@pptkit/pptx-parser/node";

export async function preserveImportedFile(
  sourcePath: string,
  destinationPath: string,
  expectedDestination?: PptxFileSnapshot,
) {
  const imported = await importPptxFile(sourcePath, {
    projection: "content",
    editSaveProfile: "text-runs-v3",
  });
  try {
    const prepared = await imported.session.prepareSave({ mode: "no-op" });
    if (prepared.status !== "ready") {
      throw new Error(prepared.report.diagnostics.map(item => item.message).join("\n"));
    }
    const result = await writePptxFile(
      imported.session,
      prepared.plan,
      destinationPath,
      expectedDestination
        ? { mode: "replace", expected: expectedDestination }
        : { mode: "create" },
    );
    return { result, source: imported.file };
  } finally {
    imported.session.dispose();
  }
}
