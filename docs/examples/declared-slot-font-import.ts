import { importPptx, type PptxImportOptions, type PptxSourceLocation } from "@pptkit/pptx-parser";

export async function inspectDeclaredFontRules(bytes: Uint8Array) {
  const options: PptxImportOptions = {
    projection: "content",
    textFontProfile: "script-families-v3",
    notesTextProfile: "master-style-v1",
  };
  const session = await importPptx(bytes, options);
  try {
    const unresolved: { source: PptxSourceLocation; target: PptxSourceLocation; reason?: string }[] = [];
    for (const property of [
      ...(session.report.textStyleInheritance?.properties ?? []),
      ...(session.report.tableStyleInheritance?.properties ?? []),
    ]) {
      if (property.property.startsWith("run.scriptFontFamilies.") && property.status === "unsupported") {
        unresolved.push({ source: property.source, target: property.target, ...(property.reason ? { reason: property.reason } : {}) });
      }
    }
    return { profile: session.report.textFontProfile, unresolved };
  } finally { session.dispose(); }
}
