import { importPptx } from "@pptkit/pptx-parser";
import { normalizePresentation } from "@pptkit/core";

export async function inspectSourceScriptFonts(bytes: Uint8Array) {
  const session = await importPptx(bytes, {
    projection: "content",
    textFontProfile: "script-families-v1",
    editSaveProfile: "text-runs-v3",
    sessionLimits: { maxTextStyleProperties: 100_000, maxTextStyleDependencies: 200_000 },
  });
  try {
    const fields = new Set(["run.fontFamily", "run.eastAsianFontFamily", "run.complexScriptFontFamily"]);
    const fonts = [
      ...(session.report.textStyleInheritance?.properties ?? []),
      ...(session.report.tableStyleInheritance?.properties ?? []),
    ].filter(property => fields.has(property.property));
    // Mapped values establish the named projection, not source font-write permission.
    const blocked = fonts.filter(property => property.status === "unsupported");
    const ready = await session.prepareSave({ mode: "no-op" });
    if (ready.status !== "ready") throw new Error("Expected unchanged source eligibility.");
    const saved = await session.save(ready.plan);
    return { profile: session.report.textFontProfile, fonts, blocked, document: normalizePresentation(session.document), original: saved.bytes };
  } finally {
    session.dispose();
  }
}
