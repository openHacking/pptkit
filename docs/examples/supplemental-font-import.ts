import { importPptx } from "@pptkit/pptx-parser";
import { normalizePresentation, resolveTextFontSegments } from "@pptkit/core";

export async function inspectSupplementalFonts(bytes: Uint8Array) {
  const session = await importPptx(bytes, {
    projection: "content", textFontProfile: "script-families-v2", editSaveProfile: "text-runs-v3",
    sessionLimits: { maxTextStyleProperties: 100_000, maxTextStyleDependencies: 300_000 },
  });
  try {
    const document = normalizePresentation(session.document);
    const rules = [...session.report.textStyleInheritance!.properties, ...session.report.tableStyleInheritance!.properties]
      .filter(p => p.property.startsWith("run.scriptFontFamilies."));
    const fragments = document.slides.flatMap(slide => slide.elements.flatMap(element => element.type === "text"
      ? element.content.flatMap(p => p.runs.flatMap(run => resolveTextFontSegments(run.text, run.style, document.theme))) : []));
    const prepared = await session.prepareSave({ mode: "no-op" });
    if (prepared.status !== "ready") throw new Error("Unchanged source must remain eligible.");
    return { rules, fragments, original: (await session.save(prepared.plan)).bytes };
  } finally { session.dispose(); }
}
