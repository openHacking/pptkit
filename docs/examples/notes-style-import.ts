import { importPptx } from "@pptkit/pptx-parser";
import { normalizePresentation } from "@pptkit/core";

export async function inspectInheritedNotes(bytes: Uint8Array) {
  const session = await importPptx(bytes, {
    projection: "content", notesTextProfile: "master-style-v1", textFontProfile: "script-families-v2",
    editSaveProfile: "text-runs-v3",
    sessionLimits: { maxTextStyleProperties: 100_000, maxTextStyleDependencies: 300_000, maxNotesStyleWork: 500_000 },
  });
  try {
    const document = normalizePresentation(session.document);
    const inherited = session.report.textStyleInheritance!.properties.filter(property => property.level.startsWith("notes-master"));
    const prepared = await session.prepareSave({ mode: "no-op" });
    if (prepared.status !== "ready") throw new Error("Expected unchanged notes source eligibility.");
    return { inherited, notes: document.slides.map(slide => slide.notes), original: (await session.save(prepared.plan)).bytes };
  } finally { session.dispose(); }
}
