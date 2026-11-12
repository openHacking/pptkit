import { importPptx } from "@pptkit/pptx-parser";
import { generatePptx } from "@pptkit/pptx-exporter";

/** Caller explicitly requests a fresh projection, after reviewing its limitations. */
export async function editProjectedNotes(bytes: Uint8Array) {
  const session = await importPptx(bytes, { projection: "content" });
  try {
    const noteSources = session.report.identities.filter(identity => identity.kind === "notes");
    const limitations = session.report.findings.filter(finding =>
      finding.operation === "import" && finding.tier !== "editable");
    const first = session.document.slides[0];
    if (first) first.setNotes([
      { runs: [{ text: "Updated speaker notes", style: { bold: true } }] },
      { runs: [{ text: "Continue", action: { type: "slide", slideId: first.id } }] },
    ]);
    // Creates a new PPTX containing the mapped projection. Source edit-save
    // remains blocked; this does not preserve unknown source parts or notes layout.
    const output = await generatePptx(session.document);
    return { output, noteSources, limitations };
  } finally {
    session.dispose();
  }
}
