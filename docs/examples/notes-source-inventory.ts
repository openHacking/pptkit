import { inspectPptxNotes, type PptxNotesInspection, type PptxNotesRole } from "@pptkit/pptx-parser";

export async function notesSources(bytes: Uint8Array, signal?: AbortSignal): Promise<PptxNotesInspection> {
  const report = await inspectPptxNotes(bytes, { ...(signal ? { signal } : {}), notesLimits: { maxRecords: 100_000 } });
  const roles: readonly PptxNotesRole[] = report.parts.flatMap(part => part.drawings.map(drawing => drawing.role));
  void roles;
  return report;
}
