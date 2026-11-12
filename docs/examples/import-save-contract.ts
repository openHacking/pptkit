import type { NormalizedPresentation } from "@pptkit/core";
import type { ImportedPresentation, PptxSaveMode as SaveMode, PptxSaveReport as SaveReport, PptxSavePlan } from "@pptkit/pptx-parser";
export type { ImportedPresentation, PptxSaveMode as SaveMode, PptxSaveReport as SaveReport, PptxSavePlan, PptxSavePreparation as SavePreparation, ImportReport, ImportFinding, ImportDiagnostic, ImportSourceIdentity as SourceIdentity, PptxSourceLocation as SourceLocation, ImportOperation, ImportCapabilityTier as CapabilityTier, SaveLoss } from "@pptkit/pptx-parser";

// Actual experimental contracts: structural import and original-byte no-op saves are
// implemented; edit/rebuild preparation currently returns an explicit blocked report.
// Applications supply consent policy; no generic allowLossy boolean is used.
export async function reviewAndSave(
  session: ImportedPresentation,
  mode: SaveMode,
  decide: (report: SaveReport) => Promise<readonly string[]>,
) {
  const prepared = await session.prepareSave({ mode });
  if (prepared.status === "blocked") return prepared;
  try {
    const acceptLossIds = await decide(prepared.plan.report);
    return await session.save(prepared.plan, { acceptLossIds });
  } finally {
    prepared.plan.dispose();
  }
}

// Compile-only negative consumers verify the public contract boundaries.
function contractTypeChecks(session: ImportedPresentation, ir: NormalizedPresentation, report: SaveReport) {
  // @ts-expect-error Common IR cannot carry parser-owned source package state.
  ir.preservedParts;
  // @ts-expect-error The session's document identity is fixed.
  session.document = session.document;
  // @ts-expect-error Capability findings cannot be edited through the report.
  session.report.findings.push({});
  // @ts-expect-error Consumers cannot construct a usable preservation save plan.
  const forged: PptxSavePlan = { report, byteLength: 0, dispose() {} };
  // @ts-expect-error Consent must name the exact losses for this plan.
  session.save(forged, { allowLossy: true });
  // @ts-expect-error There is no ambiguous generic save mode.
  session.prepareSave({ mode: "save" });
}
void contractTypeChecks;
