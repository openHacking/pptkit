import type { OpcClosureInspection } from "../types/closure.js";
import { PptxImportError } from "../types/import.js";
import type { ChartSeriesOrderBinding } from "./chart-series-order.js";

/** Relationship graph proof only; arbitrary non-relationship references remain outside this check. */
export async function verifyChartOrderRelationshipClosure(binding: ChartSeriesOrderBinding, inspection: OpcClosureInspection, checkpoint: () => Promise<void>): Promise<void> {
  const fail = (reason: string): never => { throw new PptxImportError("save-candidate-invalid", reason, { part: binding.part }); };
  if (inspection.roots.length !== 1 || inspection.roots[0] !== binding.part || !inspection.coverage.incomingComplete || !inspection.coverage.outgoingComplete || inspection.findings.length) fail("Chart order relationship scan is incomplete.");
  if (inspection.incomingRelationships.length !== 1) fail("Chart order requires exactly one incoming relationship to its owned closure.");
  const edge = inspection.incomingRelationships[0]!;
  if (edge.source !== binding.slidePart || edge.targetPart !== binding.part || edge.targetMode !== "Internal" || edge.fragment !== undefined || !edge.matchesEffectiveRelationship ||
      !["http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "http://purl.oclc.org/ooxml/officeDocument/relationships/chart"].includes(edge.type ?? "")) fail("Chart order incoming relationship differs from its owning slide.");
  for (const relation of inspection.relationships) {
    await checkpoint();
    if (relation.source !== binding.part || relation.targetMode !== "Internal" || !relation.targetPart || relation.fragment !== undefined || !relation.matchesEffectiveRelationship ||
        !["http://schemas.openxmlformats.org/officeDocument/2006/relationships/package", "http://purl.oclc.org/ooxml/officeDocument/relationships/package"].includes(relation.type ?? "")) fail("Chart order closure contains an unsupported relationship.");
  }
  if (inspection.relationships.length > 1) fail("Chart order closure has multiple workbook relationships.");
}
