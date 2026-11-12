import type { ImportCapabilityAssessment, ImportCapabilityOptions, ImportCapabilityReport, ImportCapabilitySubject } from "../types/capability.js";
import type { ImportFinding, ImportOperation, ImportReport } from "../types/import.js";
import { PptxImportError } from "../types/import.js";
import type { PptxSourceLocation } from "../types/source.js";
import { freezeReport } from "./state.js";
import { capabilityReferences } from "./capability-catalog.js";
const operations: readonly ImportOperation[] = ["create", "import", "edit", "preview", "export", "no-op-save", "edit-save"];
const tiers = ["editable", "preserved", "fallback", "unsupported"];
export const DEFAULT_IMPORT_CAPABILITY_LIMITS = Object.freeze({ maxFindings: 150000, maxTextCodeUnits: 16 * 1024 * 1024, maxWorkUnits: 32 * 1024 * 1024, maxSubjects: 1000 });
function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new TypeError("Expected a plain capability report record.");
  return value as Record<string, unknown>;
}
/** Read only recognized own data properties; ignore future fields without executing getters. */
function field(value: Record<string, unknown>, key: string): unknown {
  const descriptor = Object.getOwnPropertyDescriptor(value, key);
  if (!descriptor) return undefined;
  if (!("value" in descriptor)) throw new TypeError("Capability report fields cannot be accessors.");
  return descriptor.value;
}
/** Pure, bounded observation transport. No source/session/save authority is created. */
export function describeImportCapabilities(report: Pick<ImportReport, "projection" | "findings">, options: ImportCapabilityOptions = {}): ImportCapabilityReport {
  const config = record(options), limitsInput = field(config, "limits"), filterInput = field(config, "filter"), limits: Record<keyof typeof DEFAULT_IMPORT_CAPABILITY_LIMITS, number> = { ...DEFAULT_IMPORT_CAPABILITY_LIMITS };
  if (limitsInput !== undefined) {
    const input = record(limitsInput);
    if (Object.keys(input).some(k => !Object.hasOwn(limits, k))) throw new TypeError("Unknown capability report limit.");
    for (const key of Object.keys(limits) as (keyof typeof limits)[]) {
      const value = field(input, key); if (value === undefined) continue;
      if (!Number.isSafeInteger(value) || (value as number) < (key === "maxSubjects" ? 0 : 1)) throw new RangeError("Invalid capability report limit.");
      limits[key] = value as number;
    }
  }
  let work = 0, textUnits = 0;
  const step = (units = 1) => { work += units; if (work > limits.maxWorkUnits) throw new PptxImportError("import-limit-exceeded", "Capability report exceeds maxWorkUnits."); };
  const text = (value: unknown, optional = false): string | undefined => {
    if (value === undefined && optional) return undefined;
    if (typeof value !== "string") throw new TypeError("Expected a capability report string.");
    textUnits += value.length; step(Math.max(1, value.length));
    if (textUnits > limits.maxTextCodeUnits) throw new PptxImportError("import-limit-exceeded", "Capability report exceeds maxTextCodeUnits."); return value;
  };
  const filter: { feature?: string; part?: string; slideId?: string; elementId?: string } = {};
  if (filterInput !== undefined) for (const key of ["feature", "part", "slideId", "elementId"] as const) {
    const value = text(field(record(filterInput), key), true); if (value !== undefined) filter[key] = value;
  }
  const input = record(report), projection = field(input, "projection");
  if (projection !== "structure" && projection !== "content") throw new TypeError("Unknown capability report projection.");
  const findings = field(input, "findings");
  if (!Array.isArray(findings)) throw new TypeError("Expected capability findings.");
  if (findings.length > limits.maxFindings) throw new PptxImportError("import-limit-exceeded", "Capability report exceeds maxFindings.");
  type Group = { feature: string; source: PptxSourceLocation; slideId?: string; elementId?: string; entries: Map<ImportOperation, Pick<ImportFinding, "id" | "reason" | "tier">[]> };
  const groups = new Map<string, Group>(); let matchedFindings = 0;
  for (let index = 0; index < findings.length; index++) {
    step(); const descriptor = Object.getOwnPropertyDescriptor(findings, String(index));
    if (!descriptor || !("value" in descriptor)) throw new TypeError("Capability findings must be dense data arrays.");
    const f = record(descriptor.value), feature = text(field(f, "feature"))!, id = text(field(f, "id"))!, reason = text(field(f, "reason"))!;
    if (!feature.length || !id.length) throw new TypeError("Capability finding identities cannot be empty.");
    const operation = field(f, "operation") as ImportOperation, tier = field(f, "tier") as ImportFinding["tier"];
    if (!operations.includes(operation) || !tiers.includes(tier)) throw new TypeError("Unknown capability operation/tier.");
    const raw = record(field(f, "source")), part = text(field(raw, "part"))!;
    if (!part.startsWith("/")) throw new TypeError("Capability source part must be a logical absolute part.");
    const source: { part: string; path?: string; sourceObjectId?: string; relationshipId?: string; elementIndex?: number } = { part };
    for (const key of ["path", "sourceObjectId", "relationshipId"] as const) { const value = text(field(raw, key), true); if (value !== undefined) source[key] = value; }
    const elementIndex = field(raw, "elementIndex");
    if (elementIndex !== undefined) { if (!Number.isSafeInteger(elementIndex) || (elementIndex as number) < 1) throw new TypeError("Invalid raw source element index."); source.elementIndex = elementIndex as number; }
    const slideId = text(field(f, "slideId"), true), elementId = text(field(f, "elementId"), true);
    if (filter.feature !== undefined && feature !== filter.feature || filter.part !== undefined && part !== filter.part || filter.slideId !== undefined && slideId !== filter.slideId || filter.elementId !== undefined && elementId !== filter.elementId) continue;
    matchedFindings++;
    const key = JSON.stringify([feature, part, source.path ?? null, source.sourceObjectId ?? null, source.relationshipId ?? null, source.elementIndex ?? null, slideId ?? null, elementId ?? null]); step(key.length);
    let group = groups.get(key);
    if (!group) { group = { feature, source, ...(slideId !== undefined ? { slideId } : {}), ...(elementId !== undefined ? { elementId } : {}), entries: new Map() }; groups.set(key, group); }
    const entries = group.entries.get(operation) ?? []; entries.push({ id, reason, tier }); group.entries.set(operation, entries);
  }
  const compare = (a: string, b: string) => { step(1 + Math.min(a.length, b.length)); return a < b ? -1 : a > b ? 1 : 0; };
  const sorted = [...groups].sort(([a], [b]) => compare(a, b));
  const totals = operations.map(operation => ({ operation, reported: 0, unreported: 0, conflicting: 0 }));
  const subjects: ImportCapabilitySubject[] = []; let returnedFindings = 0;
  for (let index = 0; index < sorted.length; index++) {
    const [, group] = sorted[index]!; const assessments: ImportCapabilityAssessment[] = [];
    for (const [operationIndex, operation] of operations.entries()) {
      step(); const entries = group.entries.get(operation) ?? [], tier = entries[0]?.tier, consistent = entries.every(e => { step(); return e.tier === tier; });
      const status = !entries.length ? "unreported" : consistent ? "reported" : "conflicting"; totals[operationIndex]![status]++;
      if (index < limits.maxSubjects) {
        const observations = [...entries].sort((a, b) => compare(a.id, b.id) || compare(a.tier, b.tier) || compare(a.reason, b.reason));
        returnedFindings += observations.length;
        assessments.push({ operation, status, ...(status === "reported" ? { tier: tier! } : {}), observations });
      }
    }
    if (index < limits.maxSubjects) {
      const references = capabilityReferences(group.feature);
      subjects.push({ id: `subject-${index + 1}`, feature: group.feature, source: group.source, ...(group.slideId !== undefined ? { slideId: group.slideId } : {}), ...(group.elementId !== undefined ? { elementId: group.elementId } : {}), operations: assessments, ...(references ? { references } : {}) });
    }
  }
  return freezeReport({ schemaVersion: "pptkit-import-capabilities-v1", catalogVersion: "pptkit-capability-references-v1", coverage: "reported-source-findings", projection,
    nativeVerification: "not-performed", candidateSaveValidation: "not-performed", filter, totalFindings: findings.length, matchedFindings, totalSubjects: groups.size,
    returnedSubjects: subjects.length, returnedFindings, truncated: subjects.length < groups.size, totals, subjects, workUnits: work });
}
