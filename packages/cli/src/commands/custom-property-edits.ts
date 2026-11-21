import type { CustomDocumentProperty, PresentationMetadataInput } from "@pptkit/core";
import type { ImportedPresentation } from "@pptkit/pptx-parser";
import { CliError } from "./arguments.js";
import { checkAbort } from "./read-file.js";

const record = (v: unknown): v is Record<string, unknown> => typeof v === "object" && v !== null && !Array.isArray(v);
const valueText = (value: CustomDocumentProperty["value"]): string => typeof value === "number" && Object.is(value, -0) ? "-0" : String(value);
function targets(session: ImportedPresentation, signal?: AbortSignal, names = false) {
  const entries = new Map<string, { part: string; sourceObjectId: string; index: number; expected: { name: string; type: string; value: string } }>();
  const properties = session.document.metadata.customProperties ?? [];
  for (const item of session.report.customPropertySources ?? []) {
    checkAbort(signal);
    if (item.status !== "bound" || item.projectedIndex === undefined || names && item.nameEditable !== true) continue;
    const { part, sourceObjectId } = item.source, property = properties[item.projectedIndex];
    if (!sourceObjectId || !property || entries.has(sourceObjectId)) throw new CliError("custom-property-candidates-inconsistent", "Custom property bindings are ambiguous.", 3);
    entries.set(sourceObjectId, { part, sourceObjectId, index: item.projectedIndex, expected: { name: property.name, type: property.type, value: valueText(property.value) } });
  }
  return entries;
}
export function customPropertyCandidates(session: ImportedPresentation, sourceSha256: string, signal?: AbortSignal, names = false, removals = false, additions = false, create = false, currency = false) {
  const bindings = targets(session, signal, names && !removals && !additions);
  return { version: currency ? "pptkit-custom-property-edits-v6" : create ? "pptkit-custom-property-edits-v5" : additions ? "pptkit-custom-property-edits-v4" : removals ? "pptkit-custom-property-edits-v3" : names ? "pptkit-custom-property-edits-v2" : "pptkit-custom-property-edits-v1", sourceSha256, ...(additions ? { insertion: session.report.customPropertyInsertion } : {}), targets: [...bindings.values()].map(({ index: _index, ...rest }) => rest), unboundCount: (session.report.customPropertySources?.length ?? 0) - bindings.size };
}
export function parseCustomPropertyEdits(bytes: Uint8Array, sourceSha256: string): { version: "pptkit-custom-property-edits-v1" | "pptkit-custom-property-edits-v2" | "pptkit-custom-property-edits-v3" | "pptkit-custom-property-edits-v4" | "pptkit-custom-property-edits-v5" | "pptkit-custom-property-edits-v6"; edits: readonly unknown[]; additions: readonly unknown[] } {
  let value: unknown;
  try { value = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)); } catch { throw new CliError("custom-property-edits-invalid", "Custom property edits must be valid UTF-8 JSON.", 3); }
  if (!record(value) || (value.version !== "pptkit-custom-property-edits-v1" && value.version !== "pptkit-custom-property-edits-v2" && value.version !== "pptkit-custom-property-edits-v3" && value.version !== "pptkit-custom-property-edits-v4" && value.version !== "pptkit-custom-property-edits-v5" && value.version !== "pptkit-custom-property-edits-v6") || Object.keys(value).some(k => !["version", "sourceSha256", "edits", ...(["pptkit-custom-property-edits-v4", "pptkit-custom-property-edits-v5"].includes(String(value.version)) ? ["additions"] : [])].includes(k)) || typeof value.sourceSha256 !== "string" || !/^[0-9a-f]{64}$/.test(value.sourceSha256) || !Array.isArray(value.edits) || (value.additions !== undefined && !Array.isArray(value.additions)) || value.edits.length + (Array.isArray(value.additions) ? value.additions.length : 0) === 0 || value.edits.length + (Array.isArray(value.additions) ? value.additions.length : 0) > 10000) throw new CliError("custom-property-edits-invalid", "Expected a bounded custom property edit manifest.", 3);
  if (value.sourceSha256 !== sourceSha256) throw new CliError("custom-property-edits-source-mismatch", "Custom property edits do not match the source SHA-256.", 3);
  return { version: value.version, edits: value.edits, additions: Array.isArray(value.additions) ? value.additions : [] };
}
function convert(property: CustomDocumentProperty, text: string): CustomDocumentProperty {
  if (property.type === "string" || property.type === "date") return { ...property, value: text };
  if (property.type === "boolean" && (text === "true" || text === "false")) return { ...property, value: text === "true" };
  if (property.type === "integer" || property.type === "real") {
    const number = Number(text), canonical = Object.is(number, -0) ? "-0" : String(number);
    if (Number.isFinite(number) && canonical === text && (property.type !== "integer" || Number.isInteger(number) && number >= -2147483648 && number <= 2147483647 && !Object.is(number, -0))) return { ...property, value: number };
  }
  throw new CliError("custom-property-edits-invalid", "Replacement must use the canonical representation for its type.", 3);
}
export function applyCustomPropertyEdits(session: ImportedPresentation, plan: ReturnType<typeof parseCustomPropertyEdits>, signal?: AbortSignal): { removals: readonly string[]; additions: readonly number[] } {
  const bindings = targets(session, signal), seen = new Set<string>(), staged: PresentationMetadataInput = structuredClone(session.document.metadata);
  if (typeof session.document.setMetadata !== "function" || !staged.customProperties && !(plan.version === "pptkit-custom-property-edits-v5" && session.report.customPropertyInsertion?.status === "available")) throw new CliError("custom-property-edits-target-unbound", "Custom properties cannot be edited.", 3);
  const properties = [...(staged.customProperties ?? [])], removedIndexes = new Set<number>(), removalIds: string[] = [];
  for (const edit of plan.edits) {
    checkAbort(signal);
    if (!record(edit) || Object.keys(edit).some(k => !["part", "sourceObjectId", "expected", "replacement", ...(!["pptkit-custom-property-edits-v1", "pptkit-custom-property-edits-v6"].includes(plan.version) ? ["replacementName"] : []), ...(["pptkit-custom-property-edits-v3", "pptkit-custom-property-edits-v4", "pptkit-custom-property-edits-v5"].includes(plan.version) ? ["delete"] : [])].includes(k)) || typeof edit.part !== "string" || typeof edit.sourceObjectId !== "string" || !record(edit.expected) || Object.keys(edit.expected).some(k => !["name", "type", "value"].includes(k)) || typeof edit.expected.name !== "string" || typeof edit.expected.type !== "string" || typeof edit.expected.value !== "string" || (edit.replacement !== undefined && typeof edit.replacement !== "string") || (edit.replacementName !== undefined && (typeof edit.replacementName !== "string" || edit.replacementName.length === 0)) || (edit.delete !== undefined && edit.delete !== true) || (edit.delete === true && (edit.replacement !== undefined || edit.replacementName !== undefined)) || (["pptkit-custom-property-edits-v1", "pptkit-custom-property-edits-v6"].includes(plan.version) ? typeof edit.replacement !== "string" : edit.replacement === undefined && edit.replacementName === undefined && edit.delete !== true)) throw new CliError("custom-property-edits-invalid", "Invalid custom property edit.", 3);
    if (seen.has(edit.sourceObjectId)) throw new CliError("custom-property-edits-invalid", "Duplicate custom property identity.", 3); seen.add(edit.sourceObjectId);
    const binding = bindings.get(edit.sourceObjectId);
    if (!binding || binding.part !== edit.part) throw new CliError("custom-property-edits-target-unbound", "Custom property is not source-bound.", 3);
    if (binding.expected.name !== edit.expected.name || binding.expected.type !== edit.expected.type || binding.expected.value !== edit.expected.value) throw new CliError("custom-property-edits-baseline-mismatch", "Expected custom property differs from the imported baseline.", 3);
    if (edit.delete === true) { removedIndexes.add(binding.index); removalIds.push(binding.sourceObjectId); continue; }
    const property = properties[binding.index]!;
    const value = typeof edit.replacement === "string" ? convert(property, edit.replacement) : property;
    properties[binding.index] = typeof edit.replacementName === "string" ? { ...value, name: edit.replacementName } : value;
  }
  const remaining = properties.filter((_, index) => !removedIndexes.has(index)), additionIndexes: number[] = [];
  if (plan.additions.length && session.report.customPropertyInsertion?.status !== "available") throw new CliError("custom-property-edits-target-unbound", "A safe custom property insertion source is required.", 3);
  for (const addition of plan.additions) {
    checkAbort(signal);
    if (!record(addition) || Object.keys(addition).some(k => !["name", "type", "value"].includes(k)) || typeof addition.name !== "string" || !addition.name || typeof addition.value !== "string") throw new CliError("custom-property-edits-invalid", "Invalid custom property addition.", 3);
    const base: CustomDocumentProperty = addition.type === "string" || addition.type === "date" ? { name: addition.name, type: addition.type, value: "" } : addition.type === "boolean" ? { name: addition.name, type: "boolean", value: false } : addition.type === "integer" || addition.type === "real" ? { name: addition.name, type: addition.type, value: 0 } : (() => { throw new CliError("custom-property-edits-invalid", "Unknown custom property type.", 3); })();
    additionIndexes.push(remaining.length); remaining.push(convert(base, addition.value));
  }
  staged.customProperties = remaining; checkAbort(signal);
  try { session.document.setMetadata(staged); } catch (error) { if (error instanceof TypeError) throw new CliError("custom-property-edits-invalid", "Invalid custom property replacement or addition.", 3); throw error; }
  return { removals: removalIds, additions: additionIndexes };
}
