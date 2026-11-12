import type { ImportedPresentation } from "@pptkit/pptx-parser";
import type { PresentationMetadataInput } from "@pptkit/core";
import { CliError } from "./arguments.js";
import { checkAbort } from "./read-file.js";

const fields = ["title", "author", "subject", "description", "language"] as const;
type StringField = typeof fields[number];
type Field = StringField | "revision" | "keywords" | "company" | "createdAt" | "modifiedAt";
type Version = "pptkit-metadata-edits-v1" | "pptkit-metadata-edits-v2" | "pptkit-metadata-edits-v3" | "pptkit-metadata-edits-v4" | "pptkit-metadata-edits-v5" | "pptkit-metadata-edits-v6" | "pptkit-metadata-edits-v7" | "pptkit-metadata-edits-v8" | "pptkit-metadata-edits-v9" | "pptkit-metadata-edits-v10" | "pptkit-metadata-edits-v11" | "pptkit-metadata-edits-v12" | "pptkit-metadata-edits-v13" | "pptkit-metadata-edits-v14" | "pptkit-metadata-edits-v15" | "pptkit-metadata-edits-v16" | "pptkit-metadata-edits-v17";
const record = (value: unknown): value is Record<string, unknown> => typeof value === "object" && value !== null && !Array.isArray(value);
const isStringField = (value: unknown): value is StringField => typeof value === "string" && (fields as readonly string[]).includes(value);
const isDateField = (value: unknown): value is "createdAt" | "modifiedAt" => value === "createdAt" || value === "modifiedAt";
const isField = (value: unknown, revision: boolean, keywords: boolean, company: boolean, dates: boolean): value is Field => isStringField(value) || revision && value === "revision" || keywords && value === "keywords" || company && value === "company" || dates && (value === "createdAt" || value === "modifiedAt");
const isDate = (value: unknown): value is string => typeof value === "string" && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(value) && Number.isFinite(Date.parse(value)) && new Date(value).toISOString() === value;
const isRevision = (value: unknown): value is number => typeof value === "number" && Number.isSafeInteger(value) && value >= 0 && !Object.is(value, -0);
function sources(session: ImportedPresentation, revision: boolean, keywords: boolean, company: boolean, dates: boolean, signal?: AbortSignal): Map<Field, "bound" | "insertable" | "unmapped"> {
  const result = new Map<Field, "bound" | "insertable" | "unmapped">();
  for (const binding of session.report.coreMetadataSources ?? []) {
    checkAbort(signal); if (!isField(binding.field, revision, keywords, company, dates) || result.has(binding.field)) throw new CliError("metadata-candidates-inconsistent", "Core metadata field identity is ambiguous or unknown.", 3);
    result.set(binding.field, binding.status);
  }
  return result;
}
export function metadataCandidates(session: ImportedPresentation, sourceSha256: string, signal?: AbortSignal, revision = false, keywords = false, company = false, insert = false, dates = false, dateInsert = false, create = false, revisionInsert = false, keywordInsert = false, companyInsert = false, extendedCreate = false, remove = false, dateRemove = false, keywordsRemove = false, revisionRemove = false, defaultRemove = false) {
  revisionRemove ||= defaultRemove;
  keywordsRemove ||= revisionRemove;
  dateRemove ||= keywordsRemove;
  remove ||= dateRemove;
  extendedCreate ||= remove;
  companyInsert ||= extendedCreate;
  const bindings = sources(session, revision || keywords || company || insert || dates || dateInsert || create || revisionInsert || keywordInsert || companyInsert, keywords || company || insert || dates || dateInsert || create || revisionInsert || keywordInsert || companyInsert, company || insert || dates || dateInsert || create || revisionInsert || keywordInsert || companyInsert, dates || dateInsert || create || revisionInsert || keywordInsert || companyInsert, signal), targets: { field: Field; expected: string | number | null }[] = []; let unboundCount = 0;
  for (const [field, status] of bindings) {
    checkAbort(signal); if ((insert || dates || dateInsert || create || revisionInsert || keywordInsert || companyInsert) && status === "insertable" && (isStringField(field) || (dateInsert || create || revisionInsert || keywordInsert || companyInsert) && isDateField(field) || (revisionInsert || keywordInsert || companyInsert) && field === "revision" || (keywordInsert || companyInsert) && field === "keywords" || companyInsert && field === "company")) { targets.push({ field, expected: null }); continue; }
    if (status !== "bound") { unboundCount++; continue; }
    const expected = field === "keywords" ? session.document.metadata.keywords?.length === 1 ? session.document.metadata.keywords[0] : undefined : session.document.metadata[field];
    if ((field === "revision" ? !isRevision(expected) : field === "createdAt" || field === "modifiedAt" ? !isDate(expected) : typeof expected !== "string") || typeof session.document.setMetadata !== "function") throw new CliError("metadata-candidates-inconsistent", "Bound core metadata is unavailable.", 3);
    targets.push({ field, expected: expected as string | number });
  }
  return { version: defaultRemove ? "pptkit-metadata-edits-v17" : revisionRemove ? "pptkit-metadata-edits-v16" : keywordsRemove ? "pptkit-metadata-edits-v15" : dateRemove ? "pptkit-metadata-edits-v14" : remove ? "pptkit-metadata-edits-v13" : extendedCreate ? "pptkit-metadata-edits-v12" : companyInsert ? "pptkit-metadata-edits-v11" : keywordInsert ? "pptkit-metadata-edits-v10" : revisionInsert ? "pptkit-metadata-edits-v9" : create ? "pptkit-metadata-edits-v8" : dateInsert ? "pptkit-metadata-edits-v7" : dates ? "pptkit-metadata-edits-v6" : insert ? "pptkit-metadata-edits-v5" : company ? "pptkit-metadata-edits-v4" : keywords ? "pptkit-metadata-edits-v3" : revision ? "pptkit-metadata-edits-v2" : "pptkit-metadata-edits-v1", sourceSha256, targets, unboundCount, ...(remove ? { removableFields: (session.report.coreMetadataSources ?? []).filter(source => source.removable === true).map(source => source.field) } : {}) };
}
export function parseMetadataEdits(bytes: Uint8Array, sourceSha256: string): { version: Version; edits: readonly unknown[] } {
  let value: unknown; try { value = JSON.parse(new TextDecoder("utf-8", { fatal: true }).decode(bytes)); }
  catch { throw new CliError("metadata-edits-invalid", "Metadata edits must be valid UTF-8 JSON.", 3); }
  if (!record(value) || value.version !== "pptkit-metadata-edits-v1" && value.version !== "pptkit-metadata-edits-v2" && value.version !== "pptkit-metadata-edits-v3" && value.version !== "pptkit-metadata-edits-v4" && value.version !== "pptkit-metadata-edits-v5" && value.version !== "pptkit-metadata-edits-v6" && value.version !== "pptkit-metadata-edits-v7" && value.version !== "pptkit-metadata-edits-v8" && value.version !== "pptkit-metadata-edits-v9" && value.version !== "pptkit-metadata-edits-v10" && value.version !== "pptkit-metadata-edits-v11" && value.version !== "pptkit-metadata-edits-v12" && value.version !== "pptkit-metadata-edits-v13" && value.version !== "pptkit-metadata-edits-v14" && value.version !== "pptkit-metadata-edits-v15" && value.version !== "pptkit-metadata-edits-v16" && value.version !== "pptkit-metadata-edits-v17" || Object.keys(value).some(k => !["version", "sourceSha256", "edits"].includes(k)) || typeof value.sourceSha256 !== "string" || !/^[0-9a-f]{64}$/.test(value.sourceSha256) || !Array.isArray(value.edits) || !value.edits.length || value.edits.length > ((value.version === "pptkit-metadata-edits-v6" || value.version === "pptkit-metadata-edits-v7" || value.version === "pptkit-metadata-edits-v8" || value.version === "pptkit-metadata-edits-v9" || value.version === "pptkit-metadata-edits-v10" || value.version === "pptkit-metadata-edits-v11" || value.version === "pptkit-metadata-edits-v12" || value.version === "pptkit-metadata-edits-v13" || value.version === "pptkit-metadata-edits-v14" || value.version === "pptkit-metadata-edits-v15" || value.version === "pptkit-metadata-edits-v16" || value.version === "pptkit-metadata-edits-v17") ? fields.length + 5 : (value.version === "pptkit-metadata-edits-v5" || value.version === "pptkit-metadata-edits-v4") ? fields.length + 3 : value.version === "pptkit-metadata-edits-v3" ? fields.length + 2 : value.version === "pptkit-metadata-edits-v2" ? fields.length + 1 : fields.length)) throw new CliError("metadata-edits-invalid", "Expected a bounded metadata edit manifest.", 3);
  if (value.sourceSha256 !== sourceSha256) throw new CliError("metadata-edits-source-mismatch", "Metadata edits do not match the captured source SHA-256.", 3);
  return { version: value.version, edits: value.edits };
}
export function applyMetadataEdits(session: ImportedPresentation, plan: ReturnType<typeof parseMetadataEdits>, signal?: AbortSignal): { insertions: readonly (StringField | "revision")[]; removals: readonly ("title" | "author" | "language" | "revision")[] } {
  const companyInsert = plan.version === "pptkit-metadata-edits-v11" || plan.version === "pptkit-metadata-edits-v12" || plan.version === "pptkit-metadata-edits-v13" || plan.version === "pptkit-metadata-edits-v14" || plan.version === "pptkit-metadata-edits-v15" || plan.version === "pptkit-metadata-edits-v16" || plan.version === "pptkit-metadata-edits-v17", keywordInsert = companyInsert || plan.version === "pptkit-metadata-edits-v10", revisionInsert = keywordInsert || plan.version === "pptkit-metadata-edits-v9", dateInsert = revisionInsert || plan.version === "pptkit-metadata-edits-v7" || plan.version === "pptkit-metadata-edits-v8", dates = dateInsert || plan.version === "pptkit-metadata-edits-v6", insert = dates || plan.version === "pptkit-metadata-edits-v5", company = insert || plan.version === "pptkit-metadata-edits-v4", keywords = company || plan.version === "pptkit-metadata-edits-v3", revision = plan.version !== "pptkit-metadata-edits-v1", bindings = sources(session, revision, keywords, company, dates, signal), seen = new Set<Field>(), insertions: (StringField | "revision")[] = [], removals: ("title" | "author" | "language" | "revision")[] = [], staged: PresentationMetadataInput = structuredClone(session.document.metadata);
  if (typeof session.document.setMetadata !== "function") throw new CliError("metadata-edits-target-unbound", "Document metadata cannot be edited.", 3);
  for (const edit of plan.edits) {
    checkAbort(signal);
    if (!record(edit) || Object.keys(edit).some(k => !["field", "expected", "replacement"].includes(k)) || !isField(edit.field, revision, keywords, company, dates) || (edit.field === "revision" ? !(revisionInsert && edit.expected === null) && !isRevision(edit.expected) || !(isRevision(edit.replacement) || (plan.version === "pptkit-metadata-edits-v16" || plan.version === "pptkit-metadata-edits-v17") && edit.replacement === null) : (typeof edit.expected !== "string" && !(insert && (isStringField(edit.field) || dateInsert && isDateField(edit.field) || keywordInsert && edit.field === "keywords" || companyInsert && edit.field === "company") && edit.expected === null)) || !(typeof edit.replacement === "string" || (plan.version === "pptkit-metadata-edits-v13" || plan.version === "pptkit-metadata-edits-v14" || plan.version === "pptkit-metadata-edits-v15" || plan.version === "pptkit-metadata-edits-v16" || plan.version === "pptkit-metadata-edits-v17") && edit.replacement === null && (edit.field === "subject" || edit.field === "description" || edit.field === "company" || (plan.version === "pptkit-metadata-edits-v14" || plan.version === "pptkit-metadata-edits-v15" || plan.version === "pptkit-metadata-edits-v16" || plan.version === "pptkit-metadata-edits-v17") && isDateField(edit.field) || (plan.version === "pptkit-metadata-edits-v15" || plan.version === "pptkit-metadata-edits-v16" || plan.version === "pptkit-metadata-edits-v17") && edit.field === "keywords" || plan.version === "pptkit-metadata-edits-v17" && (edit.field === "title" || edit.field === "author" || edit.field === "language"))))) throw new CliError("metadata-edits-invalid", "Metadata edits require a supported field and correctly typed values.", 3);
    if (isDateField(edit.field) && (!(dateInsert && edit.expected === null) && !isDate(edit.expected) || edit.replacement !== null && !isDate(edit.replacement))) throw new CliError("metadata-edits-invalid", "Date edits require canonical UTC timestamps with millisecond precision.", 3);
    if (seen.has(edit.field)) throw new CliError("metadata-edits-invalid", "Duplicate metadata field in manifest.", 3); seen.add(edit.field);
    const status = bindings.get(edit.field);
    if (edit.replacement === null && edit.field === "revision") {
      const source = session.report.coreMetadataSources?.find(item => item.field === "revision");
      if (status !== "bound" || !source?.removable || !isRevision(edit.expected) || session.document.metadata.revision !== edit.expected) throw new CliError("metadata-edits-target-unbound", "Revision removal requires an eligible source and matching baseline.", 3);
      delete staged.revision; removals.push("revision"); continue;
    }
    if (edit.replacement === null && (edit.field === "title" || edit.field === "author" || edit.field === "language")) {
      const source = session.report.coreMetadataSources?.find(item => item.field === edit.field);
      if (status !== "bound" || !source?.removable || typeof edit.expected !== "string" || session.document.metadata[edit.field] !== edit.expected) throw new CliError("metadata-edits-target-unbound", "Metadata removal requires an eligible source and matching baseline.", 3);
      delete staged[edit.field]; removals.push(edit.field); continue;
    }
    if (edit.replacement === null) {
      const source = session.report.coreMetadataSources?.find(item => item.field === edit.field);
      const actual = edit.field === "keywords" ? session.document.metadata.keywords?.length === 1 ? session.document.metadata.keywords[0] : undefined : session.document.metadata[edit.field as "subject" | "description" | "company" | "createdAt" | "modifiedAt"];
      if (status !== "bound" || !source?.removable || typeof edit.expected !== "string" || actual !== edit.expected) throw new CliError("metadata-edits-target-unbound", "Metadata removal requires an eligible source and matching baseline.", 3);
      if (edit.field === "keywords") staged.keywords = []; else delete staged[edit.field as "subject" | "description" | "company" | "createdAt" | "modifiedAt"]; continue;
    }
    if (status === "insertable" && insert && isStringField(edit.field) && edit.expected === null) { staged[edit.field] = edit.replacement as string; insertions.push(edit.field); continue; }
    if (status === "insertable" && dateInsert && isDateField(edit.field) && edit.expected === null) { staged[edit.field] = edit.replacement as string; continue; }
    if (status === "insertable" && revisionInsert && edit.field === "revision" && edit.expected === null) { staged.revision = edit.replacement as number; insertions.push("revision"); continue; }
    if (status === "insertable" && keywordInsert && edit.field === "keywords" && edit.expected === null) { staged.keywords = [edit.replacement as string]; continue; }
    if (status === "insertable" && companyInsert && edit.field === "company" && edit.expected === null) { staged.company = edit.replacement as string; continue; }
    if (status !== "bound") throw new CliError("metadata-edits-target-unbound", "Metadata field is missing or not source-bound.", 3);
    const actual = edit.field === "keywords" ? session.document.metadata.keywords?.length === 1 ? session.document.metadata.keywords[0] : undefined : session.document.metadata[edit.field];
    if (actual !== edit.expected) throw new CliError("metadata-edits-baseline-mismatch", "Expected metadata does not match the imported baseline.", 3);
    if (edit.field === "revision") staged.revision = edit.replacement as number;
    else if (edit.field === "keywords") staged.keywords = [edit.replacement as string];
    else staged[edit.field] = edit.replacement as string;
  }
  checkAbort(signal); session.document.setMetadata(staged); return { insertions, removals };
}
