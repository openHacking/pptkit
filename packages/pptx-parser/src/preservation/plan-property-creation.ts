import { pathKey } from "../archive/read-archive.js";
import { readOpcSource, snapshotOptions } from "../opc/read-package.js";
import { PackageInspectionError } from "../types/archive.js";
import type { PptxInspectionOptions } from "../types/structure.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { appendPropertyDeclarations } from "./append-property-declarations.js";
import { inspectClosure } from "./inspect-closure.js";
import { XmlSourceEditError } from "./xml-text-source.js";

/** Prove that creating a new part cannot activate any pre-existing raw relationship.
 * This only plans OPC declarations: a caller must separately validate the final ZIP and projection. */
export async function planPropertyCreation(input: Uint8Array, kind: "core" | "custom" | "extended", dialect: "strict" | "transitional", options: PptxInspectionOptions, maxOutputBytes: number, maxFindings: number): Promise<{ part: string; relationshipId: string; relationships: Uint8Array; contentTypes: Uint8Array }> {
  if (!(input instanceof Uint8Array) || byteViewInfo(input).shared) throw new TypeError("Property creation requires an ordinary byte buffer.");
  if (kind !== "core" && kind !== "custom" && kind !== "extended") throw new TypeError("Invalid property creation kind.");
  if (dialect !== "strict" && dialect !== "transitional") throw new TypeError("Invalid property dialect.");
  if (!Number.isSafeInteger(maxOutputBytes) || maxOutputBytes <= 0 || !Number.isSafeInteger(maxFindings) || maxFindings <= 0) throw new RangeError("Invalid property creation limits.");
  const captured = snapshotOptions(options);
  const inspectionOptions = { limits: captured.limits, xmlLimits: captured.xmlLimits, zipLimits: captured.zipLimits, ...(captured.signal ? { signal: captured.signal } : {}) };
  if (byteViewInfo(input).length > captured.zipLimits.maxInputBytes) throw new PackageInspectionError("archive-limit-exceeded", "Property creation input exceeds maxInputBytes.");
  const bytes = copyByteView(input);
  // Findings are collected from every raw relationship set, even outside the
  // root closure. Missing targets and opaque markup make activation unsafe.
  const closure = await inspectClosure(bytes, { ...inspectionOptions, roots: ["/"], closureLimits: { maxFindings } });
  if (closure.findings.length) throw new XmlSourceEditError("xml-edit-unsupported", "Raw package relationships do not prove a closed source graph.");
  const original = await readOpcSource(bytes, inspectionOptions, false, true);
  const rel = original.payloads!.get("_rels/.rels"), types = original.payloads!.get("[Content_Types].xml");
  if (!rel || !types) throw new PackageInspectionError("package-invalid", "Root declarations are unavailable.");
  const paths = new Set(original.inspection.archive.entries.map(entry => pathKey(entry.name).split("/")[0]));
  const ids = new Set(closure.relationships.filter(edge => edge.source === "/").map(edge => edge.id));
  let serial = 1, part: string, relationshipId: string;
  do {
    if (captured.signal?.aborted) throw new PackageInspectionError("operation-aborted", "Property creation was aborted.");
    part = `/pptkit-${kind}-properties-${serial}.xml`; relationshipId = `pptkit${kind === "core" ? "Core" : kind === "custom" ? "Custom" : "Extended"}Properties${serial++}`;
    if (serial % 256 === 0) await new Promise<void>(resolve => setTimeout(resolve, 0));
  } while (ids.has(relationshipId) || paths.has(pathKey(part.slice(1))));
  const declarations = await appendPropertyDeclarations(rel, types, { part, relationshipId, dialect, kind }, { limits: captured.xmlLimits, ...(captured.signal ? { signal: captured.signal } : {}), maxOutputBytes });
  return { part, relationshipId, ...declarations };
}
