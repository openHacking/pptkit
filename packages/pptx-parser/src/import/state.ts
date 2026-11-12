import { normalizePresentation, type NormalizedPresentation, type PresentationDocument } from "@pptkit/core";
import { PptxImportError } from "../types/import.js";

/** Iterative equality over detached normalized JSON-like state; no hash-only freshness. */
export function equalState(a: unknown, b: unknown, maxValues: number): boolean {
  const stack: [unknown, unknown][] = [[a, b]], compared = new WeakMap<object, WeakSet<object>>();
  let values = 1;
  while (stack.length) {
    if (values > maxValues) throw new PptxImportError("save-limit-exceeded", "Normalized comparison exceeds maxSnapshotValues.");
    const [left, right] = stack.pop()!;
    if (Object.is(left, right)) continue;
    if (!left || !right || typeof left !== "object" || typeof right !== "object" || Array.isArray(left) !== Array.isArray(right)) return false;
    let rights = compared.get(left);
    if (rights?.has(right)) continue;
    if (!rights) { rights = new WeakSet(); compared.set(left, rights); } rights.add(right);
    const l = left as Record<string, unknown>, r = right as Record<string, unknown>, keys = Object.keys(l);
    if (keys.length !== Object.keys(r).length || (Array.isArray(left) && left.length !== (right as unknown[]).length)) return false;
    if (keys.length > maxValues - values) throw new PptxImportError("save-limit-exceeded", "Normalized comparison exceeds maxSnapshotValues.");
    values += keys.length;
    for (const key of keys) { if (!Object.hasOwn(r, key)) return false; stack.push([l[key], r[key]]); }
  }
  return true;
}
/** Only called for owned report records, never for authoring state or caller objects. */
export function freezeReport<T>(report: T): T {
  const stack: unknown[] = [report], seen = new Set<object>();
  while (stack.length) {
    const value = stack.pop();
    if (!value || typeof value !== "object" || seen.has(value)) continue;
    seen.add(value); for (const child of Object.values(value)) stack.push(child); Object.freeze(value);
  }
  return report;
}
export function snapshot(document: PresentationDocument, maxValues: number): NormalizedPresentation {
  try {
    const result = normalizePresentation(document);
    const active = new Set<object>(), complete = new Set<object>(), stack: { value: unknown; exit?: boolean }[] = [{ value: result }];
    let values = 1;
    while (stack.length) {
      const { value, exit } = stack.pop()!;
      if (exit) { active.delete(value as object); complete.add(value as object); continue; }
      if (values > maxValues) throw new PptxImportError("save-limit-exceeded", "Normalized document exceeds maxSnapshotValues.");
      if (value === null || value === undefined || typeof value === "string" || typeof value === "boolean") continue;
      if (typeof value === "number" && Number.isFinite(value)) continue;
      if (typeof value !== "object") throw new TypeError("Normalized document contains an unsupported scalar value.");
      if (active.has(value)) throw new TypeError("Normalized document contains a cycle.");
      if (complete.has(value)) continue;
      if (!Array.isArray(value) && Object.getPrototypeOf(value) !== Object.prototype && Object.getPrototypeOf(value) !== null) throw new TypeError("Normalized document contains a non-record object.");
      const keys = Object.keys(value);
      if (keys.length > maxValues - values) throw new PptxImportError("save-limit-exceeded", "Normalized document exceeds maxSnapshotValues.");
      values += keys.length;
      active.add(value); stack.push({ value, exit: true });
      for (const key of keys) stack.push({ value: (value as Record<string, unknown>)[key] });
    }
    return result;
  } catch (cause) {
    if (cause instanceof PptxImportError) throw cause;
    throw new PptxImportError("save-invalid-document", "The current Core document cannot be normalized into bounded acyclic record state for saving.", undefined, { cause });
  }
}
export function notSerializable(): never { throw new PptxImportError("session-not-serializable", "Import sessions and save plans must be re-created by importing source bytes, not serialized."); }
