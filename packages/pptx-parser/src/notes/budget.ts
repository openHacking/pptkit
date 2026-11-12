import type { PptxNotesInspectionLimits } from "../types/notes.js";
import { PackageInspectionError } from "../types/archive.js";

export const DEFAULT_NOTES_INSPECTION_LIMITS: Readonly<PptxNotesInspectionLimits> = Object.freeze({ maxWorkUnits: 1000000, maxRecords: 100000, maxFindings: 10000, maxTextCharacters: 1000000 });
export class NotesBudget {
  readonly limits: PptxNotesInspectionLimits;
  workUnits = 0; recordCount = 0; textCharacters = 0; findingCount = 0;
  constructor(value?: Partial<PptxNotesInspectionLimits>) {
    if (value !== undefined && (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value)))) throw new TypeError("Invalid notes inspection limits.");
    const captured: Record<string, number> = { ...DEFAULT_NOTES_INSPECTION_LIMITS };
    if (value) for (const key of Reflect.ownKeys(value)) {
      const descriptor = Object.getOwnPropertyDescriptor(value, key)!;
      if (typeof key !== "string" || !Object.hasOwn(captured, key) || !("value" in descriptor)) throw new TypeError("Unknown/accessor notes limit.");
      if (descriptor.value !== undefined) captured[key] = descriptor.value;
    }
    for (const [key, n] of Object.entries(captured)) if (!Number.isSafeInteger(n) || n <= 0) throw new RangeError(`Invalid notes limit ${key}.`);
    this.limits = captured as unknown as PptxNotesInspectionLimits;
  }
  private fail(name: string): never { throw new PackageInspectionError("package-limit-exceeded", `Notes inventory exceeds ${name}.`); }
  work(): void { if (++this.workUnits > this.limits.maxWorkUnits) this.fail("maxWorkUnits"); }
  record(): void { if (++this.recordCount > this.limits.maxRecords) this.fail("maxRecords"); }
  text(value: string): string { this.textCharacters += value.length; if (this.textCharacters > this.limits.maxTextCharacters) this.fail("maxTextCharacters"); return value; }
  finding(): void { if (++this.findingCount > this.limits.maxFindings) this.fail("maxFindings"); }
}
