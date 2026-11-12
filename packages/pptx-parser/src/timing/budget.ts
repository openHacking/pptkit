import type { PptxTimingInspectionLimits } from "../types/timing.js";
import { PackageInspectionError } from "../types/archive.js";
export class TimingBudget {
  readonly limits: PptxTimingInspectionLimits;
  workUnits = 0; recordCount = 0; textCharacters = 0; findingCount = 0;
  constructor(value?: Partial<PptxTimingInspectionLimits>) {
    const defaults = { maxWorkUnits: 1000000, maxRecords: 100000, maxFindings: 10000, maxTextCharacters: 1000000 };
    if (value !== undefined && (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value)))) throw new TypeError("Invalid timing limits.");
    const captured: Record<string, number> = { ...defaults };
    if (value) for (const key of Reflect.ownKeys(value)) {
      const descriptor = Object.getOwnPropertyDescriptor(value, key)!;
      if (typeof key !== "string" || !Object.hasOwn(defaults, key) || !("value" in descriptor)) throw new TypeError("Unknown/accessor timing limit.");
      if (descriptor.value !== undefined) captured[key] = descriptor.value;
    }
    for (const [key, number] of Object.entries(captured)) if (!Number.isSafeInteger(number) || number <= 0) throw new RangeError(`Invalid timing limit ${key}.`);
    this.limits = captured as unknown as PptxTimingInspectionLimits;
  }
  fail(message: string): never { throw new PackageInspectionError("package-limit-exceeded", message); }
  work(): void { if (++this.workUnits > this.limits.maxWorkUnits) this.fail("Timing inventory exceeds maxWorkUnits."); }
  record(): void { if (++this.recordCount > this.limits.maxRecords) this.fail("Timing inventory exceeds maxRecords."); }
  text(value: string): string { this.charge(value.length); return value; }
  charge(characters: number): void { this.textCharacters += characters; if (this.textCharacters > this.limits.maxTextCharacters) this.fail("Timing inventory exceeds maxTextCharacters."); }
  finding(): void { if (++this.findingCount > this.limits.maxFindings) this.fail("Timing inventory exceeds maxFindings."); }
}
