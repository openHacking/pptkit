import { TableEditError, type TableEditErrorCode, type TableEditLimits } from "../types/table-edit.js";

export function fail(code: TableEditErrorCode, path: string, message: string): never { throw new TableEditError(code, message, path); }
export function object(value: unknown, path: string): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) return fail("invalid-table-edit-input", path, "Expected a plain data record.");
  return value as Record<string, unknown>;
}
export function fields(value: unknown, allowed: readonly string[], path: string): Record<string, unknown> {
  const data = object(value, path), result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(data)) {
    if (typeof key !== "string" || !allowed.includes(key)) fail("invalid-table-edit-input", path, "Unknown table edit field.");
    const descriptor = Object.getOwnPropertyDescriptor(data, key)!;
    if (!("value" in descriptor) || !descriptor.enumerable) fail("invalid-table-edit-input", path, "Expected enumerable data properties, not accessors.");
    result[key] = descriptor.value;
  }
  return result;
}
export function integer(value: unknown, path: string, min = 0): number {
  if (typeof value !== "number" || !Number.isSafeInteger(value) || value < min) return fail("invalid-table-edit-input", path, "Expected a safe integer within the supported range.");
  return Object.is(value, -0) ? 0 : value;
}
export function positive(value: unknown, path: string): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) return fail("invalid-table-edit-input", path, "Expected a positive finite dimension.");
  return value;
}
export class EditBudget {
  readonly limits: Required<TableEditLimits>;
  private work = 0;
  private values = 0;
  private characters = 0;
  constructor(input: unknown) {
    const defaults = { maxRows: 1000, maxColumns: 1000, maxGridSlots: 1000000, maxCells: 100000, maxOperations: 1000, maxDataValues: 1000000, maxTextCharacters: 16000000, maxWorkUnits: 10000000, maxDepth: 128 };
    const caps = { maxRows: 10000, maxColumns: 10000, maxGridSlots: 2000000, maxCells: 200000, maxOperations: 10000, maxDataValues: 2000000, maxTextCharacters: 32000000, maxWorkUnits: 100000000, maxDepth: 256 };
    const data = fields(input === undefined ? {} : input, Object.keys(defaults), "/options/limits");
    for (const key of Object.keys(defaults) as (keyof typeof defaults)[]) {
      const value = integer(data[key] === undefined ? defaults[key] : data[key], `/options/limits/${key}`, 1);
      if (value > caps[key]) fail("invalid-table-edit-input", `/options/limits/${key}`, "Limit exceeds the supported cap.");
      defaults[key] = value;
    }
    this.limits = defaults;
  }
  step(units = 1, path = "/"): void {
    this.work += units;
    if (!Number.isSafeInteger(this.work) || this.work > this.limits.maxWorkUnits) fail("table-edit-limit-exceeded", path, "Table transaction exceeds maxWorkUnits.");
  }
  size(rows: number, columns: number, cells = 0): number {
    const count = rows * columns;
    if (rows > this.limits.maxRows || columns > this.limits.maxColumns || count > this.limits.maxGridSlots || cells > this.limits.maxCells) fail("table-edit-limit-exceeded", "/table", "Table dimensions, cells or grid slots exceed the transaction budget.");
    return count;
  }
  /** Bounded descriptor capture: aliases are copied, cycles/accessors never executed. */
  capture<T>(value: T, path: string, depth = 0, ancestors = new Set<object>()): T {
    this.step(1, path);
    if (++this.values > this.limits.maxDataValues || depth > this.limits.maxDepth) fail("table-edit-limit-exceeded", path, "Table data values/depth exceed the transaction budget.");
    if (typeof value === "string") {
      this.characters += value.length;
      if (this.characters > this.limits.maxTextCharacters) fail("table-edit-limit-exceeded", path, "Table data exceeds maxTextCharacters.");
      return value;
    }
    if (value === null || value === undefined || typeof value === "boolean") return value;
    if (typeof value === "number" && Number.isFinite(value)) return value;
    if (typeof value !== "object") return fail("invalid-table-edit-input", path, "Table edits require finite plain data values.");
    if (ancestors.has(value)) fail("invalid-table-edit-input", path, "Cyclic table data is unsupported.");
    if (!Array.isArray(value)) object(value, path);
    else if (value.length > this.limits.maxDataValues - this.values) fail("table-edit-limit-exceeded", path, "Array exceeds the remaining data budget.");
    const keys = Reflect.ownKeys(value);
    if (Array.isArray(value) && keys.length !== value.length + 1) fail("invalid-table-edit-input", path, "Expected a dense array without custom properties.");
    const output = (Array.isArray(value) ? [] : {}) as Record<string, unknown>;
    ancestors.add(value);
    for (const key of keys) {
      if (Array.isArray(value) && key === "length") continue;
      if (typeof key !== "string") fail("invalid-table-edit-input", path, "Symbol fields are unsupported.");
      if (Array.isArray(value) && (!Number.isSafeInteger(Number(key)) || String(Number(key)) !== key || Number(key) < 0 || Number(key) >= value.length)) fail("invalid-table-edit-input", path, "Expected dense numeric array indices.");
      const descriptor = Object.getOwnPropertyDescriptor(value, key)!;
      if (!("value" in descriptor) || !descriptor.enumerable) fail("invalid-table-edit-input", path, "Accessors/non-enumerable fields are unsupported.");
      // defineProperty also preserves an ordinary own __proto__ field safely.
      this.characters += key.length;
      if (this.characters > this.limits.maxTextCharacters) fail("table-edit-limit-exceeded", path, "Table keys exceed maxTextCharacters.");
      this.step(path.length + key.length + 1, path);
      Object.defineProperty(output, key, { value: this.capture(descriptor.value, `${path}/${key}`, depth + 1, ancestors), enumerable: true, writable: true, configurable: true });
    }
    ancestors.delete(value);
    return output as T;
  }
}

export function freeze<T>(input: T): T {
  const stack: unknown[] = [input];
  while (stack.length) {
    const value = stack.pop();
    if (value && typeof value === "object" && !Object.isFrozen(value)) { Object.freeze(value); for (const child of Object.values(value)) stack.push(child); }
  }
  return input;
}
