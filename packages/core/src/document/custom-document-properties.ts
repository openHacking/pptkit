import type { CustomDocumentProperty } from "../types/presentation.js";

/** Canonical UTC milliseconds give date fields a deterministic, clock-free contract. */
export function isCanonicalUtcDate(value: unknown): value is string {
  return typeof value === "string" && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(value) && Number.isFinite(Date.parse(value)) && new Date(value).toISOString() === value;
}

/** Exact, canonical decimal text; bounded without any floating-point conversion. */
export function isCanonicalDecimal(value: unknown): value is string {
  if (typeof value !== "string" || value.length > 10000 || value === "-0") return false;
  const match = /^-?(?:0|[1-9]\d*)(?:\.\d*[1-9])?$/.exec(value);
  return match !== null && match[0].length === value.length;
}

/** Reject ambiguous names and values before storing authoring state. */
export function assertCustomDocumentProperties(input: unknown): asserts input is readonly CustomDocumentProperty[] {
  if (!Array.isArray(input) || Reflect.ownKeys(input).length !== input.length + 1) throw new TypeError("customProperties must be a dense array.");
  const names = new Set<string>();
  for (const item of input) {
    if (!item || typeof item !== "object" || Array.isArray(item) || ![Object.prototype, null].includes(Object.getPrototypeOf(item)) || Reflect.ownKeys(item).some(key => typeof key !== "string" || !["name", "type", "value"].includes(key) || !Object.prototype.propertyIsEnumerable.call(item, key))) throw new TypeError("Invalid custom property record.");
    const { name, type, value } = item;
    if (typeof name !== "string" || name.length === 0 || names.has(name)) throw new TypeError("Custom property names must be nonempty and unique.");
    names.add(name);
    const date = isCanonicalUtcDate(value);
    if (!(type === "string" && typeof value === "string" || type === "boolean" && typeof value === "boolean" || type === "integer" && typeof value === "number" && Number.isInteger(value) && value >= -2147483648 && value <= 2147483647 || type === "real" && typeof value === "number" && Number.isFinite(value) || type === "date" && date || type === "decimal" && isCanonicalDecimal(value))) throw new TypeError("Invalid custom property type or value.");
  }
}
