import { FiniteTimelineError, type FiniteTimelineErrorCode, type FiniteTimelineLimits } from "../types/timeline.js";
export type Limits = Required<FiniteTimelineLimits>;
export const fail = (code: FiniteTimelineErrorCode, path: string, message: string): never => { throw new FiniteTimelineError(code, message, path); };
export function record(value: unknown, keys: readonly string[], path: string, code: FiniteTimelineErrorCode = "invalid-timeline-input"): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) return fail(code, path, "Expected a plain data record.");
  const result = Object.create(null) as Record<string, unknown>;
  for (const key of Reflect.ownKeys(value)) {
    if (typeof key !== "string" || !keys.includes(key)) return fail(code === "invalid-timeline-input" ? "unsupported-timeline-feature" : code, path, "Unknown timeline field; no approximation is applied.");
    const descriptor = Object.getOwnPropertyDescriptor(value, key)!;
    if (!("value" in descriptor) || !descriptor.enumerable) return fail(code, `${path}/${key}`, "Timeline fields must be enumerable data properties, not accessors.");
    result[key] = descriptor.value;
  }
  return result;
}
export function array(value: unknown, maximum: number, path: string, code: FiniteTimelineErrorCode = "invalid-timeline-input"): readonly unknown[] {
  if (!Array.isArray(value)) return fail(code, path, "Expected a dense data array.");
  const length = value.length;
  if (length > maximum) return fail("timeline-limit-exceeded", path, "Timeline array exceeds its node budget.");
  // Snapshot by descriptors: holes, index getters and custom iterator fields
  // cannot execute or quietly change the semantic child sequence.
  const keys = Reflect.ownKeys(value);
  if (keys.length !== length + 1) return fail(code, path, "Timeline arrays must be dense and have no extra fields.");
  const output: unknown[] = [];
  for (let i = 0; i < length; i++) {
    const descriptor = Object.getOwnPropertyDescriptor(value, String(i));
    if (!descriptor || !("value" in descriptor) || !descriptor.enumerable) return fail(code, `${path}/${i}`, "Timeline array indices must be data properties.");
    output.push(descriptor.value);
  }
  return output;
}
export function integer(value: unknown, path: string, min = 0, code: FiniteTimelineErrorCode = "invalid-timeline-input"): number {
  if (typeof value !== "number" || !Number.isSafeInteger(value) || value < min) return fail(code, path, "Expected a nonnegative safe integer (positive for counts).");
  return Object.is(value, -0) ? 0 : value;
}
export function limits(input: FiniteTimelineLimits | undefined): Limits {
  const defaults = { maxNodes: 10000, maxDepth: 128, maxTimeMs: 86400000, maxRepeat: 1000000, maxIdLength: 256, maxTotalIdLength: 1000000 };
  const caps = { maxNodes: 100000, maxDepth: 256, maxTimeMs: Number.MAX_SAFE_INTEGER, maxRepeat: Number.MAX_SAFE_INTEGER, maxIdLength: 65536, maxTotalIdLength: 16000000 };
  const data = record(input === undefined ? {} : input, Object.keys(defaults), "/limits");
  for (const key of Object.keys(defaults) as (keyof Limits)[]) {
    const value = integer(data[key] === undefined ? defaults[key] : data[key], `/limits/${key}`, key === "maxTimeMs" ? 0 : 1);
    if (value > caps[key]) fail("invalid-timeline-input", `/limits/${key}`, "Timeline limit exceeds its supported cap.");
    defaults[key] = value;
  }
  return defaults;
}
export function id(value: unknown, path: string, budget: Limits, code: FiniteTimelineErrorCode = "invalid-timeline-input"): string {
  if (typeof value !== "string" || !value.length) return fail(code, path, "Timeline IDs must be nonempty strings.");
  if (value.length > budget.maxIdLength) return fail("timeline-limit-exceeded", path, "Timeline ID exceeds maxIdLength.");
  return value;
}
export function time(value: unknown, path: string, budget: Limits, code: FiniteTimelineErrorCode = "invalid-timeline-input"): number {
  const number = integer(value, path, 0, code);
  if (number > budget.maxTimeMs) return fail("timeline-limit-exceeded", path, "Timeline exceeds maxTimeMs.");
  return number;
}
export function repeat(value: unknown, path: string, budget: Limits, code: FiniteTimelineErrorCode = "invalid-timeline-input"): number {
  if (code === "invalid-timeline-input" && (value === Infinity || value === "indefinite" || (typeof value === "number" && Number.isFinite(value) && value > 0 && !Number.isInteger(value)))) return fail("unsupported-timeline-feature", path, "Fractional/infinite repetition requires a separate timing contract.");
  const number = integer(value, path, 1, code);
  if (number > budget.maxRepeat) return fail("timeline-limit-exceeded", path, "Timeline exceeds maxRepeat.");
  return number;
}
/** Exact arithmetic before applying the caller's finite-time ceiling. */
export function arithmetic(a: number, b: number, operation: "add" | "multiply", path: string, budget: Limits, code: FiniteTimelineErrorCode = "timeline-time-overflow"): number {
  const result = operation === "add" ? BigInt(a) + BigInt(b) : BigInt(a) * BigInt(b);
  if (result > BigInt(Number.MAX_SAFE_INTEGER)) return fail(code, path, "Timeline arithmetic exceeds safe integer milliseconds.");
  if (result > BigInt(budget.maxTimeMs)) return fail("timeline-limit-exceeded", path, "Timeline arithmetic exceeds maxTimeMs.");
  return Number(result);
}
