import type { CustomDocumentProperty } from "@pptkit/core";
import type { CustomPropertyInfo } from "../types/metadata.js";
import { assessPptxCurrencyNumber } from "../metadata/currency-text.js";

// Core's integer contract is int32. Intersect every source variant's range
// with that contract so neither projection nor edits silently change its type.
const integerRanges = new Map<string, readonly [number, number]>([
  ["i1", [-128, 127]], ["i2", [-32768, 32767]], ["i4", [-2147483648, 2147483647]], ["int", [-2147483648, 2147483647]], ["i8", [-2147483648, 2147483647]],
  ["ui1", [0, 255]], ["ui2", [0, 65535]], ["ui4", [0, 2147483647]], ["uint", [0, 2147483647]], ["ui8", [0, 2147483647]],
]);
export function sourceIntegerAccepts(type: string, value: number): boolean | undefined {
  const range = integerRanges.get(type);
  return range ? Number.isInteger(value) && !Object.is(value, -0) && value >= range[0] && value <= range[1] : undefined;
}
/** Decimal text must equal the binary64 value mathematically, not just after rounding. */
function exactDecimal(value: number): boolean {
  if (!Number.isFinite(value) || Object.is(value, -0)) return false;
  const text = String(value), match = /^(-?)(\d+)(?:\.(\d+))?$/.exec(text);
  if (!match || text.length > 32) return false;
  const digits = BigInt(match[2]! + (match[3] ?? "")), scale = 10n ** BigInt((match[3] ?? "").length);
  const bytes = new DataView(new ArrayBuffer(8)); bytes.setFloat64(0, Math.abs(value), false);
  const bits = bytes.getBigUint64(0, false), exponent = Number((bits >> 52n) & 2047n), fraction = bits & ((1n << 52n) - 1n);
  const mantissa = exponent === 0 ? fraction : (1n << 52n) | fraction;
  const power = exponent === 0 ? -1074 : exponent - 1075;
  return power >= 0 ? digits === (mantissa << BigInt(power)) * scale : (digits << BigInt(-power)) === mantissa * scale;
}

/** r4 and decimal are admitted only when the source value is exact in Core. */
export function sourceRealAccepts(type: string, value: number): boolean | undefined {
  if (type === "decimal") return exactDecimal(value);
  if (type !== "r4") return undefined;
  return Number.isFinite(value) && Object.is(Math.fround(value), value);
}

/** Convert explicitly zoned values only when fractional precision is exactly milliseconds. */
export function sourceDateToUtc(value: string): string | undefined {
  const match = /^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(?:\.(\d+))?(Z|[+-]\d{2}:\d{2})$/.exec(value);
  if (!match || match[3] === "-00:00") return undefined;
  const fraction = match[2] ?? "";
  if (/[1-9]/.test(fraction.slice(3))) return undefined;
  const local = `${match[1]}.${fraction.slice(0, 3).padEnd(3, "0")}Z`, timestamp = Date.parse(local);
  if (!Number.isFinite(timestamp) || new Date(timestamp).toISOString() !== local) return undefined;
  const zone = match[3]!;
  let offset = 0;
  if (zone !== "Z") {
    const hours = Number(zone.slice(1, 3)), minutes = Number(zone.slice(4, 6));
    if (minutes > 59 || hours > 14 || hours === 14 && minutes !== 0) return undefined;
    offset = (hours * 60 + minutes) * (zone[0] === "+" ? 1 : -1);
  }
  const instant = new Date(timestamp - offset * 60000);
  if (!Number.isFinite(instant.getTime())) return undefined;
  const canonical = instant.toISOString();
  return /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(canonical) ? canonical : undefined;
}

/** Decode escaped literal underscores once, without interpreting other variants. */
export function decodeLiteralEscapes(input: string): string | undefined {
  const pattern = /_x([0-9a-fA-F]{4})_/g;
  let output = "", cursor = 0;
  for (;;) {
    pattern.lastIndex = cursor;
    const match = pattern.exec(input);
    if (!match) {
      const decoded = output + input.slice(cursor);
      return /_x[0-9a-fA-F]{4}_x[0-9a-fA-F]{4}_/.test(decoded) ? undefined : decoded;
    }
    output += input.slice(cursor, match.index);
    const rest = input.slice(match.index + match[0].length, match.index + match[0].length + 6);
    if (match[1]!.toLowerCase() !== "005f" || !/^x[0-9a-fA-F]{4}_/.test(rest)) return undefined;
    output += `_${rest.slice(0, 6)}`;
    cursor = match.index + match[0].length + 6;
  }
}

/** Decode one OOXML escape layer solely to detect collisions, without projecting unsupported names. */
export function customNameCollisionKey(input: string): string {
  return input.replace(/_x([0-9a-fA-F]{4})_/g, (_match, hex: string) => String.fromCharCode(Number.parseInt(hex, 16)));
}

/** Optional read-only projection; currency remains an opaque source value. */
export function projectCurrencyProperty(property: CustomPropertyInfo, text: string): CustomDocumentProperty | undefined {
  if (property.valueType !== "cy" || property.value.status !== "uninterpreted" || !property.name || property.linkTarget !== undefined) return undefined;
  const name = decodeLiteralEscapes(property.name);
  if (!name || text.length > 1000000) return undefined;
  const value = assessPptxCurrencyNumber(text);
  return value.status === "exact" ? { name, type: "real", value: value.value } : undefined;
}

/** Map only an explicitly typed source subset without interpreting opaque variants. */
export function projectCustomProperty(property: CustomPropertyInfo): CustomDocumentProperty | undefined {
  if (!property.name || property.linkTarget !== undefined || property.value.status !== "scalar") return undefined;
  const name = decodeLiteralEscapes(property.name), value = property.value.value;
  if (!name) return undefined;
  switch (property.valueType) {
    case "lpstr": case "lpwstr": case "bstr": {
      const decoded = decodeLiteralEscapes(value);
      return decoded === undefined ? undefined : { name, type: "string", value: decoded };
    }
    case "bool":
      if (value === "true" || value === "1") return { name, type: "boolean", value: true };
      if (value === "false" || value === "0") return { name, type: "boolean", value: false };
      return undefined;
    case "i1": case "i2": case "i4": case "int": case "i8": case "ui1": case "ui2": case "ui4": case "uint": case "ui8": {
      const number = Number(value);
      return sourceIntegerAccepts(property.valueType, number) && String(number) === value ? { name, type: "integer", value: number } : undefined;
    }
    case "decimal": case "r4": case "r8": {
      const number = Number(value);
      return Number.isFinite(number) && sourceRealAccepts(property.valueType, number) !== false && (String(number) === value || value === "-0" && Object.is(number, -0)) ? { name, type: "real", value: number } : undefined;
    }
    case "date": case "filetime": {
      const canonical = sourceDateToUtc(value);
      return canonical === undefined ? undefined : { name, type: "date", value: canonical };
    }
    default: return undefined;
  }
}
