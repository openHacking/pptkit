export type PptxCurrencyTextAssessment =
  | { readonly status: "valid"; readonly canonical: string; readonly tenThousandths: string }
  | { readonly status: "invalid" };

export type PptxCurrencyNumberAssessment =
  | { readonly status: "exact"; readonly value: number }
  | { readonly status: "inexact" | "out-of-range" | "invalid" };

/** Assess ST_Cy text without floating-point conversion or assumptions about currency units. */
export function assessPptxCurrencyText(value: string): PptxCurrencyTextAssessment {
  if (typeof value !== "string") throw new TypeError("Currency text must be a string.");
  if (value.length > 1000000) throw new RangeError("Currency text exceeds one million UTF-16 code units.");
  // ECMA-376 Strict and Transitional ST_Cy restrict whitespace, digits and scale.
  const match = /^[\x20\t\r\n]*([0-9]*)\.([0-9]{4})[\x20\t\r\n]*$/.exec(value);
  if (!match || match[0].length !== value.length) return Object.freeze({ status: "invalid" });
  const whole = match[1]!.replace(/^0+/, "") || "0", fractional = match[2]!;
  const tenThousandths = (whole + fractional).replace(/^0+/, "") || "0";
  return Object.freeze({ status: "valid", canonical: `${whole}.${fractional}`, tenThousandths });
}

/** Return a number only when its binary64 value is exactly the source decimal. */
export function assessPptxCurrencyNumber(value: string): PptxCurrencyNumberAssessment {
  const text = assessPptxCurrencyText(value);
  if (text.status === "invalid") return Object.freeze({ status: "invalid" });
  // Any value with more than 309 integer digits is larger than finite binary64.
  if (text.tenThousandths.length > 313) return Object.freeze({ status: "out-of-range" });
  const number = Number(text.canonical);
  if (!Number.isFinite(number)) return Object.freeze({ status: "out-of-range" });
  const units = BigInt(text.tenThousandths);
  const bytes = new DataView(new ArrayBuffer(8)); bytes.setFloat64(0, number, false);
  const bits = bytes.getBigUint64(0, false), exponent = Number((bits >> 52n) & 2047n), fraction = bits & ((1n << 52n) - 1n);
  const mantissa = exponent === 0 ? fraction : (1n << 52n) | fraction;
  const power = exponent === 0 ? -1074 : exponent - 1075;
  const exact = power >= 0 ? units === (mantissa << BigInt(power)) * 10000n : (units << BigInt(-power)) === mantissa * 10000n;
  return exact ? Object.freeze({ status: "exact", value: number }) : Object.freeze({ status: "inexact" });
}
