export type PptxCurrencyTextAssessment =
  | { readonly status: "valid"; readonly canonical: string; readonly tenThousandths: string }
  | { readonly status: "invalid" };

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
