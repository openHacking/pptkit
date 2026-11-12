import { children, invalidStructure } from "../presentation/structure-xml.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { ContentReader } from "./content-reader.js";
import { chartString } from "./chart-string.js";

export interface ChartReference {
  values?: string[];
  /** Declared positions after a dense prefix, retained only for explicit bubble projection. */
  trailingPadding?: number;
  /** Explicit blank-display cache bounded by a category count or aggregate point budget. */
  blankValues?: (string | null)[];
  points?: readonly { readonly node: XmlElement; readonly lexical: string; readonly index: number }[];
  /** Present source points indexed by their original position in a sparse cache. */
  sparsePoints?: ReadonlyMap<number, { readonly node: XmlElement; readonly lexical: string; readonly index: number }>;
  formula?: string;
  reason?: string;
}

const NUMBER_LEXICAL = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

/** Nonzero source significands must not silently become signed zero. */
export function chartNumberUnderflows(value: string): boolean {
  const lexical = value.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "");
  return NUMBER_LEXICAL.test(lexical) && Number(lexical) === 0 && /[1-9]/.test(lexical.split(/[eE]/, 1)[0]!);
}

/** XML Schema doubles excluding infinities/NaN and values lost to underflow. */
export function chartNumber(value: string): number | undefined {
  const lexical = value.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "");
  if (!NUMBER_LEXICAL.test(lexical)) return undefined;
  const number = Number(lexical);
  // A nonzero source significand that underflows to signed zero is not a
  // representable Core value. Treating it as zero can falsely bind a cache
  // and workbook that disagree and silently lose the source value.
  if (!Number.isFinite(number) || number === 0 && /[1-9]/.test(lexical.split(/[eE]/, 1)[0]!)) return undefined;
  return number;
}

/** Compare mathematical decimal values without equating distinct values that round to one double. */
export function chartNumericSourcesAgree(left: string, right: string): boolean {
  const a = chartNumber(left), b = chartNumber(right);
  if (a === undefined || b === undefined || !Object.is(a, b)) return false;
  if (a === 0 || left === right) return true;
  const canonical = (input: string): string | undefined => {
    const lexical = input.replace(/^[\t\r\n ]+|[\t\r\n ]+$/g, "");
    const match = /^([+-]?)(?:(\d+)(?:\.(\d*))?|\.(\d+))(?:[eE]([+-]?\d+))?$/.exec(lexical);
    if (!match) return undefined;
    const fractional = match[3] ?? match[4] ?? "", digits = ((match[2] ?? "") + fractional).replace(/^0+/, "");
    const significant = digits.replace(/0+$/, ""), trailing = digits.length - significant.length, exponent = Number(match[5] ?? "0");
    if (!Number.isSafeInteger(exponent)) return undefined;
    const scale = exponent - fractional.length + trailing;
    if (!Number.isSafeInteger(scale)) return undefined;
    return `${match[1] === "-" ? "-" : "+"}:${significant}:${scale}`;
  };
  const first = canonical(left), second = canonical(right);
  return first !== undefined && first === second;
}

/** Dense point arrays are indexed by source idx, never by XML document order. */
export async function chartReference(
  node: XmlElement, numeric: boolean, c: string, r: ContentReader,
  checkpoint: () => Promise<void>, countPoint: () => void, captureSources = false, blankLength?: number | "declared", captureSparseSources = false, allowTrailingBubblePadding = false,
): Promise<ChartReference> {
  r.known(node);
  const choices = children(node).filter(n => n.namespaceUri === c && ["strRef", "numRef", "strLit", "numLit", "multiLvlStrRef", "v"].includes(n.local));
  if (choices.length > 1) invalidStructure("Conflicting chart data-source choices.", r.part);
  const choice = choices[0];
  if (!choice) return { reason: "Chart data source is absent." };
  if (choice.local === "multiLvlStrRef") return { reason: "Multilevel chart categories cannot map to Core's flat category array." };
  if (choice.local === "v") return numeric ? { reason: "Numeric series values require an indexed numeric source." } : { values: [chartString(r.text(choice))] };
  if (numeric && !["numRef", "numLit"].includes(choice.local)) return { reason: "Series values do not use a numeric source." };
  r.known(choice);
  const reference = choice.local.endsWith("Ref");
  const formulaNode = reference ? r.one(choice, c, "f", true)! : undefined;
  const formula = formulaNode ? r.text(formulaNode) : undefined;
  const cache = reference ? r.one(choice, c, choice.local === "strRef" ? "strCache" : "numCache") : choice;
  if (!cache) return formula === undefined ? {} : { formula };
  r.known(cache);
  const count = r.one(cache, c, "ptCount");
  const points = children(cache).filter(n => n.namespaceUri === c && n.local === "pt");
  // An omitted count may be inferred only from a genuinely contiguous cache.
  const length = count ? r.number(count, "val", 0, 4294967295) : points.length;
  const indexed = new Map<number, string>(), sourcePoints = new Map<number, XmlElement>(), sourceLexicals = new Map<number, string>();
  for (const point of points) {
    await checkpoint(); countPoint();
    const index = r.number(point, "idx", 0, 4294967295);
    if (index >= length || indexed.has(index)) invalidStructure("Chart cache point index is duplicated or outside ptCount.", r.part);
    const value = r.one(point, c, "v", true)!;
    const text = r.text(value); indexed.set(index, numeric ? text : chartString(text)); if (captureSources) { sourcePoints.set(index, value); sourceLexicals.set(index, text); }
  }
  const result: ChartReference = formula === undefined ? {} : { formula };
  if (!length || indexed.size !== length) {
    // Category charts use an already parsed dense category count. Scatter may
    // use its declared count, charging each missing point to the shared budget
    // before allocation. Both require an explicit supported blank policy.
    if (numeric && allowTrailingBubblePadding && count && indexed.size > 0 && [...indexed.keys()].every(index => index < indexed.size)) {
      return { ...result, values: Array.from({ length: indexed.size }, (_, index) => indexed.get(index)!), trailingPadding: length - indexed.size };
    }
    if (numeric && count && length > 0 && (length === blankLength || blankLength === "declared")) {
      const values: (string | null)[] = [];
      const sparsePoints: Map<number, { index: number; lexical: string; node: XmlElement }> | undefined = captureSources && captureSparseSources ? new Map() : undefined;
      for (let index = 0; index < length; index++) {
        await checkpoint();
        if (!indexed.has(index)) countPoint();
        values.push(indexed.get(index) ?? null);
        const lexical = indexed.get(index);
        if (sparsePoints && lexical !== undefined) sparsePoints.set(index, { index, lexical: sourceLexicals.get(index)!, node: sourcePoints.get(index)! });
      }
      return { ...result, blankValues: values, ...(sparsePoints ? { sparsePoints } : {}) };
    }
    return { ...result, reason: "Empty/sparse chart data cannot map without an explicit supported blank policy and matching bounded category count; missing values are not replaced with zero." };
  }
  // size === length and all distinct indices < length proves density. No
  // allocation or loop is controlled by an untrusted declared ptCount alone.
  result.values = [...indexed].sort((a, b) => a[0] - b[0]).map(([, value]) => value);
  if (captureSources) result.points = [...indexed].sort((a, b) => a[0] - b[0]).map(([index]) => ({ index, lexical: sourceLexicals.get(index)!, node: sourcePoints.get(index)! }));
  return result;
}
