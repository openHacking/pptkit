import { chartNumberText } from "./chart-workbook.js";
import { chartStringXml } from "./chart-string.js";

export function chartStringReference(formula: string, values: readonly string[]): string {
  const points = values.map((value, index) => `<c:pt idx="${index}"><c:v>${chartStringXml(value)}</c:v></c:pt>`).join("");
  return `<c:strRef><c:f>${formula}</c:f><c:strCache><c:ptCount val="${values.length}"/>${points}</c:strCache></c:strRef>`;
}

export function chartNumberReference(formula: string, values: readonly (number | null)[]): string {
  const points = values.map((value, index) => value === null ? "" : `<c:pt idx="${index}"><c:v>${chartNumberText(value)}</c:v></c:pt>`).join("");
  return `<c:numRef><c:f>${formula}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache></c:numRef>`;
}
