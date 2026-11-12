import { CliError } from "./arguments.js";
/** Encode bounded output without first materializing the complete serialized report. */
export function boundedJson(value: unknown, limit: number): string {
  const tokens: string[] = []; let bytes = 1; // final newline belongs to the bound
  const append = (token: string) => { bytes += Buffer.byteLength(token); if (bytes > limit) throw new CliError("output-limit-exceeded", "Report exceeds max-output-bytes; increase the limit.", 4); tokens.push(token); };
  const quote = (text: string) => { if (text.length + bytes + 2 > limit) throw new CliError("output-limit-exceeded", "Report exceeds max-output-bytes; increase the limit.", 4); append(JSON.stringify(text)); };
  const visit = (item: unknown): void => {
    if (item === null || typeof item === "boolean" || typeof item === "number") { append(JSON.stringify(item)); return; }
    if (typeof item === "string") { quote(item); return; }
    if (Array.isArray(item)) { append("["); item.forEach((v, i) => { if (i) append(","); visit(v); }); append("]"); return; }
    if (typeof item !== "object") throw new TypeError("Unexpected CLI report value.");
    append("{"); let first = true;
    for (const [key, entry] of Object.entries(item)) if (entry !== undefined) { if (!first) append(","); first = false; quote(key); append(":"); visit(entry); }
    append("}");
  };
  visit(value); return tokens.join("");
}
export const terminalText = (value: string): string => value.replace(/[\x00-\x1f\x7f-\x9f\u200e-\u200f\u2028-\u202e\u2066-\u2069]/g, c => `\\u${c.charCodeAt(0).toString(16).padStart(4, "0")}`);
export function boundedText(lines: readonly string[], limit: number): string {
  const parts: string[] = []; let bytes = 1;
  for (const line of lines) {
    if (line.length + bytes > limit) throw new CliError("output-limit-exceeded", "Report exceeds max-output-bytes; increase the limit.", 4);
    const safe = terminalText(line); bytes += Buffer.byteLength(safe) + (parts.length ? 1 : 0);
    if (bytes > limit) throw new CliError("output-limit-exceeded", "Report exceeds max-output-bytes; increase the limit.", 4);
    parts.push(safe);
  }
  return parts.join("\n");
}
