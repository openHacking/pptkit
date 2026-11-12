import type { MarkupCompatibilityInspection, MarkupCompatibilityOptions } from "../types/compatibility.js";
import { readXml } from "../xml/read-xml.js";
import { captureConfiguration } from "./configuration.js";
import { processCompatibility } from "./process.js";

export async function inspectCompatibility(input: Uint8Array, options: MarkupCompatibilityOptions): Promise<MarkupCompatibilityInspection> {
  const configuration = captureConfiguration(options), signal = options.signal, part = options.partName;
  // readXml captures source bytes and XML limits before its first yield.
  const parsed = await readXml(input, options, true);
  const { report } = await processCompatibility(parsed.root!, configuration, part, signal);
  return { source: parsed.inspection, ...report };
}
