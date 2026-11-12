import { readOpcSource, snapshotOptions } from "../opc/read-package.js";
import { byteViewInfo, copyByteView } from "../utils/byte-view.js";
import { PackageInspectionError } from "../types/archive.js";
import type { OpcPartInfo, OpcRelationshipInfo } from "../types/opc.js";
import type { OpcComparisonLimits, OpcComparisonOptions, OpcPackageComparison, OpcPackageDifference } from "../types/comparison.js";
import { xmlTokens } from "./xml-tokens.js";

export const DEFAULT_COMPARISON_LIMITS: Readonly<OpcComparisonLimits> = Object.freeze({
  maxTotalExpandedBytes: 128 * 1024 * 1024,
  maxWorkUnits: 512 * 1024 * 1024,
  maxDifferences: 1000,
});
const TYPES = "/[Content_Types].xml";
const order = (a: string, b: string): number => a < b ? -1 : a > b ? 1 : 0;
type Source = Awaited<ReturnType<typeof readOpcSource>>;

/** Independent source/package comparison: never imports a Core document or uses Exporter. */
export async function comparePackages(beforeInput: Uint8Array, afterInput: Uint8Array, options: OpcComparisonOptions = {}): Promise<OpcPackageComparison> {
  const inspectionOptions = snapshotOptions(options);
  const profile = options.profile ?? "opc-payload-v1";
  if (profile !== "opc-payload-v1" && profile !== "opc-xml-v1") throw new TypeError("Invalid OPC comparison profile");
  if (options.comparisonLimits !== undefined && (!options.comparisonLimits || typeof options.comparisonLimits !== "object" || Array.isArray(options.comparisonLimits))) throw new TypeError("Invalid OPC comparison limits");
  const limits = { ...DEFAULT_COMPARISON_LIMITS, ...options.comparisonLimits };
  for (const [name, value] of Object.entries(limits)) if (!Number.isSafeInteger(value) || value <= 0) throw new RangeError(`Invalid OPC comparison limit ${name}`);
  for (const input of [beforeInput, afterInput]) {
    if (!(input instanceof Uint8Array)) throw new TypeError("OPC comparison inputs must be Uint8Array");
    const view = byteViewInfo(input);
    if (view.shared) throw new TypeError("OPC comparison inputs must not use SharedArrayBuffer.");
    if (view.length > inspectionOptions.zipLimits.maxInputBytes) throw new PackageInspectionError("archive-limit-exceeded", "Comparison input exceeds maxInputBytes.");
  }
  const { signal } = inspectionOptions;
  const abort = (): void => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "Package comparison was aborted."); };
  const failLimit = (message: string): never => { throw new PackageInspectionError("package-limit-exceeded", message); };
  abort();
  // Both are captured before the first await, including when callers reuse/mutate a buffer.
  const beforeBytes = copyByteView(beforeInput), afterBytes = copyByteView(afterInput);
  const read = (input: Uint8Array, available: number) => readOpcSource(input, {
    limits: inspectionOptions.limits, xmlLimits: inspectionOptions.xmlLimits, ...(signal ? { signal } : {}),
    zipLimits: { ...inspectionOptions.zipLimits, maxExpandedBytes: Math.min(inspectionOptions.zipLimits.maxExpandedBytes, available) },
  }, profile === "opc-xml-v1", true);
  const before = await read(beforeBytes, limits.maxTotalExpandedBytes);
  const available = limits.maxTotalExpandedBytes - before.inspection.archive.expandedBytes;
  if (available <= 0) failLimit("Comparison inputs exceed maxTotalExpandedBytes.");
  const after = await read(afterBytes, available);
  const differences: OpcPackageDifference[] = [];
  let differenceCount = 0, workUnits = 0, lastYield = 0;
  const work = async (units: number): Promise<void> => {
    abort();
    if (units > limits.maxWorkUnits - workUnits) failLimit("Comparison exceeds maxWorkUnits.");
    workUnits += units;
    if (workUnits - lastYield >= 32768) {
      lastYield = workUnits;
      await new Promise<void>((resolve) => setTimeout(resolve, 0)); abort();
    }
  };
  const add = (difference: OpcPackageDifference): void => {
    differenceCount++;
    if (differences.length < limits.maxDifferences) differences.push(Object.freeze(difference));
  };
  const parts = (source: Source): Map<string, Pick<OpcPartInfo, "entryName" | "contentType">> => new Map<string, Pick<OpcPartInfo, "entryName" | "contentType">>([
    [TYPES, { entryName: "[Content_Types].xml", contentType: "application/xml" }],
    ...source.inspection.parts.map((part) => [part.name, part] as const),
  ]);
  const beforeParts = parts(before), afterParts = parts(after);
  for (const name of [...new Set([...beforeParts.keys(), ...afterParts.keys()])].sort(order)) {
    await work(name.length + 1);
    const left = beforeParts.get(name), right = afterParts.get(name);
    if (!left || !right) {
      add({ code: left ? "part-removed" : "part-added", part: name }); continue;
    }
    if (left.contentType !== right.contentType) add({ code: "content-type-changed", part: name });
    const leftXml = before.xmlDocuments.get(name), rightXml = after.xmlDocuments.get(name);
    if (profile === "opc-xml-v1" && leftXml && rightXml) {
      const a = xmlTokens(leftXml), b = xmlTokens(rightXml);
      while (true) {
        const x = a.next(), y = b.next();
        await work((x.value?.value.length ?? 0) + (y.value?.value.length ?? 0) + 1);
        if (x.done && y.done) break;
        if (x.done || y.done || x.value.value !== y.value.value) {
          add({ code: "xml-changed", part: name, path: x.value?.path ?? y.value!.path }); break;
        }
      }
    } else {
      const a = before.payloads!.get(left.entryName)!, b = after.payloads!.get(right.entryName)!;
      let offset: number | undefined;
      for (let start = 0; start < Math.min(a.length, b.length); start += 16384) {
        const end = Math.min(start + 16384, a.length, b.length);
        await work(2 * (end - start));
        for (let i = start; i < end; i++) if (a[i] !== b[i]) { offset = i; break; }
        if (offset !== undefined) break;
      }
      if (offset === undefined && a.length !== b.length) offset = Math.min(a.length, b.length);
      if (offset !== undefined) add({ code: "payload-changed", part: name, byteOffset: offset });
    }
  }
  const relationships = (source: Source): Map<string, OpcRelationshipInfo> => new Map(source.inspection.relationships.map((edge) => [JSON.stringify([edge.source, edge.id]), edge]));
  const beforeEdges = relationships(before), afterEdges = relationships(after);
  for (const key of [...new Set([...beforeEdges.keys(), ...afterEdges.keys()])].sort(order)) {
    const left = beforeEdges.get(key), right = afterEdges.get(key), edge = left ?? right!;
    const value = (item: OpcRelationshipInfo | undefined): string => JSON.stringify(item ? [item.type, item.targetMode, item.target, item.targetPart ?? null, item.fragment ?? null] : null);
    const a = value(left), b = value(right);
    await work(key.length + a.length + b.length + 1);
    if (a !== b) add({ code: !left ? "relationship-added" : !right ? "relationship-removed" : "relationship-changed", part: edge.relationshipPart, relationshipId: edge.id });
  }
  abort();
  const summary = (source: Source) => Object.freeze({ parts: source.inspection.parts.length, relationships: source.inspection.relationships.length, expandedBytes: source.inspection.archive.expandedBytes });
  return Object.freeze({
    profile, equal: differenceCount === 0, before: summary(before), after: summary(after),
    differenceCount, differences: Object.freeze(differences), truncated: differenceCount > differences.length, workUnits,
  });
}
