import { readOpcSource, snapshotOptions } from "../opc/read-package.js";
import { asciiKey, partName } from "../opc/names.js";
import { PackageInspectionError } from "../types/archive.js";
import type { OpcClosureFinding, OpcClosureInspection, OpcClosureLimits, OpcClosureOptions, OpcClosurePart, OpcSourceRelationship } from "../types/closure.js";
import { sourceRelationships } from "./source-relationships.js";

export const DEFAULT_CLOSURE_LIMITS: Readonly<OpcClosureLimits> = Object.freeze({
  maxRoots: 1000, maxSourceRelationships: 100000, maxClosureParts: 20000,
  maxClosureExpandedBytes: 128 * 1024 * 1024, maxFindings: 10000, maxWorkUnits: 16000000,
});
const order = (a: string, b: string): number => a < b ? -1 : a > b ? 1 : 0;

/** Graph inventory only: never authorizes garbage collection or creates a preservation writer. */
export async function inspectClosure(input: Uint8Array, options: OpcClosureOptions): Promise<OpcClosureInspection> {
  const captured = snapshotOptions(options);
  const providedRoots = options.roots, providedLimits = options.closureLimits;
  if (!Array.isArray(providedRoots) || !providedRoots.length) throw new TypeError("OPC closure requires a nonempty array of logical roots.");
  if (providedLimits !== undefined && (!providedLimits || typeof providedLimits !== "object" || Array.isArray(providedLimits))) throw new TypeError("Invalid OPC closure limits");
  const limits = { ...DEFAULT_CLOSURE_LIMITS, ...providedLimits };
  for (const [name, value] of Object.entries(limits)) if (!Number.isSafeInteger(value) || value <= 0) throw new RangeError(`Invalid OPC closure limit ${name}`);
  const failLimit = (name: keyof OpcClosureLimits): never => { throw new PackageInspectionError("package-limit-exceeded", `OPC closure exceeds ${name}.`); };
  if (providedRoots.length > limits.maxRoots) failLimit("maxRoots");
  let rootUnits = 0;
  const requestedRoots = providedRoots.map((root: unknown) => {
    if (typeof root !== "string") throw new TypeError("OPC closure roots must be strings.");
    if (root.length + 1 > limits.maxWorkUnits - rootUnits) failLimit("maxWorkUnits");
    rootUnits += root.length + 1;
    return root === "/" ? root : partName(root, root);
  });
  const { signal } = captured;
  const abort = () => { if (signal?.aborted) throw new PackageInspectionError("operation-aborted", "OPC closure inspection was aborted."); };
  abort();
  // The source reader snapshots input synchronously; ordinary XML is validated without retaining its tree.
  const source = await readOpcSource(input, { limits: captured.limits, xmlLimits: captured.xmlLimits, zipLimits: captured.zipLimits, ...(signal ? { signal } : {}) }, "relationships");
  const parts = new Map(source.inspection.parts.map((part) => [part.name, part]));
  const byKey = new Map(source.inspection.parts.map((part) => [asciiKey(part.name), part]));
  const roots = [...new Set(requestedRoots.map((root) => {
    if (root === "/") return root;
    const part = byKey.get(asciiKey(root));
    if (!part) throw new PackageInspectionError("package-invalid", "OPC closure root is absent.", root);
    if (part.relationshipPart) throw new TypeError("OPC closure roots must be ordinary parts or '/'.");
    return part.name;
  }))].sort(order);
  let workUnits = 0, lastYield = 0, sourceRelationshipCount = 0;
  const work = async (units: number): Promise<void> => {
    abort();
    if (units > limits.maxWorkUnits - workUnits) failLimit("maxWorkUnits");
    workUnits += units;
    if (workUnits - lastYield >= 4096) {
      lastYield = workUnits; await new Promise<void>((resolve) => setTimeout(resolve, 0)); abort();
    }
  };
  const findings: OpcClosureFinding[] = [];
  await work(rootUnits);
  const addFinding = (finding: OpcClosureFinding): void => {
    if (findings.length >= limits.maxFindings) failLimit("maxFindings");
    findings.push(Object.freeze(finding));
  };
  const { edges, sets } = await sourceRelationships(source.inspection, source.xmlDocuments, work, addFinding, () => {
    if (++sourceRelationshipCount > limits.maxSourceRelationships) failLimit("maxSourceRelationships");
  });
  const outgoing = new Map<string, OpcSourceRelationship[]>();
  for (const edge of edges) {
    await work(1);
    let list = outgoing.get(edge.source);
    if (!list) { list = []; outgoing.set(edge.source, list); }
    list.push(edge);
  }
  const controlStream = Object.freeze({ name: "/[Content_Types].xml" as const, expandedBytes: source.inspection.archive.entries.find((entry) => entry.name === "[Content_Types].xml")!.expandedBytes });
  let expandedBytes = controlStream.expandedBytes;
  if (expandedBytes > limits.maxClosureExpandedBytes) failLimit("maxClosureExpandedBytes");
  const selected = new Map<string, OpcClosurePart>(), visited = new Set<string>(), queue = [...roots];
  const rootSet = new Set(roots);
  const select = (name: string, reason: OpcClosurePart["reason"], via?: OpcSourceRelationship): void => {
    if (selected.has(name)) return;
    if (selected.size >= limits.maxClosureParts) failLimit("maxClosureParts");
    const part = parts.get(name)!;
    if (part.expandedBytes > limits.maxClosureExpandedBytes - expandedBytes) failLimit("maxClosureExpandedBytes");
    expandedBytes += part.expandedBytes;
    selected.set(name, Object.freeze({
      name, entryName: part.entryName, contentType: part.contentType, expandedBytes: part.expandedBytes,
      reason: rootSet.has(name) ? "root" : reason,
      ...(!rootSet.has(name) && via ? { via: Object.freeze({ source: via.source, ...(via.id !== undefined ? { relationshipId: via.id } : {}), elementIndex: via.elementIndex }) } : {}),
    }));
  };
  // Enqueue each ordinary source once, including cyclic/shared dependency graphs.
  const enqueued = new Set(roots);
  for (let index = 0; index < queue.length; index++) {
    const name = queue[index]!;
    await work(name.length + 1);
    visited.add(name);
    if (name !== "/") select(name, "root");
    const relationshipPart = sets.get(name);
    if (relationshipPart) select(relationshipPart, "relationship-set");
    for (const edge of outgoing.get(name) ?? []) {
      await work(1);
      if (!edge.targetPart) continue;
      select(edge.targetPart, "dependency", edge);
      if (!enqueued.has(edge.targetPart)) { enqueued.add(edge.targetPart); queue.push(edge.targetPart); }
    }
  }
  const relationships: OpcSourceRelationship[] = [], incoming: OpcSourceRelationship[] = [];
  for (const edge of edges) {
    await work(1);
    if (visited.has(edge.source)) relationships.push(edge);
    else if (edge.targetPart && selected.has(edge.targetPart)) incoming.push(edge);
  }
  let outgoingComplete = true, incomingComplete = true;
  for (const finding of findings) {
    await work(1);
    if (visited.has(finding.source)) outgoingComplete = false;
    else incomingComplete = false;
  }
  const unreachedParts: string[] = [];
  for (const part of source.inspection.parts) {
    await work(1);
    if (!selected.has(part.name)) unreachedParts.push(part.name);
  }
  abort();
  return Object.freeze({
    profile: "source-relationships-v1", roots: Object.freeze(roots),
    parts: Object.freeze([...selected.values()].sort((a, b) => order(a.name, b.name))),
    relationships: Object.freeze(relationships), incomingRelationships: Object.freeze(incoming),
    unreachedParts: Object.freeze(unreachedParts.sort(order)), findings: Object.freeze(findings),
    coverage: Object.freeze({ outgoingComplete, incomingComplete }), controlStream, expandedBytes,
    sourceRelationshipCount, effectiveRelationshipCount: source.inspection.relationships.length, workUnits,
  });
}
