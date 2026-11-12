import { inspectOpcPackage } from "@pptkit/pptx-parser";

export async function inspectPackage(bytes: Uint8Array, signal?: AbortSignal) {
  const result = await inspectOpcPackage(bytes, {
    zipLimits: { maxInputBytes: 32 * 1024 * 1024 },
    xmlLimits: { maxDepth: 128 },
    limits: { maxXmlNodes: 500000, maxRelationships: 20000 },
    ...(signal ? { signal } : {}),
  });
  return {
    parts: result.parts,
    externalReferences: result.relationships.filter((reference) => reference.targetMode === "External"),
    xmlTotals: result.xmlTotals,
  };
}
