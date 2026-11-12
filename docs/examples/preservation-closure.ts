import {
  inspectOpcClosure,
  type OpcClosureInspection,
} from "@pptkit/pptx-parser";

export async function inspectSlideDependencies(
  source: Uint8Array,
  slidePart: string,
  signal?: AbortSignal,
): Promise<OpcClosureInspection> {
  return inspectOpcClosure(source, {
    roots: [slidePart],
    closureLimits: {
      maxSourceRelationships: 10000,
      maxClosureParts: 5000,
      maxClosureExpandedBytes: 64 * 1024 * 1024,
    },
    ...(signal ? { signal } : {}),
  });
}
