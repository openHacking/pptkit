import {
  compareOpcPackages,
  type OpcPackageComparison,
} from "@pptkit/pptx-parser";

export async function comparePreservedCandidate(
  original: Uint8Array,
  candidate: Uint8Array,
  signal?: AbortSignal,
): Promise<OpcPackageComparison> {
  return compareOpcPackages(original, candidate, {
    profile: "opc-payload-v1",
    comparisonLimits: {
      maxTotalExpandedBytes: 64 * 1024 * 1024,
      maxWorkUnits: 256 * 1024 * 1024,
      maxDifferences: 100,
    },
    ...(signal ? { signal } : {}),
  });
}
