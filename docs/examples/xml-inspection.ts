import { inspectXmlPart, PackageInspectionError } from "@pptkit/pptx-parser";

export async function inspectPresentationXml(bytes: Uint8Array, signal?: AbortSignal) {
  try {
    return await inspectXmlPart(bytes, {
      partName: "ppt/presentation.xml",
      limits: { maxDepth: 64, maxNodes: 20000 },
      ...(signal ? { signal } : {}),
    });
  } catch (error) {
    if (error instanceof PackageInspectionError) {
      return { code: error.code, part: error.entry, location: error.location };
    }
    throw error;
  }
}
