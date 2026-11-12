import { describeImportCapabilities, importPptx } from "@pptkit/pptx-parser";

export async function inspectCapabilitySubjects(bytes: Uint8Array) {
  const session = await importPptx(bytes, { projection: "content", diagramPreview: "cache" });
  try {
    return describeImportCapabilities(session.report, {
      filter: { feature: "smartart.cache-preview" },
      limits: { maxSubjects: 100 },
    });
  } finally {
    session.dispose();
  }
}
