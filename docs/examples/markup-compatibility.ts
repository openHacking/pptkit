import { inspectMarkupCompatibility } from "@pptkit/pptx-parser";

export async function inspectCompatibilityPart(bytes: Uint8Array) {
  return inspectMarkupCompatibility(bytes, {
    understoodNamespaces: ["http://schemas.openxmlformats.org/presentationml/2006/main"],
    partName: "ppt/slides/slide1.xml",
    limits: { maxDepth: 256, maxNodes: 100000 },
  });
}
