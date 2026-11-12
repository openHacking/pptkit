import { inspectPptx } from "@pptkit/pptx-parser";

export async function readSlideOrder(bytes: Uint8Array) {
  const result = await inspectPptx(bytes, {
    limits: { maxXmlNodes: 500000, maxRelationships: 20000 },
  });
  return {
    metadata: result.metadata,
    sections: result.sectionList,
    extensions: result.presentationExtensions,
    size: result.slideSize,
    slides: result.slides,
    unresolved: result.findings,
  };
}
