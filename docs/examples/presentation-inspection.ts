import { inspectPresentationPackage } from "@pptkit/pptx-parser";

export async function identifyPresentation(bytes: Uint8Array) {
  const result = await inspectPresentationPackage(bytes, {
    limits: { maxRelationships: 20000 },
  });
  return {
    mainPart: result.mainPart,
    dialect: result.dialect,
    kind: result.kind,
    suggestedExtension: result.suggestedExtension,
  };
}
