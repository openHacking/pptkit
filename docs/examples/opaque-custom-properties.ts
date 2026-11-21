import { inspectPptxOpaqueCustomProperties, type OpaqueCustomPropertyValue } from "@pptkit/pptx-parser";

export async function inspectOpaqueValues(bytes: Uint8Array): Promise<readonly OpaqueCustomPropertyValue[]> {
  return inspectPptxOpaqueCustomProperties(bytes, { limits: { maxXmlTextCharacters: 500_000 } });
}
