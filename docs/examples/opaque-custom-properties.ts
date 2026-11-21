import { assessPptxCurrencyText, inspectPptxOpaqueCustomProperties, type OpaqueCustomPropertyValue } from "@pptkit/pptx-parser";

export async function inspectOpaqueValues(bytes: Uint8Array): Promise<readonly OpaqueCustomPropertyValue[]> {
  return inspectPptxOpaqueCustomProperties(bytes, { limits: { maxXmlTextCharacters: 500_000 } });
}

export function inspectCurrencyText(property: OpaqueCustomPropertyValue) {
  return property.valueType === "cy" && property.content.status === "text"
    ? assessPptxCurrencyText(property.content.value)
    : undefined;
}
