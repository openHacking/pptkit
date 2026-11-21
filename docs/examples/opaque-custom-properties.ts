import { assessPptxCurrencyNumber, assessPptxCurrencyText, importPptx, inspectPptxOpaqueCustomProperties, type OpaqueCustomPropertyValue } from "@pptkit/pptx-parser";

export async function inspectOpaqueValues(bytes: Uint8Array): Promise<readonly OpaqueCustomPropertyValue[]> {
  return inspectPptxOpaqueCustomProperties(bytes, { limits: { maxXmlTextCharacters: 500_000 } });
}

export function currencyNumberIfExact(text: string): number | undefined {
  const result = assessPptxCurrencyNumber(text);
  return result.status === "exact" ? result.value : undefined;
}

export function inspectCurrencyText(property: OpaqueCustomPropertyValue) {
  return property.valueType === "cy" && property.content.status === "text"
    ? assessPptxCurrencyText(property.content.value)
    : undefined;
}

export async function changeExactCurrency(bytes: Uint8Array, name: string, value: number): Promise<Uint8Array> {
  const session = await importPptx(bytes, { projection: "content", customCurrencyProjection: "exact-save-v2", customPropertySaveProfile: "scalar-v1" });
  try {
    const properties = session.document.metadata.customProperties ?? [];
    const index = properties.findIndex(property => property.name === name && property.type === "real");
    const source = session.report.customPropertySources?.find(item => item.projectedIndex === index);
    const currency = session.report.inspection.metadata.custom?.properties.find(item => `${item.formatId}:${item.id}` === source?.source.sourceObjectId && item.valueType === "cy");
    if (index < 0 || !currency || !session.document.setMetadata) throw new Error("Bound currency property not found.");
    session.document.setMetadata({ ...session.document.metadata, customProperties: properties.map(property => property.name === name && property.type === "real" ? { ...property, value } : property) });
    const prepared = await session.prepareSave({ mode: "edit" });
    if (prepared.status !== "ready") throw new Error("Source edit was rejected; inspect the preparation report.");
    return (await session.save(prepared.plan)).bytes;
  } finally { session.dispose(); }
}
