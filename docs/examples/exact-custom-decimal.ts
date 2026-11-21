import { createPresentation } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";
import { importPptx } from "@pptkit/pptx-parser";

export async function makeExactDecimal(): Promise<Uint8Array> {
  const document = createPresentation({ metadata: { customProperties: [{ name: "Exact", type: "decimal", value: "9007199254740993.0001" }] } });
  return (await generatePptx(document)).bytes;
}

export async function readExactDecimals(bytes: Uint8Array): Promise<readonly string[]> {
  const session = await importPptx(bytes, { projection: "content", customDecimalProjection: "exact-v1" });
  try { return (session.document.metadata.customProperties ?? []).filter(property => property.type === "decimal").map(property => property.value); }
  finally { session.dispose(); }
}

export async function changeBoundDecimal(bytes: Uint8Array, name: string, value: string): Promise<Uint8Array> {
  const session = await importPptx(bytes, { projection: "content", customDecimalProjection: "exact-save-v2", customPropertySaveProfile: "scalar-v1" });
  try {
    const properties = session.document.metadata.customProperties ?? [];
    const index = properties.findIndex(property => property.name === name && property.type === "decimal");
    if (index < 0 || !session.report.customPropertySources?.some(source => source.projectedIndex === index && source.status === "bound") || !session.document.setMetadata) throw new Error("Bound decimal property not found.");
    session.document.setMetadata({ ...session.document.metadata, customProperties: properties.map(property => property.name === name && property.type === "decimal" ? { ...property, value } : property) });
    const prepared = await session.prepareSave({ mode: "edit" });
    if (prepared.status !== "ready") throw new Error("Decimal source edit was rejected; inspect its report.");
    return (await session.save(prepared.plan)).bytes;
  } finally { session.dispose(); }
}
