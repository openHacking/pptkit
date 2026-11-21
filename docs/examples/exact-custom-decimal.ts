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
