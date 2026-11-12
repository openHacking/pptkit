import { importPptx, type PptxTextStylePropertyInfo } from "@pptkit/pptx-parser";

export async function projectImportedTextStyles(bytes: Uint8Array) {
  const session = await importPptx(bytes, {
    projection: "content",
    sessionLimits: { maxTextStyleProperties: 500_000 },
  });
  try {
    const properties: readonly PptxTextStylePropertyInfo[] =
      session.report.textStyleInheritance?.properties ?? [];
    return properties.map(({ property, level, status, value, source, target, resolvedFrom, reason }) => ({
      property, level, status, value, source, target, resolvedFrom, reason,
    }));
  } finally {
    session.dispose();
  }
}
