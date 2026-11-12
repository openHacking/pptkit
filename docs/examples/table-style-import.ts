import { importPptx, type PptxTableStyleInheritance } from "@pptkit/pptx-parser";

export async function inspectImportedTableStyles(bytes: Uint8Array) {
  const session = await importPptx(bytes, {
    projection: "content",
    editSaveProfile: "text-runs-v3",
    sessionLimits: { maxTableStyleProperties: 50_000, maxTableStyleDependencies: 100_000, maxTableStyleWork: 500_000 },
  });
  try {
    const styles: PptxTableStyleInheritance | undefined = session.report.tableStyleInheritance;
    const selected = styles?.tables.map(table => ({ id: table.styleId, status: table.status, flags: table.flags, definition: table.definition }));
    const unsupported = styles?.properties.filter(property => property.status === "unsupported");
    // Matrix selections use the reference index as source and retain theme dependencies.
    const matrixSelections = styles?.properties.filter(property => property.source.path?.endsWith("/@idx"))
      .map(property => ({ source: property.source, status: property.status, dependencies: property.resolvedFrom }));
    const textAppearance = styles?.properties.filter(property => property.property.startsWith("run."))
      .map(property => ({ property: property.property, source: property.source, status: property.status, value: property.value, reason: property.reason, dependencies: property.resolvedFrom }));
    const prepared = await session.prepareSave({ mode: "no-op" });
    if (prepared.status !== "ready") throw new Error("Expected unchanged source eligibility.");
    const saved = await session.save(prepared.plan);
    return { selected, unsupported, matrixSelections, textAppearance, source: saved.bytes };
  } finally {
    session.dispose();
  }
}
