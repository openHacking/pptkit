import { attribute, children } from "../presentation/structure-xml.js";
import type { structureContext } from "../presentation/structure-context.js";
import type { ImportedCustomPropertySource, PptxSessionLimits } from "../types/import.js";
import { PptxImportError } from "../types/import.js";

/** Inventory raw direct-child identities; a bound entry is not write authority. */
export async function customPropertySources(context: Awaited<ReturnType<typeof structureContext>>, limits: Readonly<PptxSessionLimits>, checkpoint: () => Promise<void>, currencyLexical?: readonly (string | undefined)[]): Promise<ImportedCustomPropertySource[]> {
  const part = context.metadata.custom;
  if (!part) return [];
  const root = context.rawXml.get(part.part), byIdentity = new Map<string, ReturnType<typeof children>>();
  const hasMce = context.reports.some(report => report.part === part.part);
  const valueNamespace = context.entry.dialect === "strict" ? "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes" : "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
  if (root) for (const node of children(root)) {
    await checkpoint();
    const pid = attribute(node, "pid"), fmtid = attribute(node, "fmtid");
    if (pid === undefined || fmtid === undefined) continue;
    const key = `${fmtid.trim()}:${Number(pid.trim())}`;
    const group = byIdentity.get(key) ?? []; group.push(node); byIdentity.set(key, group);
  }
  const rootIdentity = root ? await context.sourceElement(part.part, root) : undefined;
  const results: ImportedCustomPropertySource[] = [];
  for (const [index, property] of part.properties.entries()) {
    await checkpoint();
    if (results.length >= limits.maxReportFindings) throw new PptxImportError("import-limit-exceeded", "Custom property sources exceed maxReportFindings.", { part: part.part });
    const key = `${property.formatId}:${property.id}`, nodes = byIdentity.get(key) ?? [], node = nodes[0], values = node ? children(node) : [];
    const value = values[0], propertyIdentity = node ? await context.sourceElement(part.part, node) : undefined, valueIdentity = value ? await context.sourceElement(part.part, value) : undefined;
    const rawText = value?.children.map(child => child.kind === "text" || child.kind === "cdata" ? child.value : "").join("");
    const bound = !!(root && rootIdentity && rootIdentity.elementIndex === 1 && rootIdentity.parent === undefined && !hasMce && nodes.length === 1 && node?.namespaceUri === root.namespaceUri && node.local === "property" && propertyIdentity?.parent === root && values.length === 1 && valueIdentity?.parent === node && value?.namespaceUri === valueNamespace && value.local === property.valueType && value.children.every(child => child.kind === "text" || child.kind === "cdata") && ((property.value.status === "scalar" && rawText === property.value.value) || (property.valueType === "cy" && currencyLexical?.[index] !== undefined && rawText === currencyLexical[index])) && attribute(node, "name") === property.name && attribute(node, "linkTarget") === property.linkTarget);
    results.push({ source: { part: part.part, path: "Properties/property", sourceObjectId: key, ...(bound ? { elementIndex: propertyIdentity!.elementIndex } : {}) }, ...(bound ? { valueElementIndex: valueIdentity!.elementIndex } : {}), status: bound ? "bound" : "unmapped", ...(!bound ? { reason: "Raw direct scalar ancestry or source identity is ambiguous, altered by compatibility processing, or unavailable." } : {}) });
  }
  return results;
}
