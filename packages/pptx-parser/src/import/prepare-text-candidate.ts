import { planPropertyCreation } from "../preservation/plan-property-creation.js";
import { appendPropertyDeclarations } from "../preservation/append-property-declarations.js";
import { appendXmlRootChild } from "../preservation/append-xml-root-child.js";
import { planCustomPropertyCreation } from "../preservation/plan-custom-property-creation.js";
import { readXml } from "../xml/read-xml.js";
import { removeXmlRootChildren } from "../preservation/remove-xml-root-children.js";
import type { NormalizedPresentation } from "@pptkit/core";
import { readOpcSource } from "../opc/read-package.js";
import { inspectClosure } from "../preservation/inspect-closure.js";
import { captureXmlTextSource, XmlSourceEditError, type XmlTextSourceElement } from "../preservation/xml-text-source.js";
import { editXmlNumericAttribute, editXmlTextAttributes } from "../preservation/replace-xml-relationship-id.js";
import { replaceArchiveEntries } from "../preservation/replace-archive-entries.js";
import { comparePackages } from "../comparison/compare-packages.js";
import { structureContext } from "../presentation/structure-context.js";
import { readStructureFromContext } from "../presentation/read-structure.js";
import { PackageInspectionError } from "../types/archive.js";
import { PptxImportError, type ImportDiagnostic, type PptxSessionLimits, type PptxImportOptions } from "../types/import.js";
import type { PptxInspectionOptions } from "../types/structure.js";
import { equalState, snapshot } from "./state.js";
import { projectStructure } from "./project-structure.js";
import { projectCurrencies } from "./currency-projection.js";
import { projectContent } from "./project-content.js";
import { chartNumber } from "./chart-data.js";
import type { ChartValueBinding } from "./chart-value-bindings.js";
import type { ChartLabelBinding } from "./chart-label-bindings.js";
import { verifyChartSeriesOrderRebound, type ChartSeriesOrderBinding } from "./chart-series-order.js";
import { captureColumnLegendSources } from "./chart-legend-source.js";
import { verifyChartOrderRelationshipClosure } from "./chart-order-closure.js";
import { chartString } from "./chart-string.js";
import { prepareChartValues } from "./prepare-chart-values.js";
import type { BoundXmlTextEdit, CoreMetadataBinding, TextEditAnalysis } from "./analyze-text-edits.js";

/** Cross-check optional raw ancestry evidence against the separately captured lexical tree. */
export async function verifyTextSourceAncestry(elements: readonly XmlTextSourceElement[], edits: readonly BoundXmlTextEdit[], part: string, checkpoint: () => Promise<void>): Promise<void> {
  for (const edit of edits) {
    await checkpoint();
    const atom = elements[edit.elementIndex - 1];
    if (edit.rawMetadataRootElementIndex !== undefined) {
      const root = elements[edit.rawMetadataRootElementIndex - 1];
      const extended = edit.rawExtendedRootNamespace;
      if (!atom || atom.parentElementIndex !== edit.rawMetadataRootElementIndex || !root || root.elementIndex !== 1 || root.parentElementIndex !== undefined || root.namespaceUri !== (extended ?? "http://schemas.openxmlformats.org/package/2006/metadata/core-properties") || root.local !== (extended ? "Properties" : "coreProperties") || extended && (atom.namespaceUri !== extended || atom.local !== "Company")) throw new PptxImportError("save-candidate-invalid", "Metadata source ancestry disagrees with its binding.", { part });
    }
    if (edit.rawCustomPropertyElementIndex !== undefined) {
      const property = elements[edit.rawCustomPropertyElementIndex - 1], root = elements[0];
      if (!atom || atom.parentElementIndex !== edit.rawCustomPropertyElementIndex || !property || property.local !== "property" || property.namespaceUri !== edit.rawCustomRootNamespace || property.parentElementIndex !== 1 || !root || root.local !== "Properties" || root.namespaceUri !== edit.rawCustomRootNamespace || root.parentElementIndex !== undefined) throw new PptxImportError("save-candidate-invalid", "Custom property source ancestry disagrees with its binding.", { part });
    }
    if (edit.rawParagraphElementIndex !== undefined && edit.rawRunElementIndex === undefined) throw new PptxImportError("save-candidate-invalid", "Raw paragraph evidence lacks its run identity.", { part });
    if (edit.rawRunElementIndex === undefined) continue;
    const run = elements[edit.rawRunElementIndex - 1];
    if (!atom || atom.parentElementIndex !== edit.rawRunElementIndex || !run || run.local !== "r" || run.namespaceUri !== edit.namespaceUri) throw new PptxImportError("save-candidate-invalid", "Raw text/run ancestry disagrees with the source binding.", { part });
    if (edit.rawParagraphElementIndex !== undefined) {
      const paragraph = elements[edit.rawParagraphElementIndex - 1];
      if (run.parentElementIndex !== edit.rawParagraphElementIndex || !paragraph || paragraph.local !== "p" || paragraph.namespaceUri !== edit.namespaceUri) throw new PptxImportError("save-candidate-invalid", "Raw run/paragraph ancestry disagrees with the source binding.", { part });
    }
  }
}

export async function prepareTextCandidate(
  source: Uint8Array, candidate: NormalizedPresentation, analysis: TextEditAnalysis,
  options: PptxInspectionOptions, limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>, diagramPreview?: PptxImportOptions["diagramPreview"], textFontProfile?: PptxImportOptions["textFontProfile"], notesTextProfile?: PptxImportOptions["notesTextProfile"], fieldTextProfile?: PptxImportOptions["fieldTextProfile"], currencySave?: boolean,
): Promise<{ bytes: Uint8Array; validationBytes: number; validationWork: number } | { diagnostics: readonly ImportDiagnostic[] }> {
  try {
    await checkpoint();
    const coreCreating = analysis.metadataInsertions?.find(item => item.binding.create && !item.binding.extendedRootNamespace)?.binding;
    const extendedCreating = analysis.metadataInsertions?.find(item => item.binding.create && item.binding.extendedRootNamespace)?.binding;
    const creating = analysis.customInsertions?.find(item => item.source.create)?.source;
    let orderClosureBytes = 0, orderClosureWork = 0;
    const orderParts = new Set<string>();
    for (const proposal of analysis.chartOrders ?? []) {
      await checkpoint();
      if (orderParts.has(proposal.binding.part)) throw new PptxImportError("save-candidate-invalid", "Chart order source part is repeated.", { part: proposal.binding.part });
      orderParts.add(proposal.binding.part);
      const remainingBytes = limits.maxSaveValidationBytes - orderClosureBytes, remainingWork = limits.maxSaveValidationWork - orderClosureWork;
      if (remainingBytes <= 0 || remainingWork <= 0) throw new PptxImportError("save-limit-exceeded", "Chart order relationship scans exceed the shared validation budget.");
      const graph = await inspectClosure(source, { ...options, roots: [proposal.binding.part], closureLimits: { maxRoots: 1, maxClosureExpandedBytes: remainingBytes, maxWorkUnits: remainingWork, maxFindings: limits.maxSaveChanges } });
      await verifyChartOrderRelationshipClosure(proposal.binding, graph, checkpoint);
      orderClosureBytes += graph.expandedBytes; orderClosureWork += graph.workUnits + graph.relationships.length + 1;
      if (orderClosureBytes > limits.maxSaveValidationBytes || orderClosureWork > limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Chart order relationship scans exceed the shared validation budget.");
    }
    const closure = await inspectClosure(source, { ...options, roots: [...new Set([...analysis.parts.keys(), ...(analysis.metadataInsertions ?? []).filter(item => !item.binding.create).map(item => item.binding.part), ...(analysis.metadataRemovals ?? []).map(item => item.part), ...(analysis.firstSlideNumber ? [analysis.firstSlideNumber.part] : []), ...(analysis.descriptions ?? []).map(item => item.binding.part), ...(analysis.slideVisibility ?? []).map(item => item.binding.part), ...(analysis.customNames ?? []).map(item => item.binding.part), ...(analysis.customRemovals ?? []).map(item => item.part), ...(analysis.customInsertions ?? []).filter(item => !item.source.create).map(item => item.source.part), ...(creating || coreCreating || extendedCreating ? ["/"] : []), ...analysis.chartValues.flatMap(v => [v.binding.source.part, ...(v.binding.workbook ? [v.binding.workbook.part] : [])]), ...(analysis.chartLabels ?? []).flatMap(v => v.binding.closure.uses.flatMap(use => [use.cache.part, use.cell.workbookPart])), ...(analysis.chartOrders ?? []).map(v => v.binding.part)])], closureLimits: { maxFindings: limits.maxSaveChanges } });
    await checkpoint();
    if (closure.findings.length) return { diagnostics: closure.findings.map(finding => ({
      code: "save-source-closure-incomplete", severity: "error" as const, source: { part: finding.relationshipPart, elementIndex: finding.elementIndex }, message: finding.reason,
    })) };
    const original = await readOpcSource(source, options, false, true);
    const partByName = new Map(original.inspection.parts.map(part => [part.name, part]));
    const replacements: { name: string; bytes: Uint8Array }[] = [];
    const replacementByName = new Map<string, { name: string; bytes: Uint8Array }>();
    const additions: { name: string; bytes: Uint8Array }[] = [];
    let replacementBytes = 0, descriptionValidationBytes = orderClosureBytes, descriptionValidationWork = orderClosureWork;
    for (const [part, edits] of [...analysis.parts].sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0)) {
      await checkpoint();
      const inventory = partByName.get(part);
      const payload = inventory && original.payloads!.get(inventory.entryName);
      if (!inventory || !payload) throw new PptxImportError("save-candidate-invalid", "Bound source text part is unavailable.", { part });
      const xml = await captureXmlTextSource(payload, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part });
      try {
        await verifyTextSourceAncestry(xml.elements, edits, part, checkpoint);
        const changed = await xml.replaceText(edits, {
          ...(options.signal ? { signal: options.signal } : {}),
          limits: { maxEdits: limits.maxSaveChanges, maxOutputBytes: limits.maxPreparedOutputBytes },
        });
        await checkpoint();
        replacementBytes += changed.bytes.length;
        if (replacementBytes > limits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Prepared XML payloads exceed maxPreparedOutputBytes.");
        const replacement = { name: inventory.entryName, bytes: changed.bytes }; replacements.push(replacement); replacementByName.set(replacement.name, replacement);
      } catch (cause) {
        if (cause instanceof XmlSourceEditError && cause.code === "xml-edit-unsupported") return { diagnostics: [{ code: edits.some(edit => edit.rawMetadataRootElementIndex !== undefined) ? "save-metadata-candidate-unsupported" : "save-text-candidate-unsupported", severity: "error", source: { part }, message: cause.message }] };
        throw cause;
      } finally { xml.dispose(); }
    }
    const coreCreationPlan = coreCreating ? await planPropertyCreation(source, "core", coreCreating.create!.dialect, options, limits.maxPreparedOutputBytes, limits.maxSaveChanges) : undefined;
    const extendedCreationPlan = extendedCreating ? await planPropertyCreation(source, "extended", extendedCreating.create!.dialect, options, limits.maxPreparedOutputBytes, limits.maxSaveChanges) : undefined;
    let finalDeclarations: { relationships: Uint8Array; contentTypes: Uint8Array } | undefined;
    const metadataInsertionsByPart = new Map<string, NonNullable<TextEditAnalysis["metadataInsertions"]>[number][]>();
    for (const item of analysis.metadataInsertions ?? []) { const list = metadataInsertionsByPart.get(item.binding.part) ?? []; list.push(item); metadataInsertionsByPart.set(item.binding.part, list); }
    for (const [part, insertions] of metadataInsertionsByPart) {
      await checkpoint();
      const inventory = partByName.get(part), payload = inventory && original.payloads!.get(inventory.entryName);
      const create = insertions[0]!.binding.create;
      if (create ? Boolean(inventory) : !inventory || !payload) throw new PptxImportError("save-candidate-invalid", "Core metadata insertion part is unavailable or already exists.", { part });
      if (insertions.some(item => Boolean(item.binding.create) !== Boolean(create) || create && (item.binding.create?.relationshipId !== create.relationshipId || item.binding.create?.dialect !== create.dialect))) throw new PptxImportError("save-candidate-invalid", "Core metadata creation bindings disagree.", { part });
      const existing = inventory && replacementByName.get(inventory.entryName), current = existing?.bytes ?? payload ?? new TextEncoder().encode(insertions[0]!.binding.extendedRootNamespace ? `<Properties xmlns="${insertions[0]!.binding.extendedRootNamespace}"/>` : '<coreProperties xmlns="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>');
      const xmlOptions = { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part };
      const parsed = await readXml(current, xmlOptions, true), root = parsed.root;
      const ns = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", dc = "http://purl.org/dc/elements/1.1/", terms = "http://purl.org/dc/terms/", xsi = "http://www.w3.org/2001/XMLSchema-instance";
      const extendedNamespace = insertions[0]!.binding.extendedRootNamespace;
      if (!root || root.namespaceUri !== (extendedNamespace ?? ns) || root.local !== (extendedNamespace ? "Properties" : "coreProperties") || insertions.some(item => !item.binding.insert || item.binding.rootElementIndex !== 1 || item.binding.extendedRootNamespace !== extendedNamespace || !(extendedNamespace ? item.binding.namespace === extendedNamespace && item.binding.local === "Company" : (item.binding.namespace === dc && ["title", "creator", "subject", "description", "language"].includes(item.binding.local)) || (item.binding.namespace === terms && ["created", "modified"].includes(item.binding.local)) || (item.binding.namespace === ns && ["revision", "keywords"].includes(item.binding.local)))) || new Set(insertions.map(item => item.binding.local)).size !== insertions.length) throw new PptxImportError("save-candidate-invalid", "Metadata insertion binding is invalid.", { part });
      let child = "", work = current.length;
      for (const { binding, value } of insertions) {
        if (root.children.some(node => node.kind === "element" && node.namespaceUri === binding.namespace && node.local === binding.local)) throw new PptxImportError("save-candidate-invalid", "Core metadata insertion target already exists.", { part });
        work += value.length * 6 + 256;
        if (work > limits.maxPreparedOutputBytes || work > limits.maxSaveValidationWork - descriptionValidationWork) throw new PptxImportError("save-limit-exceeded", "Core metadata insertion exceeds its serialization budget.");
        let count = 0;
        for (const character of value) { if (++count % 16384 === 0) await checkpoint(); const code = character.codePointAt(0)!; if (code !== 9 && code !== 10 && code !== 13 && (code < 32 || code >= 0xd800 && code <= 0xdfff || code === 0xfffe || code === 0xffff)) return { diagnostics: [{ code: "save-metadata-candidate-unsupported", severity: "error", source: { part }, message: "Core metadata contains unsupported XML characters." }] }; }
        const text = value.replace(/[&<>\r]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\r": "&#13;" })[c]!);
        child += extendedNamespace ? `<ep:Company xmlns:ep="${extendedNamespace}">${text}</ep:Company>` : binding.namespace === ns ? `<cp:${binding.local} xmlns:cp="${ns}">${text}</cp:${binding.local}>` : binding.namespace === terms ? `<dct:${binding.local} xmlns:dct="${terms}" xmlns:xsi="${xsi}" xsi:type="dct:W3CDTF">${text}</dct:${binding.local}>` : `<dc:${binding.local} xmlns:dc="${dc}">${text}</dc:${binding.local}>`;
      }
      descriptionValidationBytes += current.length; descriptionValidationWork += work;
      if (descriptionValidationBytes >= limits.maxSaveValidationBytes || descriptionValidationWork >= limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Core metadata insertion exceeds its validation budget.");
      const changed = await appendXmlRootChild(current, root, parsed.inspection.encoding, child, limits.maxPreparedOutputBytes, xmlOptions);
      for (const { binding, value } of insertions) {
        const matches = changed.root.children.filter(node => node.kind === "element" && node.namespaceUri === binding.namespace && node.local === binding.local);
        const target = matches[0];
        if (matches.length !== 1 || !target || target.kind !== "element" || (binding.namespace === terms && !target.attributes.some(attr => attr.namespaceUri === xsi && attr.local === "type" && attr.value === "dct:W3CDTF")) || target.children.some(node => node.kind !== "text" && node.kind !== "cdata") || target.children.map(node => node.kind === "text" || node.kind === "cdata" ? node.value : "").join("") !== value) throw new PptxImportError("save-candidate-invalid", "Core metadata insertion failed verification.", { part });
      }
      replacementBytes += changed.bytes.length - (existing?.bytes.length ?? 0);
      if (replacementBytes > limits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Prepared XML payloads exceed maxPreparedOutputBytes.");
      if (create) {
        const plan = extendedNamespace ? extendedCreationPlan : coreCreationPlan, kind = extendedNamespace ? "extended" : "core";
        if (!plan || plan.part !== part || plan.relationshipId !== create.relationshipId) throw new PptxImportError("save-candidate-invalid", "Property creation identity changed.", { part });
        additions.push({ name: part.slice(1), bytes: changed.bytes });
        finalDeclarations = finalDeclarations ? await appendPropertyDeclarations(finalDeclarations.relationships, finalDeclarations.contentTypes, { kind, part, relationshipId: plan.relationshipId, dialect: create.dialect }, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), maxOutputBytes: limits.maxPreparedOutputBytes }) : plan;
        for (const [name, bytes] of [["_rels/.rels", finalDeclarations.relationships], ["[Content_Types].xml", finalDeclarations.contentTypes]] as const) { const existingDeclaration = replacementByName.get(name); if (existingDeclaration) { replacementBytes += bytes.length - existingDeclaration.bytes.length; existingDeclaration.bytes = bytes; } else { replacementBytes += bytes.length; const replacement = { name, bytes }; replacements.push(replacement); replacementByName.set(name, replacement); } }
        if (replacementBytes > limits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Core creation declarations exceed the output budget.");
      } else if (existing) existing.bytes = changed.bytes; else { const replacement = { name: inventory!.entryName, bytes: changed.bytes }; replacements.push(replacement); replacementByName.set(replacement.name, replacement); }
    }
    const metadataRemovalsByPart = new Map<string, CoreMetadataBinding[]>();
    for (const binding of analysis.metadataRemovals ?? []) { const list = metadataRemovalsByPart.get(binding.part) ?? []; list.push(binding); metadataRemovalsByPart.set(binding.part, list); }
    for (const [part, removals] of metadataRemovalsByPart) {
      await checkpoint();
      const inventory = partByName.get(part), payload = inventory && original.payloads!.get(inventory.entryName), existing = inventory && replacementByName.get(inventory.entryName);
      if (!inventory || !payload || removals.some(binding => !binding.removable || binding.insert || binding.rootElementIndex !== 1 || binding.part !== part)) throw new PptxImportError("save-candidate-invalid", "Metadata removal source is unavailable or invalid.", { part });
      const extended = removals[0]!.extendedRootNamespace, rootNamespace = extended ?? "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", rootLocal = extended ? "Properties" : "coreProperties";
      if (removals.some(binding => binding.extendedRootNamespace !== extended || (extended ? binding.namespace !== extended || binding.local !== "Company" : !((binding.namespace === "http://purl.org/dc/elements/1.1/" && ["subject", "description", "title", "creator", "language"].includes(binding.local)) || (binding.namespace === "http://purl.org/dc/terms/" && ["created", "modified"].includes(binding.local)) || (binding.namespace === "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" && ["keywords", "revision"].includes(binding.local)))))) throw new PptxImportError("save-candidate-invalid", "Metadata removal identity disagrees with its profile.", { part });
      const current = existing?.bytes ?? payload;
      descriptionValidationBytes += current.length; descriptionValidationWork += current.length;
      if (descriptionValidationBytes >= limits.maxSaveValidationBytes || descriptionValidationWork >= limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Metadata removal exceeds its validation budget.");
      const changed = await removeXmlRootChildren(current, removals.map(binding => ({ elementIndex: binding.elementIndex, namespaceUri: binding.namespace, local: binding.local })), { namespaceUri: rootNamespace, local: rootLocal }, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part, maxOutputBytes: limits.maxPreparedOutputBytes, maxEdits: limits.maxSaveChanges });
      replacementBytes += changed.length - (existing?.bytes.length ?? 0);
      if (replacementBytes > limits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Prepared XML payloads exceed maxPreparedOutputBytes.");
      if (existing) existing.bytes = changed; else { const replacement = { name: inventory.entryName, bytes: changed }; replacements.push(replacement); replacementByName.set(replacement.name, replacement); }
    }
    const descriptionsByPart = new Map<string, NonNullable<TextEditAnalysis['descriptions']>[number][]>();
    for (const edit of analysis.descriptions ?? []) { const list = descriptionsByPart.get(edit.binding.part) ?? []; list.push(edit); descriptionsByPart.set(edit.binding.part, list); }
    const visibilityByPart = new Map<string, NonNullable<TextEditAnalysis['slideVisibility']>[number]>();
    for (const item of analysis.slideVisibility ?? []) { const prior = visibilityByPart.get(item.binding.part); if (prior && prior.hidden !== item.hidden) throw new PptxImportError('save-candidate-invalid', 'Shared slide visibility proposals disagree.', { part: item.binding.part }); visibilityByPart.set(item.binding.part, item); }
    const drawingPropertyParts = [...new Set([...descriptionsByPart.keys(), ...visibilityByPart.keys()])].sort();
    for (const part of drawingPropertyParts) {
      const edits = descriptionsByPart.get(part) ?? [], visibility = visibilityByPart.get(part);
      const inventory = partByName.get(part), payload = inventory && original.payloads!.get(inventory.entryName);
      if (!inventory || !payload) throw new PptxImportError('save-candidate-invalid', 'Bound drawing property source is unavailable.', { part });
      descriptionValidationBytes += payload.length; descriptionValidationWork += payload.length;
      if (descriptionValidationBytes >= limits.maxSaveValidationBytes || descriptionValidationWork >= limits.maxSaveValidationWork) throw new PptxImportError('save-limit-exceeded', 'Description source inspection exceeds the shared validation budget.');
      const xml = await captureXmlTextSource(payload, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part });
      try {
        if (visibility) {
          const root = xml.elements[visibility.binding.elementIndex - 1];
          if (!root || root.elementIndex !== 1 || root.parentElementIndex !== undefined || root.local !== 'sld' || root.namespaceUri !== visibility.binding.namespace) throw new PptxImportError('save-candidate-invalid', 'Slide visibility source root disagrees with its binding.', { part });
        }
        for (const { binding, value } of edits) {
          await checkpoint();
          if (binding.status !== 'bound' || binding.elementIndex === undefined || binding.nonvisualIndex === undefined || binding.ownerIndex === undefined || binding.expected === undefined && !binding.allowPresenceChange || value === undefined && !binding.allowPresenceChange) throw new PptxImportError('save-candidate-invalid', 'Bound drawing property source is unavailable.', { part });
          const property = xml.elements[(binding.property === 'decorative' ? binding.propertyIndex! : binding.elementIndex) - 1], nonvisual = xml.elements[binding.nonvisualIndex - 1], owner = xml.elements[binding.ownerIndex - 1];
          const containers: Record<string, string> = { sp: 'nvSpPr', grpSp: 'nvGrpSpPr', pic: 'nvPicPr', cxnSp: 'nvCxnSpPr', graphicFrame: 'nvGraphicFramePr' };
          if (!property || property.local !== 'cNvPr' || property.namespaceUri !== binding.namespace || property.parentElementIndex !== binding.nonvisualIndex || !nonvisual || nonvisual.namespaceUri !== binding.namespace || nonvisual.parentElementIndex !== binding.ownerIndex || !owner || owner.namespaceUri !== binding.namespace || containers[owner.local] !== nonvisual.local) throw new PptxImportError('save-candidate-invalid', 'Drawing property source ancestry disagrees with its binding.', { part });
          if (binding.property === 'decorative') {
            const list = xml.elements[(binding.listIndex ?? 0) - 1], extension = xml.elements[(binding.extensionIndex ?? 0) - 1], target = xml.elements[binding.elementIndex - 1];
            if (!list || list.namespaceUri !== 'http://schemas.openxmlformats.org/drawingml/2006/main' && list.namespaceUri !== 'http://purl.oclc.org/ooxml/drawingml/main' || list.local !== 'extLst' || list.parentElementIndex !== binding.propertyIndex || !extension || extension.namespaceUri !== list.namespaceUri || extension.local !== 'ext' || extension.parentElementIndex !== binding.listIndex || !target || target.namespaceUri !== 'http://schemas.microsoft.com/office/drawing/2017/decorative' || target.local !== 'decorative' || target.parentElementIndex !== binding.extensionIndex) throw new PptxImportError('save-candidate-invalid', 'Decorative source ancestry disagrees with its binding.', { part });
          }
        }
      } finally { xml.dispose(); }
      const existing = replacementByName.get(inventory.entryName);
      const current = existing?.bytes ?? payload;
      const requests = edits.map(({ binding, value }) => ({ elementIndex: binding.elementIndex!, elementNamespace: binding.property === 'decorative' ? 'http://schemas.microsoft.com/office/drawing/2017/decorative' : binding.namespace, elementLocal: binding.property === 'decorative' ? 'decorative' : 'cNvPr', attributeNamespace: '', attributeLocal: binding.property === 'decorative' ? 'val' : binding.property ?? 'descr', ...(binding.expected !== undefined ? { expected: binding.expected } : {}), ...(value !== undefined ? { replacement: value } : {}) }));
      if (visibility) requests.push({ elementIndex: visibility.binding.elementIndex, elementNamespace: visibility.binding.namespace, elementLocal: 'sld', attributeNamespace: '', attributeLocal: 'show', ...(visibility.binding.expected !== undefined ? { expected: visibility.binding.expected } : {}), replacement: visibility.hidden ? '0' : '1' });
      await checkpoint();
      const result = await editXmlTextAttributes(current, requests, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part, maxOutputBytes: limits.maxPreparedOutputBytes, maxEdits: limits.maxSaveChanges, maxWorkUnits: limits.maxSaveValidationWork - descriptionValidationWork });
      const changed = result.bytes; descriptionValidationWork += result.workUnits;
      if (descriptionValidationWork >= limits.maxSaveValidationWork) throw new PptxImportError('save-limit-exceeded', 'Drawing property edits exceed the shared validation work budget.');
      replacementBytes += changed.length - (existing?.bytes.length ?? 0);
      if (replacementBytes > limits.maxPreparedOutputBytes) throw new PptxImportError('save-limit-exceeded', 'Prepared XML payloads exceed maxPreparedOutputBytes.');
      if (existing) existing.bytes = changed; else { const replacement = { name: inventory.entryName, bytes: changed }; replacements.push(replacement); replacementByName.set(replacement.name, replacement); }
    }
    if (analysis.firstSlideNumber) {
      if (analysis.parts.has(analysis.firstSlideNumber.part)) return { diagnostics: [{ code: "save-text-candidate-unsupported", severity: "error", source: { part: analysis.firstSlideNumber.part }, message: "Text and presentation attribute edits sharing one source part are not supported." }] };
      const binding = analysis.firstSlideNumber, inventory = partByName.get(binding.part), payload = inventory && original.payloads!.get(inventory.entryName);
      if (!inventory || !payload) throw new PptxImportError("save-candidate-invalid", "Bound presentation part is unavailable.", { part: binding.part });
      const changed = await editXmlNumericAttribute(payload, { elementIndex: binding.elementIndex, elementNamespace: binding.namespace, elementLocal: "presentation", attributeNamespace: "", attributeLocal: "firstSlideNum", ...(binding.expected !== undefined ? { expected: binding.expected } : {}), ...(binding.value !== undefined ? { replacement: String(binding.value) } : {}) }, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: binding.part, maxOutputBytes: limits.maxPreparedOutputBytes });
      replacementBytes += changed.length;
      if (replacementBytes > limits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Prepared XML payloads exceed maxPreparedOutputBytes.");
      replacements.push({ name: inventory.entryName, bytes: changed });
    }
    const customNamesByPart = new Map<string, NonNullable<TextEditAnalysis["customNames"]>[number][]>();
    for (const item of analysis.customNames ?? []) { const list = customNamesByPart.get(item.binding.part) ?? []; list.push(item); customNamesByPart.set(item.binding.part, list); }
    for (const [part, edits] of customNamesByPart) {
      await checkpoint();
      const inventory = partByName.get(part), payload = inventory && original.payloads!.get(inventory.entryName);
      if (!inventory || !payload) throw new PptxImportError("save-candidate-invalid", "Bound custom property part is unavailable.", { part });
      descriptionValidationBytes += payload.length; descriptionValidationWork += payload.length;
      if (descriptionValidationBytes >= limits.maxSaveValidationBytes || descriptionValidationWork >= limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Custom property inspection exceeds the shared validation budget.");
      const xml = await captureXmlTextSource(payload, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part });
      try {
        await verifyTextSourceAncestry(xml.elements, edits.map(({ binding }) => ({ elementIndex: binding.valueElementIndex, namespaceUri: binding.valueNamespace, local: binding.local, expectedText: binding.lexical, text: binding.lexical, rawCustomPropertyElementIndex: binding.propertyElementIndex, rawCustomRootNamespace: binding.rootNamespace })), part, checkpoint);
      } finally { xml.dispose(); }
      const existing = replacementByName.get(inventory.entryName), current = existing?.bytes ?? payload;
      const requests = edits.map(({ binding, value }) => ({ elementIndex: binding.propertyElementIndex, elementNamespace: binding.rootNamespace, elementLocal: "property", attributeNamespace: "", attributeLocal: "name", expected: binding.lexicalName!, replacement: value }));
      const changed = await editXmlTextAttributes(current, requests, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part, maxOutputBytes: limits.maxPreparedOutputBytes, maxEdits: limits.maxSaveChanges, maxWorkUnits: limits.maxSaveValidationWork - descriptionValidationWork });
      descriptionValidationWork += changed.workUnits;
      if (descriptionValidationWork >= limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Custom property edits exceed the shared validation budget.");
      replacementBytes += changed.bytes.length - (existing?.bytes.length ?? 0);
      if (replacementBytes > limits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Prepared XML payloads exceed maxPreparedOutputBytes.");
      if (existing) existing.bytes = changed.bytes; else { const replacement = { name: inventory.entryName, bytes: changed.bytes }; replacements.push(replacement); replacementByName.set(replacement.name, replacement); }
    }
    const customRemovalsByPart = new Map<string, NonNullable<TextEditAnalysis["customRemovals"]>[number][]>();
    for (const binding of analysis.customRemovals ?? []) { const list = customRemovalsByPart.get(binding.part) ?? []; list.push(binding); customRemovalsByPart.set(binding.part, list); }
    for (const [part, removals] of customRemovalsByPart) {
      await checkpoint();
      const inventory = partByName.get(part), payload = inventory && original.payloads!.get(inventory.entryName);
      if (!inventory || !payload) throw new PptxImportError("save-candidate-invalid", "Bound custom property part is unavailable.", { part });
      descriptionValidationBytes += payload.length; descriptionValidationWork += payload.length;
      if (descriptionValidationBytes >= limits.maxSaveValidationBytes || descriptionValidationWork >= limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Custom removal inspection exceeds the shared validation budget.");
      const xml = await captureXmlTextSource(payload, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part });
      try {
        await verifyTextSourceAncestry(xml.elements, removals.map(binding => ({ elementIndex: binding.valueElementIndex, namespaceUri: binding.valueNamespace, local: binding.local, expectedText: binding.lexical, text: binding.lexical, rawCustomPropertyElementIndex: binding.propertyElementIndex, rawCustomRootNamespace: binding.rootNamespace })), part, checkpoint);
        for (const binding of removals) {
          await checkpoint(); const atom = xml.elements[binding.valueElementIndex - 1];
          if (!atom || atom.namespaceUri !== binding.valueNamespace || atom.local !== binding.local || atom.text !== binding.lexical) throw new PptxImportError("save-candidate-invalid", "Custom removal source value disagrees with its binding.", { part });
        }
      } finally { xml.dispose(); }
      const existing = replacementByName.get(inventory.entryName), current = existing?.bytes ?? payload, rootNamespace = removals[0]!.rootNamespace;
      if (removals.some(binding => binding.rootNamespace !== rootNamespace)) throw new PptxImportError("save-candidate-invalid", "Custom removal root namespaces disagree.", { part });
      const changed = await removeXmlRootChildren(current, removals.map(binding => ({ elementIndex: binding.propertyElementIndex, namespaceUri: binding.rootNamespace, local: "property" })), { namespaceUri: rootNamespace, local: "Properties" }, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part, maxOutputBytes: limits.maxPreparedOutputBytes, maxEdits: limits.maxSaveChanges });
      replacementBytes += changed.length - (existing?.bytes.length ?? 0);
      if (replacementBytes > limits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Prepared XML payloads exceed maxPreparedOutputBytes.");
      if (existing) existing.bytes = changed; else { const replacement = { name: inventory.entryName, bytes: changed }; replacements.push(replacement); replacementByName.set(replacement.name, replacement); }
    }
    const customInsertionsByPart = new Map<string, NonNullable<TextEditAnalysis["customInsertions"]>[number][]>();
    for (const item of analysis.customInsertions ?? []) { const list = customInsertionsByPart.get(item.source.part) ?? []; list.push(item); customInsertionsByPart.set(item.source.part, list); }
    const creationPlan = creating ? await planCustomPropertyCreation(source, creating.create!.dialect, options, limits.maxPreparedOutputBytes, limits.maxSaveChanges) : undefined;
    for (const [part, insertions] of customInsertionsByPart) {
      await checkpoint();
      const source = insertions[0]!.source, inventory = partByName.get(part), payload = inventory && original.payloads!.get(inventory.entryName);
      if (source.create ? Boolean(inventory) : !inventory || !payload) throw new PptxImportError("save-candidate-invalid", "Custom insertion source part is unavailable or already exists.", { part });
      const existing = inventory && replacementByName.get(inventory.entryName), current = existing?.bytes ?? payload ?? new TextEncoder().encode(`<Properties xmlns="${source.rootNamespace}"/>`);
      if (insertions.some(item => item.source.rootNamespace !== source.rootNamespace || item.source.valueNamespace !== source.valueNamespace)) throw new PptxImportError("save-candidate-invalid", "Custom insertion source namespaces disagree.", { part });
      descriptionValidationBytes += current.length; descriptionValidationWork += current.length;
      if (descriptionValidationBytes >= limits.maxSaveValidationBytes || descriptionValidationWork >= limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Custom insertion inspection exceeds the shared validation budget.");
      const escape = async (value: string, attribute: boolean): Promise<string | undefined> => {
        if (/_x[0-9a-fA-F]{4}_x[0-9a-fA-F]{4}_/.test(value)) return undefined;
        let count = 0;
        for (const character of value) {
          if (++count % 16384 === 0) await checkpoint();
          const code = character.codePointAt(0)!; if (code !== 9 && code !== 10 && code !== 13 && (code < 32 || code >= 0xd800 && code <= 0xdfff || code === 0xfffe || code === 0xffff)) return undefined;
        }
        const protectedText = value.replace(/_x[0-9a-fA-F]{4}_/g, match => `_x005F_${match.slice(1)}`);
        return protectedText.replace(/[&<>"'\r\n\t]/g, ch => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;", "\r": "&#13;", "\n": attribute ? "&#10;" : "\n", "\t": attribute ? "&#9;" : "\t" })[ch]!);
      };
      let child = "", budget = current.length;
      for (const item of insertions) {
        await checkpoint(); const property = item.property, value = property.type === "boolean" ? (property.value ? "true" : "false") : property.type === "real" && Object.is(property.value, -0) ? "-0" : String(property.value);
        budget += (property.name.length + value.length) * 6 + 512;
        if (budget > limits.maxPreparedOutputBytes || budget > limits.maxSaveValidationWork - descriptionValidationWork) throw new PptxImportError("save-limit-exceeded", "Custom insertion exceeds the bounded serialization budget.");
        const name = await escape(property.name, true), text = await escape(value, false);
        if (name === undefined || text === undefined) return { diagnostics: [{ code: "save-edit-unsupported", severity: "error", source: { part }, message: "Custom insertion contains unsupported XML characters or overlapping escape-like sequences." }] };
        const type = { string: "lpwstr", boolean: "bool", integer: "i4", real: "r8", date: "filetime" }[property.type];
        child += `<cp:property xmlns:cp="${source.rootNamespace}" xmlns:vt="${source.valueNamespace}" fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${item.id}" name="${name}"><vt:${type}>${text}</vt:${type}></cp:property>`;
      }
      const parsed = await readXml(current, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part }, true);
      if (!parsed.root || parsed.root.namespaceUri !== source.rootNamespace || parsed.root.local !== "Properties") throw new PptxImportError("save-candidate-invalid", "Custom insertion source root disagrees with its binding.", { part });
      const changed = await appendXmlRootChild(current, parsed.root, parsed.inspection.encoding, child, limits.maxPreparedOutputBytes, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part });
      descriptionValidationWork += budget;
      if (descriptionValidationWork >= limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Custom insertion exceeds the shared validation budget.");
      replacementBytes += changed.bytes.length - (existing?.bytes.length ?? 0);
      if (replacementBytes > limits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Prepared XML payloads exceed maxPreparedOutputBytes.");
      if (source.create) {
        const plan = creationPlan!;
        if (plan.part !== part || plan.relationshipId !== source.create.relationshipId) throw new PptxImportError("save-candidate-invalid", "Custom property creation identity changed.", { part });
        additions.push({ name: part.slice(1), bytes: changed.bytes });
        finalDeclarations = finalDeclarations ? await appendPropertyDeclarations(finalDeclarations.relationships, finalDeclarations.contentTypes, { kind: "custom", part, relationshipId: plan.relationshipId, dialect: source.create.dialect }, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), maxOutputBytes: limits.maxPreparedOutputBytes }) : plan;
        for (const [name, bytes] of [["_rels/.rels", finalDeclarations.relationships], ["[Content_Types].xml", finalDeclarations.contentTypes]] as const) {
          const existingDeclaration = replacementByName.get(name);
          if (existingDeclaration) { replacementBytes += bytes.length - existingDeclaration.bytes.length; existingDeclaration.bytes = bytes; }
          else { replacementBytes += bytes.length; const replacement = { name, bytes }; replacements.push(replacement); replacementByName.set(name, replacement); }
        }
        if (replacementBytes > limits.maxPreparedOutputBytes) throw new PptxImportError("save-limit-exceeded", "Custom creation declarations exceed the output budget.");
      } else if (existing) existing.bytes = changed.bytes; else { const replacement = { name: inventory!.entryName, bytes: changed.bytes }; replacements.push(replacement); replacementByName.set(replacement.name, replacement); }
    }
    const chart = await prepareChartValues(original, analysis.chartValues, options, limits, checkpoint, replacementBytes, (analysis.chartLabels ?? []).map(({ binding, value }) => ({ closure: binding.closure, value })), (analysis.chartOrders ?? []).map(({ binding, requested }) => ({ part: binding.part, sources: binding.sources, complete: binding.complete, requested, ...(binding.researchColumnLegend ? { researchColumnLegend: true } : {}) })));
    replacements.push(...chart.replacements);
    if (new Set(replacements.map(item => item.name)).size !== replacements.length) return { diagnostics: [{ code: "save-text-candidate-unsupported", severity: "error", message: "Multiple edit domains cannot replace the same source part." }] };
    const approvedParts = new Set([...analysis.parts.keys(), ...metadataInsertionsByPart.keys(), ...metadataRemovalsByPart.keys(), ...(analysis.firstSlideNumber ? [analysis.firstSlideNumber.part] : []), ...descriptionsByPart.keys(), ...visibilityByPart.keys(), ...customNamesByPart.keys(), ...customRemovalsByPart.keys(), ...customInsertionsByPart.keys(), ...chart.changedParts]);
    const output = await replaceArchiveEntries(source, replacements, { ...(options.zipLimits ? { limits: options.zipLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), maxOutputBytes: limits.maxPreparedOutputBytes, additions });
    await checkpoint();
    const validationOptions = { ...options, zipLimits: { ...options.zipLimits, maxInputBytes: Math.max(options.zipLimits?.maxInputBytes ?? source.length, output.length) } };
    if (chart.validationBytes + descriptionValidationBytes >= limits.maxSaveValidationBytes || chart.validationWork + descriptionValidationWork >= limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Candidate validation exhausted the shared comparison budget.");
    const compared = await comparePackages(source, output, { ...validationOptions, comparisonLimits: {
      maxTotalExpandedBytes: limits.maxSaveValidationBytes - chart.validationBytes - descriptionValidationBytes, maxWorkUnits: limits.maxSaveValidationWork - chart.validationWork - descriptionValidationWork, maxDifferences: Math.max(limits.maxSaveChanges, approvedParts.size + 5),
    } });
    await checkpoint();
    const expected = new Set([...approvedParts].filter(part => part !== creating?.part && part !== coreCreating?.part && part !== extendedCreating?.part).map(part => `payload-changed:${part}:`));
    if (creating) { expected.add(`part-added:${creating.part}:`); expected.add("payload-changed:/_rels/.rels:"); expected.add("payload-changed:/[Content_Types].xml:"); expected.add(`relationship-added:/_rels/.rels:${creating.create!.relationshipId}`); }
    if (coreCreating) { expected.add(`part-added:${coreCreating.part}:`); expected.add("payload-changed:/_rels/.rels:"); expected.add("payload-changed:/[Content_Types].xml:"); expected.add(`relationship-added:/_rels/.rels:${coreCreating.create!.relationshipId}`); }
    if (extendedCreating) { expected.add(`part-added:${extendedCreating.part}:`); expected.add("payload-changed:/_rels/.rels:"); expected.add("payload-changed:/[Content_Types].xml:"); expected.add(`relationship-added:/_rels/.rels:${extendedCreating.create!.relationshipId}`); }
    if (compared.truncated || compared.differenceCount !== expected.size || compared.differences.some(difference => !expected.delete(`${difference.code}:${difference.part}:${difference.relationshipId ?? ""}`)) || expected.size) {
      throw new PptxImportError("save-candidate-invalid", "Candidate changed source parts, content types or relationships outside the approved scalar regions.");
    }
    // Reproject the actual candidate without creating another session or an
    // import->session cycle. The full mapped Core state must equal the plan.
    const context = await structureContext(output, validationOptions, true), inspection = await readStructureFromContext(context);
    if (creating) {
      const relation = context.findRelationship("/", creating.create!.relationshipId), actual = context.payload(creating.part), expectedBytes = additions.find(item => item.name === creating.part.slice(1))?.bytes;
      const relPayload = context.payload("/_rels/.rels").bytes;
      const expectedType = creating.create!.dialect === "strict" ? "http://purl.oclc.org/ooxml/officeDocument/relationships/customProperties" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
      if (context.entry.dialect !== creating.create!.dialect || !relation || relation.type !== expectedType || relation.targetMode !== "Internal" || relation.targetPart !== creating.part || relation.fragment !== undefined || inspection.metadata.custom?.part !== creating.part || actual.contentType !== "application/vnd.openxmlformats-officedocument.custom-properties+xml" || !expectedBytes || !creationPlan) throw new PptxImportError("save-candidate-invalid", "Created custom property relationship or part disagrees with its plan.");
      for (const [a, b] of [[actual.bytes, expectedBytes], [relPayload, finalDeclarations!.relationships]] as const) {
        descriptionValidationBytes += a.length + b.length; descriptionValidationWork += a.length;
        if (descriptionValidationBytes + chart.validationBytes + compared.before.expandedBytes + compared.after.expandedBytes > limits.maxSaveValidationBytes || descriptionValidationWork + chart.validationWork + compared.workUnits > limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Custom creation verification exceeds the shared validation budget.");
        if (a.length !== b.length) throw new PptxImportError("save-candidate-invalid", "Created custom property payload length differs from its plan.");
        for (let i = 0; i < a.length; i++) { if (a[i] !== b[i]) throw new PptxImportError("save-candidate-invalid", "Created custom property payload differs from its plan."); if (i % 65536 === 65535) await checkpoint(); }
      }
    }
    if (coreCreating) {
      const relation = context.findRelationship("/", coreCreating.create!.relationshipId), actual = context.payload(coreCreating.part), expectedBytes = additions.find(item => item.name === coreCreating.part.slice(1))?.bytes;
      if (context.entry.dialect !== coreCreating.create!.dialect || !relation || relation.type !== "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" || relation.targetMode !== "Internal" || relation.targetPart !== coreCreating.part || relation.fragment !== undefined || inspection.metadata.core?.part !== coreCreating.part || actual.contentType !== "application/vnd.openxmlformats-package.core-properties+xml" || !expectedBytes || !coreCreationPlan || !finalDeclarations) throw new PptxImportError("save-candidate-invalid", "Created core property relationship or part disagrees with its plan.");
      for (const [a, b] of [[actual.bytes, expectedBytes], [context.payload("/_rels/.rels").bytes, finalDeclarations.relationships]] as const) {
        descriptionValidationBytes += a.length + b.length; descriptionValidationWork += a.length;
        if (descriptionValidationBytes + chart.validationBytes + compared.before.expandedBytes + compared.after.expandedBytes > limits.maxSaveValidationBytes || descriptionValidationWork + chart.validationWork + compared.workUnits > limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Core creation verification exceeds the shared validation budget.");
        if (a.length !== b.length) throw new PptxImportError("save-candidate-invalid", "Created core property payload length differs from its plan.");
        for (let i = 0; i < a.length; i++) { if (a[i] !== b[i]) throw new PptxImportError("save-candidate-invalid", "Created core property payload differs from its plan."); if (i % 65536 === 65535) await checkpoint(); }
      }
    }
    if (extendedCreating) {
      const relation = context.findRelationship("/", extendedCreating.create!.relationshipId), actual = context.payload(extendedCreating.part), expectedBytes = additions.find(item => item.name === extendedCreating.part.slice(1))?.bytes;
      const expectedType = extendedCreating.create!.dialect === "strict" ? "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties" : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
      if (context.entry.dialect !== extendedCreating.create!.dialect || !relation || relation.type !== expectedType || relation.targetMode !== "Internal" || relation.targetPart !== extendedCreating.part || relation.fragment !== undefined || inspection.metadata.extended?.part !== extendedCreating.part || actual.contentType !== "application/vnd.openxmlformats-officedocument.extended-properties+xml" || !expectedBytes || !extendedCreationPlan || !finalDeclarations) throw new PptxImportError("save-candidate-invalid", "Created extended property relationship or part disagrees with its plan.");
      for (const [a, b] of [[actual.bytes, expectedBytes], [context.payload("/_rels/.rels").bytes, finalDeclarations.relationships]] as const) {
        descriptionValidationBytes += a.length + b.length; descriptionValidationWork += a.length;
        if (descriptionValidationBytes + chart.validationBytes + compared.before.expandedBytes + compared.after.expandedBytes > limits.maxSaveValidationBytes || descriptionValidationWork + chart.validationWork + compared.workUnits > limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Extended creation verification exceeds the shared validation budget.");
        if (a.length !== b.length) throw new PptxImportError("save-candidate-invalid", "Created extended property payload length differs from its plan.");
        for (let i = 0; i < a.length; i++) { if (a[i] !== b[i]) throw new PptxImportError("save-candidate-invalid", "Created extended property payload differs from its plan."); if (i % 65536 === 65535) await checkpoint(); }
      }
    }
    const currency = currencySave ? await projectCurrencies(context, checkpoint) : undefined;
    const projection = await projectStructure(inspection, limits.maxReportFindings, checkpoint, undefined, false, currency?.values);
    const rebound: ChartValueBinding[] | undefined = analysis.chartValues.length ? [] : undefined;
    const reboundLabels: ChartLabelBinding[] | undefined = analysis.chartLabels?.length ? [] : undefined;
    const reboundOrders: ChartSeriesOrderBinding[] | undefined = analysis.chartOrders?.length ? [] : undefined;
    await projectContent(projection.document, projection.report, context, limits, checkpoint, validationOptions, undefined, undefined, diagramPreview, rebound, textFontProfile, notesTextProfile, undefined, false, false, false, false, false, false, false, false, false, false, false, undefined, fieldTextProfile, undefined, false, false, false, false, false, analysis.chartValues.some(proposal => proposal.binding.sparseProfile), analysis.chartValues.some(proposal => proposal.binding.v3Profile), reboundLabels, analysis.chartLabels?.some(proposal => proposal.binding.seriesIndex !== undefined), reboundOrders, analysis.chartOrders?.some(proposal => proposal.binding.researchColumnLegend) ?? false);
    const orderByOwner = new Map<string, { oldIds: readonly number[]; positions: ReadonlyMap<number, number> }>();
    if (reboundOrders) {
      const key = (b: ChartSeriesOrderBinding) => JSON.stringify([b.slideId, b.elementId]);
      const actual = new Map<string, ChartSeriesOrderBinding>();
      for (const binding of reboundOrders) { await checkpoint(); const id = key(binding); if (actual.has(id)) throw new PptxImportError("save-candidate-invalid", "Candidate chart order owner is ambiguous."); actual.set(id, binding); }
      for (const proposal of analysis.chartOrders ?? []) {
        await checkpoint(); const binding = proposal.binding, id = key(binding), after = actual.get(id);
        if (orderByOwner.has(id) || !after || after.part !== binding.part || after.slidePart !== binding.slidePart) throw new PptxImportError("save-candidate-invalid", "Candidate chart order owner or source part differs from its plan.");
        const count = binding.sources.length, sort = count * Math.ceil(Math.log2(count + 1)), verificationWork = 2 * (2 * count + sort) + sort + count, work = verificationWork + sort + 2 * count;
        if (!Number.isSafeInteger(work) || work > limits.maxSaveValidationWork - chart.validationWork - compared.workUnits - descriptionValidationWork) throw new PptxImportError("save-limit-exceeded", "Chart order rebinding exceeds the shared validation budget.");
        try { await verifyChartSeriesOrderRebound(binding.sources, proposal.requested, after.sources, binding.complete, after.complete, verificationWork, checkpoint, chart.movedElementIndexes.get(binding.part)); }
        catch (cause) { if (cause instanceof TypeError) throw new PptxImportError("save-candidate-invalid", cause.message, { part: binding.part }, { cause }); if (cause instanceof RangeError) throw new PptxImportError("save-limit-exceeded", cause.message, { part: binding.part }, { cause }); throw cause; }
        descriptionValidationWork += work;
        if (binding.researchColumnLegend) {
          const before = chart.legendSourceProofs.get(binding.part), remaining = limits.maxSaveValidationWork - chart.validationWork - compared.workUnits - descriptionValidationWork;
          if (!before || !after.researchColumnLegend) throw new PptxImportError("save-candidate-invalid", "Candidate legend source proof is missing.", { part: binding.part });
          if (remaining <= 0) throw new PptxImportError("save-limit-exceeded", "Legend rebinding exhausts the shared validation budget.", { part: binding.part });
          let legend: Awaited<ReturnType<typeof captureColumnLegendSources>>;
          try { legend = await captureColumnLegendSources(context.payload(binding.part).bytes, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: binding.part }, remaining, checkpoint); }
          catch (cause) { if (cause instanceof RangeError) throw new PptxImportError("save-limit-exceeded", cause.message, { part: binding.part }, { cause }); if (cause instanceof TypeError) throw new PptxImportError("save-candidate-invalid", cause.message, { part: binding.part }, { cause }); throw cause; }
          descriptionValidationWork += legend.workUnits + before.entries.length + proposal.requested.length;
          if (descriptionValidationWork + chart.validationWork + compared.workUnits > limits.maxSaveValidationWork) throw new PptxImportError("save-limit-exceeded", "Legend rebinding exceeds the shared validation budget.");
          const positions = new Map(proposal.requested.map((identity, index) => [identity, index]));
          if (JSON.stringify(legend.physicalIds) !== JSON.stringify(proposal.requested) || legend.entries.length !== before.entries.length || legend.entries.some((entry, i) => {
            const original = before.entries[i], identity = original && before.physicalIds[original.position], expected = identity === undefined ? undefined : positions.get(identity);
            return !original || expected === undefined || entry.position !== expected || entry.elementIndex !== original.elementIndex || entry.namespaceUri !== original.namespaceUri;
          })) throw new PptxImportError("save-candidate-invalid", "Reimported legend entries no longer target their original series.", { part: binding.part });
        }
        const oldIds = [...binding.sources].sort((a, b) => a.order - b.order).map(source => source.identity), positions = new Map(proposal.requested.map((identity, index) => [identity, index]));
        orderByOwner.set(id, { oldIds, positions });
      }
    }
    const remapSeries = (slideId: string, elementId: string, index: number): number => {
      const order = orderByOwner.get(JSON.stringify([slideId, elementId])); if (!order) return index;
      const identity = order.oldIds[index], position = identity === undefined ? undefined : order.positions.get(identity);
      if (position === undefined) throw new PptxImportError("save-candidate-invalid", "Edited series is absent from the order permutation."); return position;
    };
    const chartKey = (b: ChartValueBinding) => JSON.stringify([b.slideId, b.elementId, b.seriesIndex, b.pointIndex, b.coordinate]);
    const reboundByIdentity = new Map(rebound?.map(b => [chartKey(b), b]));
    if (rebound) for (const proposal of analysis.chartValues) {
      await checkpoint();
      const original = proposal.binding, binding = reboundByIdentity.get(chartKey({ ...original, seriesIndex: remapSeries(original.slideId, original.elementId, original.seriesIndex) }));
      if (!binding || binding.status !== "bound" || !Object.is(binding.baselineValue, proposal.value) || !Object.is(chartNumber(binding.workbookLexical ?? ""), proposal.value) || JSON.stringify(binding.workbook) !== JSON.stringify(original.workbook) || binding.sourceElementIndex !== (chart.movedElementIndexes.get(original.source.part)?.get(original.sourceElementIndex!) ?? original.sourceElementIndex)) throw new PptxImportError("save-candidate-invalid", "Candidate chart/workbook numeric bindings do not agree with the approved values/identities.", original.source);
    }
    if (reboundLabels) {
      const labelKey = (b: ChartLabelBinding) => JSON.stringify([b.slideId, b.elementId, b.seriesIndex === undefined ? "category" : "series", b.seriesIndex ?? b.categoryIndex]);
      const byIdentity = new Map<string, ChartLabelBinding>();
      for (const binding of reboundLabels) { await checkpoint(); const key = labelKey(binding); if (byIdentity.has(key)) throw new PptxImportError("save-candidate-invalid", "Candidate chart label identity is ambiguous."); byIdentity.set(key, binding); }
      const physical = (binding: ChartLabelBinding, ensureSpace = false, translate = false) => binding.closure.uses.map(({ cache, cell }) => JSON.stringify([cache.part, translate ? chart.movedElementIndexes.get(cache.part)?.get(cache.elementIndex) ?? cache.elementIndex : cache.elementIndex, cache.namespaceUri, cell.workbookPart, cell.worksheetPart, cell.reference, cell.elementIndex, cell.namespaceUri, ensureSpace ? "preserve" : cell.space ?? null])).sort();
      for (const proposal of analysis.chartLabels ?? []) {
        await checkpoint();
        const original = proposal.binding;
        const actual = byIdentity.get(labelKey(original.seriesIndex === undefined ? original : { ...original, seriesIndex: remapSeries(original.slideId, original.elementId, original.seriesIndex) }));
        if (!actual || actual.closure.status !== "bound" || actual.baselineValue !== proposal.value || JSON.stringify(physical(actual)) !== JSON.stringify(physical(proposal.binding, /^[\t\n\r ]|[\t\n\r ]$/u.test(proposal.value), true)) ||
            actual.closure.uses.some(use => use.projected !== proposal.value || chartString(use.cache.lexical) !== proposal.value || chartString(use.cell.lexical) !== proposal.value)) throw new PptxImportError("save-candidate-invalid", "Candidate chart label cache/workbook bindings do not agree with the approved value and source identities.");
      }
    }
    if (!equalState(snapshot(projection.document, limits.maxSnapshotValues), candidate, limits.maxSnapshotValues)) {
      throw new PptxImportError("save-candidate-invalid", "Reimported candidate differs from the prepared Core snapshot.");
    }
    await checkpoint(); return { bytes: output, validationBytes: chart.validationBytes + descriptionValidationBytes + compared.before.expandedBytes + compared.after.expandedBytes, validationWork: chart.validationWork + descriptionValidationWork + compared.workUnits };
  } catch (cause) {
    if (cause instanceof PptxImportError && cause.code.startsWith("import-")) {
      throw new PptxImportError(cause.code === "import-limit-exceeded" ? "save-limit-exceeded" : "save-candidate-invalid", cause.message, cause.source, { cause });
    }
    if (cause instanceof PackageInspectionError || cause instanceof XmlSourceEditError) {
      const code = cause.code === "operation-aborted" ? "operation-aborted" : cause.code.endsWith("limit-exceeded") ? "save-limit-exceeded" : "save-candidate-invalid";
      throw new PptxImportError(code, cause.message, undefined, { cause });
    }
    throw cause;
  }
}
