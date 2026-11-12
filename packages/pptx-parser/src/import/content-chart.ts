import type { Box, BubbleChartSeriesInput, ChartElementInput, ComboChartSeriesInput, ChartSeriesBaseInput, ChartStyleInput, ChartTextStyleInput, LineChartSeriesInput, PieChartSeriesInput, ScatterChartSeriesInput } from "@pptkit/core";
import type { structureContext } from "../presentation/structure-context.js";
import { attribute, children, invalidStructure } from "../presentation/structure-xml.js";
import { PptxImportError, type PptxSessionLimits, type PptxSourceLocation } from "../types/import.js";
import type { OpcInspectionOptions, OpcRelationshipInfo } from "../types/opc.js";
import type { ChartValueBinding } from "./chart-value-bindings.js";
import type { ChartSeriesOrderSource } from "./chart-series-order.js";
import { chartLabelClosure, type ChartLabelClosure, type ChartLabelUse } from "./chart-label-closure.js";
import type { XmlElement } from "../xml/read-xml.js";
import { contentPaths, contentReader, type ContentReader } from "./content-reader.js";
import { chartNumber, chartNumericSourcesAgree, chartReference, type ChartReference } from "./chart-data.js";
import { chartAxes, comboAxes, chartBoolean, chartLegend, chartMarker, chartScalar, chartTextStyle, chartTextProperties, scatterAxes } from "./chart-style.js";
import { chartWorkbooks, workbookRange, type ChartWorkbook } from "./chart-workbook.js";
import { captureColumnLegendSources } from "./chart-legend-source.js";

type Context = Awaited<ReturnType<typeof structureContext>>;
type Report = (feature: string, source: PptxSourceLocation, reason: string) => void;
type Result = { fields: ChartElementInput; previewFallback: boolean; valueBindings?: readonly Omit<ChartValueBinding, "slideId" | "elementId">[]; labelCandidates?: readonly ChartLabelClosure[]; seriesNameCandidates?: readonly ChartLabelClosure[]; seriesOrderCandidate?: { readonly part: string; readonly sources: readonly ChartSeriesOrderSource[]; readonly complete: boolean; readonly researchColumnLegend?: true; readonly localDependencyReason?: string } } | { reason: string };

/** One importer owns workbook reuse and aggregate budgets across all slides. */
export function contentCharts(context: Context, limits: Required<PptxSessionLimits>, options: OpcInspectionOptions, relationships: readonly OpcRelationshipInfo[], checkpoint: () => Promise<void>, captureValues = false, sparseCoordinateWrites = false, v3Writes = false, captureLabelCandidates = false, captureSeriesNameCandidates = false, captureSeriesOrder = false, researchColumnLegend = false) {
  if ((captureLabelCandidates || captureSeriesNameCandidates || captureSeriesOrder) && !captureValues) throw new TypeError("Chart source candidate inspection requires numeric source capture.");
  let points = 0;
  const countPoint = () => { if (++points > limits.maxChartDataPoints) throw new PptxImportError("import-limit-exceeded", "Chart data exceeds maxChartDataPoints."); };
  const workbook = chartWorkbooks(options, limits, countPoint, checkpoint, captureValues);
  const bySource = new Map<string, OpcRelationshipInfo[]>(), incoming = new Map<string, number>(), chartUses = new Map<string, number>(), formulas = new Map<string, number>();
  let scanned = false, bindingWork = 0, orderPackageDependency: string | undefined;
  const bindingStep = async () => { if (++bindingWork > limits.maxChartBindingWork) throw new PptxImportError("import-limit-exceeded", "Chart numeric binding exceeds maxChartBindingWork."); await checkpoint(); };
  for (const rel of relationships) { if (captureValues && rel.targetPart) incoming.set(rel.targetPart, (incoming.get(rel.targetPart) ?? 0) + 1); const list = bySource.get(rel.source) ?? []; list.push(rel); bySource.set(rel.source, list); }
  return async (data: XmlElement, box: Box, slideReader: ContentReader, sourceObjectId: string, report: Report): Promise<Result> => {
    const c = context.a.replace(/\/main$/, "/chart");
    if (captureValues && !scanned) {
      scanned = true;
      for (const [sourcePart, tree] of context.rawXml) {
        const pending = [tree];
        while (pending.length) {
          await bindingStep(); const n = pending.pop()!;
          if (captureSeriesOrder && (n.local === "extLst" || n.namespaceUri === context.p && n.local === "timing")) orderPackageDependency ??= "Package-wide timing or extensions require a dependency-aware order policy.";
          if (n.namespaceUri === c && n.local === "f") formulas.set(sourcePart, (formulas.get(sourcePart) ?? 0) + 1);
          if (n.namespaceUri === c && n.local === "chart") {
            const id = attribute(n, "id", context.r), rel = id ? context.findRelationship(sourcePart, id) : undefined;
            if (rel?.targetPart) chartUses.set(rel.targetPart, (chartUses.get(rel.targetPart) ?? 0) + 1);
          }
          pending.push(...children(n));
        }
      }
    }
    const reference = slideReader.one(data, c, "chart", true)!;
    const id = slideReader.attr(reference, "id", context.r);
    const relation = context.relationship(slideReader.part, id, "chart");
    const part = context.target(relation, "chart");
    const root = await context.root(part, "chartSpace", "drawingml.chart", c);
    let localDependencyReason: string | undefined, hasLegendEntry = false, legendProven = false;
    if (captureSeriesOrder) {
      const related = bySource.get(part) ?? [];
      if (related.some(rel => rel.targetMode !== "Internal" || !["http://schemas.openxmlformats.org/officeDocument/2006/relationships/package", "http://purl.oclc.org/ooxml/officeDocument/relationships/package"].includes(rel.type))) localDependencyReason = "Chart auxiliary or external relationships need a dependency-aware order policy.";
      const slideRoot = context.rawXml.get(slideReader.part), chartRoot = context.rawXml.get(part);
      if (!slideRoot || !chartRoot) localDependencyReason ??= "Raw chart or slide source is unavailable.";
      for (const [tree, scope] of [[slideRoot, "slide"], [chartRoot, "chart"]] as const) if (tree) {
        const pending = [{ node: tree, insideSeries: false }];
        while (pending.length) {
          await bindingStep(); const { node, insideSeries } = pending.pop()!;
          if (scope === "slide" && node.namespaceUri === context.p && (node.local === "timing" || node.local === "extLst")) localDependencyReason ??= "Slide timing or extensions require a dependency-aware order policy.";
          if (scope === "chart" && (node.namespaceUri === c && node.local === "extLst" || ![c, context.a, context.r].includes(node.namespaceUri))) localDependencyReason ??= "Chart extensions or foreign elements require a dependency-aware order policy.";
          if (scope === "chart" && node.namespaceUri === c && node.local === "legendEntry") { hasLegendEntry = true; if (!researchColumnLegend) localDependencyReason ??= "Indexed chart-level legend, label or formatting references require a dependency-aware order policy."; }
          if (scope === "chart" && node.namespaceUri === c && (node.local === "pivotFmt" || node.local === "bandFmt" || node.local === "dLbl" && !insideSeries)) localDependencyReason ??= "Indexed chart-level legend, label or formatting references require a dependency-aware order policy.";
          const childInsideSeries = insideSeries || node.namespaceUri === c && node.local === "ser";
          for (const child of children(node)) pending.push({ node: child, insideSeries: childInsideSeries });
        }
      }
      if (hasLegendEntry && researchColumnLegend) {
        try {
          const remaining = limits.maxChartBindingWork - bindingWork;
          const proof = await captureColumnLegendSources(context.payload(part).bytes, { ...(options.xmlLimits ? { limits: options.xmlLimits } : {}), ...(options.signal ? { signal: options.signal } : {}), partName: part }, remaining, checkpoint);
          bindingWork += proof.workUnits; legendProven = true;
        } catch (cause) {
          if (cause instanceof TypeError) localDependencyReason ??= "Indexed chart legend source is outside the experimental column-chart profile.";
          else if (cause instanceof RangeError) throw new PptxImportError("import-limit-exceeded", cause.message, { part }, { cause });
          else throw cause;
        }
      }
      localDependencyReason ??= orderPackageDependency;
    }
    const r = contentReader(part, context.p, context.a), paths = await contentPaths(root, checkpoint);
    const at = (node: XmlElement): PptxSourceLocation => ({ part, path: paths.get(node)!, sourceObjectId });
    const note = (feature: string, node: XmlElement, reason: string) => report(feature, at(node), reason);
    r.known(root);
    for (const rel of bySource.get(part) ?? []) {
      report("chart.relationship", { part, relationshipId: rel.id, sourceObjectId }, `${rel.type}: ${rel.targetMode} ${rel.targetPart ?? rel.target}. External targets are never fetched; auxiliary styles/user shapes/links remain source-owned.`);
    }
    const cellUses = new Map<string, number>(), labelUses: ChartLabelUse[] = [], labelKeys: string[] = [], nameUses: ChartLabelUse[] = [], nameKeys: string[] = [];
    const labelCacheSafety = new WeakMap<XmlElement, boolean>(), labelReferenceSafety = new WeakMap<XmlElement, boolean>();
    const seriesOrderSources: ChartSeriesOrderSource[] = []; let seriesOrderUnsafe = false;
    let profileUnsafe: string | undefined, labelUnsafe: string | undefined; let interpretedFormulas = 0;
    const rawLabelPoint = async (target: XmlElement, point: { node: XmlElement; index: number; lexical: string }, formula: string) => {
      await bindingStep();
      const value = await context.sourceElement(part, point.node);
      if (value.node.namespaceUri !== c || value.node.local !== "v" || value.node.attributes.length || !value.node.children.every(n => n.kind === "text" || n.kind === "cdata") ||
        value.node.children.map(n => n.kind === "text" || n.kind === "cdata" ? n.value : "").join("") !== point.lexical || !value.parent) return undefined;
      const p = await context.sourceElement(part, value.parent);
      if (p.node.namespaceUri !== c || p.node.local !== "pt" || p.node.attributes.length !== 1 || p.node.attributes[0]?.namespaceUri !== "" || p.node.attributes[0]?.local !== "idx" ||
        !/^\d+$/u.test(p.node.attributes[0].value) || Number(p.node.attributes[0].value) !== point.index ||
        !p.node.children.every(n => n === value.node || n.kind === "text" && /^[\t\r\n ]*$/u.test(n.value)) || !p.parent) return undefined;
      const cache = await context.sourceElement(part, p.parent);
      if (cache.node.namespaceUri !== c || cache.node.local !== "strCache" || cache.node.attributes.length || !cache.parent) return undefined;
      let safeCache = labelCacheSafety.get(cache.node);
      if (safeCache === undefined) {
        safeCache = true;
        for (const child of cache.node.children) {
          await bindingStep();
          if (child.kind === "text" && /^[\t\r\n ]*$/u.test(child.value)) continue;
          if (child.kind !== "element" || child.namespaceUri !== c || !["pt", "ptCount"].includes(child.local)) safeCache = false;
          else if (child.local === "ptCount" && (child.attributes.length !== 1 || child.attributes[0]?.namespaceUri !== "" || child.attributes[0]?.local !== "val" || child.children.some(n => n.kind !== "text" || !/^[\t\r\n ]*$/u.test(n.value)))) safeCache = false;
        }
        labelCacheSafety.set(cache.node, safeCache);
      }
      if (!safeCache) return undefined;
      const ref = await context.sourceElement(part, cache.parent), owner = await context.sourceElement(part, target);
      if (ref.node.namespaceUri !== c || ref.node.local !== "strRef" || ref.node.attributes.length || ref.parent !== owner.node || owner.node.namespaceUri !== c || !["cat", "tx"].includes(owner.node.local) || owner.node.attributes.length) return undefined;
      let safeReference = labelReferenceSafety.get(ref.node);
      if (safeReference === undefined) {
        const formulas = children(ref.node).filter(n => n.namespaceUri === c && n.local === "f"), caches = children(ref.node).filter(n => n.namespaceUri === c && n.local === "strCache");
        const rawFormula = formulas[0];
        safeReference = formulas.length === 1 && caches.length === 1 && caches[0] === cache.node && !!rawFormula && rawFormula.attributes.length === 0 &&
          rawFormula.children.every(n => n.kind === "text" || n.kind === "cdata") && rawFormula.children.map(n => n.kind === "text" || n.kind === "cdata" ? n.value : "").join("") === formula;
        for (const node of ref.node.children) {
          await bindingStep();
          if (!(node.kind === "text" && /^[\t\r\n ]*$/u.test(node.value) || node.kind === "element" && node.namespaceUri === c && ["f", "strCache"].includes(node.local))) safeReference = false;
        }
        for (const node of owner.node.children) {
          await bindingStep();
          if (!(node === ref.node || node.kind === "text" && /^[\t\r\n ]*$/u.test(node.value))) safeReference = false;
        }
        labelReferenceSafety.set(ref.node, safeReference);
      }
      if (!safeReference) return undefined;
      return value;
    };
    const parse = async (): Promise<Result> => {
      let book: ChartWorkbook | undefined, bookPart: string | undefined;
      const valueBindings: Omit<ChartValueBinding, "slideId" | "elementId">[] = [];
      let unsafe = captureValues && (bySource.get(part) ?? []).some(r => r.targetMode === "External") ? "Chart external dependencies are not writable by this profile." : undefined;
      if (captureValues && (incoming.get(part) !== 1 || chartUses.get(part) !== 1)) unsafe = "Shared chart parts require an explicit shared-edit policy.";
      if (captureValues && context.rawXml.get(part) !== root) unsafe = "Chart markup compatibility projection is not admitted for numeric writes.";
      const resolved = new WeakMap<XmlElement, { ref: ChartReference; stored?: Awaited<ReturnType<typeof workbookRange>> }>();
      const countedFormulas = new WeakSet<XmlElement>();
      const external = r.one(root, c, "externalData");
      if (external) {
        const rel = context.relationship(part, r.attr(external, "id", context.r), "package");
        if (rel.targetMode === "External") note("chart.data-source", external, "External workbook is not fetched. Only complete chart caches can be mapped.");
        else {
          const target = context.target(rel, "package"), payload = context.payload(target);
          if (payload.contentType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") { bookPart = target; book = await workbook(target, payload.bytes); }
          else note("chart.data-source", external, "Embedded data is not a supported XLSX workbook; only complete chart caches can be mapped.");
        }
        if (chartBoolean(external, "autoUpdate", c, r, false)) { if (captureValues) profileUnsafe = "Auto-updated chart data is not admitted for numeric writes."; note("chart.data-source", external, "Source requests auto-update. Import uses stored snapshots without refreshing links or recalculating formulas."); }
      }
      if (captureValues && bookPart && incoming.get(bookPart) !== 1) unsafe = "Shared embedded workbooks require an explicit shared-edit policy.";
      if (captureValues && book?.editBlockReason) unsafe = book.editBlockReason;
      labelUnsafe = unsafe;
      const chart = r.one(root, c, "chart", true)!; r.known(chart);
      const blankMode = chartScalar(chart, "dispBlanksAs", c, r);
      let supportedBlankMode: "gap" | "zero" | "span" | undefined, admitSparse = false;
      const resolve = async (node: XmlElement, numeric: boolean, blankLength?: number | "declared", scatterCounterpart = false, allowTrailingBubblePadding = false): Promise<ChartReference> => {
        const ref = await chartReference(node, numeric, c, r, checkpoint, countPoint, captureValues, numeric && supportedBlankMode !== undefined ? blankLength : undefined, admitSparse, allowTrailingBubblePadding);
        if (captureValues && ref.formula !== undefined && !countedFormulas.has(node)) { interpretedFormulas++; countedFormulas.add(node); }
        let stored: Awaited<ReturnType<typeof workbookRange>> | undefined;
        if (ref.formula !== undefined) {
          const cacheValues = ref.values ?? ref.blankValues;
          const expected = cacheValues?.length ?? (numeric && supportedBlankMode !== undefined && typeof blankLength === "number" ? blankLength : undefined);
          const allowWorkbookBlanks = numeric && supportedBlankMode !== undefined && blankLength !== undefined && (expected === blankLength || admitSparse && blankLength === "declared" && ref.blankValues !== undefined);
          stored = await workbookRange(book ?? { sheets: new Map() }, ref.formula, numeric, countPoint, checkpoint, expected, allowWorkbookBlanks, admitSparse);
          if (stored.reason) note("chart.data-source", node, stored.reason);
          if (stored.cachedFormula) note("chart.data-source", node, "Workbook formula cells use stored results. No formula is evaluated or recalculated.");
          const workbookValues = stored.values ?? stored.blankValues;
          if (cacheValues && workbookValues) {
            const equal = cacheValues.length === workbookValues.length && cacheValues.every((value, i) => {
              const other = workbookValues[i];
              return value === null || other === null || other === undefined ? value === other : numeric ? chartNumericSourcesAgree(value, other) : value === other;
            });
            const agreement = ref.blankValues || stored.blankValues ? "Chart cache agrees with stored workbook values including blanks." : "Complete chart cache agrees with stored workbook values.";
            const conflict = ref.blankValues && stored.values ? "Sparse chart cache conflicts with a complete nonblank workbook range; the chart cache remains authoritative for this snapshot." : "Chart cache conflicts with stored workbook values. The chart cache is authoritative for this snapshot; workbook differences are retained in the source.";
            note("chart.data-source", node, equal ? agreement : conflict);
          }
        }
        if (captureValues && bookPart) for (const item of stored?.cells ?? (admitSparse ? stored?.sparseCells?.values() : undefined) ?? []) {
          const key = JSON.stringify([bookPart, item.cell.worksheetPart, item.reference]); cellUses.set(key, (cellUses.get(key) ?? 0) + 1);
        }
        if (captureValues && ref.formula && stored?.reason) profileUnsafe = "Chart formulas outside resolved local workbook ranges are not writable.";
        resolved.set(node, { ref, ...(stored ? { stored } : {}) });
        if (ref.blankValues) {
          note("chart.data-source", node, `Explicit ${supportedBlankMode} cache maps missing points to null; ${admitSparse ? "blank positions remain source-owned; eligible existing numeric points may be bound separately" : "workbook blanks and chart data edits remain source-owned"}.`);
          if (captureValues && !admitSparse) profileUnsafe = "Sparse chart caches are not admitted for numeric preservation writes.";
        }
        if (ref.values || ref.blankValues || ref.reason) return ref; // Never repair an existing sparse cache from the workbook.
        if (stored?.values) { note("chart.data-source", node, "Missing chart cache was mapped from complete stored workbook values; source formulas are not retained in Core. The source has no cached snapshot and native first-open display may differ."); return { values: stored.values }; }
        if (stored?.blankValues) {
          note("chart.data-source", node, `Missing chart cache was mapped from a ${scatterCounterpart ? "counterpart-bounded" : "category-bounded"} workbook range with explicit blanks; source formulas are not retained in Core. The source has no cached snapshot and native first-open display may differ.`);
          if (captureValues) profileUnsafe = "Workbook blanks are not admitted for numeric preservation writes.";
          return { blankValues: stored.blankValues };
        }
        return { reason: "Chart cache is missing and no complete compatible workbook range is available." };
      };
      async function bindNumericSeries(seriesIndex: number, tx: XmlElement,
        targets: readonly (readonly ["x" | "y" | "size", XmlElement, readonly (number | null | undefined)[]])[], bubble: boolean) {
        const agree = (target: XmlElement, numericSource: boolean) => {
          const v = resolved.get(target);
          const source = v?.ref.values ?? (admitSparse ? v?.ref.blankValues : undefined), stored = v?.stored?.values ?? (admitSparse ? v?.stored?.blankValues : undefined);
          return source && stored && source.length === stored.length && source.every((text, i) => { const other = stored[i]; return text === null || other === null || other === undefined ? text === other : numericSource ? chartNumericSourcesAgree(text, other) : text === other; });
        };
        const reason = unsafe ?? (!bookPart || !book ? "An unshared internal XLSX workbook is required." : !agree(tx, false) || targets.some(([, target]) => !agree(target, true)) ? (bubble ? "Complete name/X/Y/size caches and explicit local workbook references must agree." : "Complete name/X/Y caches and explicit local workbook references must agree.") : undefined);
        for (const [coordinate, target, values] of targets) {
          const source = resolved.get(target);
          const blockedSource = reason ?? (!source?.ref.formula || !(source.ref.points || admitSparse && source.ref.sparsePoints) ? "Existing numeric reference/cache points are required." : undefined);
          for (const [pointIndex, baselineValue] of values.entries()) {
            await bindingStep();
            if (baselineValue === null || baselineValue === undefined) continue;
            const point = source?.ref.points?.[pointIndex] ?? (admitSparse ? source?.ref.sparsePoints?.get(pointIndex) : undefined), cell = source?.stored?.cells?.[pointIndex] ?? (admitSparse ? source?.stored?.sparseCells?.get(pointIndex) : undefined);
            const raw = point ? await context.sourceElement(part, point.node) : undefined;
            const scalar = raw && raw.node.children.every(n => n.kind === "text" || n.kind === "cdata");
            const blocked = blockedSource ?? (!scalar || !cell?.cell.source || !cell.cell.numeric || cell.cell.formula ? "Cache/workbook value lacks a direct numeric scalar source." : undefined);
            valueBindings.push({ seriesIndex, pointIndex, coordinate, baselineValue, ...(admitSparse ? { sparseProfile: true } : {}), ...(bubble ? { v3Profile: true } : {}), source: { ...at(target), ...(raw ? { elementIndex: raw.elementIndex } : {}) },
              ...(raw ? { sourceElementIndex: raw.elementIndex } : {}), ...(point ? { cacheLexical: point.lexical, cacheNamespace: c } : {}),
              ...(bookPart && cell?.cell.source ? { workbook: { part: bookPart, worksheetPart: cell.cell.source.part, cellReference: cell.reference, sourceElementIndex: cell.cell.source.elementIndex }, workbookLexical: cell.cell.source.lexical, workbookNamespace: cell.cell.source.namespaceUri } : {}),
              status: blocked ? "unmapped" : "bound", ...(blocked ? { reason: blocked } : {}) });
          }
        }
      }
      const plot = r.one(chart, c, "plotArea", true)!; r.known(plot);
      const families = children(plot).filter(n => n.namespaceUri === c && n.local.endsWith("Chart"));
      const combo = families.length === 2 && families.some(node => node.local === "barChart") && families.some(node => node.local === "lineChart");
      const family = combo ? families.find(node => node.local === "barChart") : families[0];
      const lineFamily = combo ? families.find(node => node.local === "lineChart") : undefined;
      if ((!combo && families.length !== 1) || !family || !["barChart", "lineChart", "areaChart", "pieChart", "doughnutChart", "scatterChart", "bubbleChart"].includes(family.local)) return { reason: "Only single-family 2D bar/line/area/pie/doughnut/scatter/bubble charts are mapped; unsupported combinations and other families remain source-only." };
      r.known(family);
      const chartType = combo ? "combo" : family.local === "barChart" ? "bar" : family.local === "lineChart" ? "line" : family.local === "areaChart" ? "area" : family.local === "doughnutChart" ? "doughnut" : family.local === "scatterChart" ? "scatter" : family.local === "bubbleChart" ? "bubble" : "pie";
      admitSparse = chartType === "scatter" && sparseCoordinateWrites;
      const areaGrouping = chartType === "area" ? chartScalar(family, "grouping", c, r) ?? "standard" : undefined;
      if (areaGrouping !== undefined && !["standard", "stacked", "percentStacked"].includes(areaGrouping)) return { reason: "Area grouping is outside Core's supported groups." };
      if (chartType === "area" && blankMode !== undefined && blankMode !== "zero" && blankMode !== "gap") return { reason: "Area blank display is unsupported; source is retained." };
      if (chartType === "area" && areaGrouping !== "standard" && blankMode !== undefined && blankMode !== "zero") return { reason: "Stacked area supports only explicit zero blank display; unsupported source is retained." };
      if (blankMode === "gap" || blankMode === "zero" || (blankMode === "span" && (chartType === "line" || chartType === "scatter"))) supportedBlankMode = blankMode;
      else if (blankMode !== undefined) r.note(chart, "Blank display policy is unsupported for this chart family; its source value is retained without projection.");
      if (chartType === "line" && ![undefined, "standard"].includes(chartScalar(family, "grouping", c, r))) return { reason: "Stacked line charts cannot map to Core's unstacked line model." };
      if (lineFamily) {
        r.known(lineFamily);
        if (chartScalar(family, "barDir", c, r) !== "col" || ![undefined, "clustered"].includes(chartScalar(family, "grouping", c, r)) ||
          ![undefined, "standard"].includes(chartScalar(lineFamily, "grouping", c, r)) ||
          ![undefined, "150", "150%"].includes(chartScalar(family, "gapWidth", c, r)) ||
          ![undefined, "0", "0%"].includes(chartScalar(family, "overlap", c, r)) ||
          ![undefined, "gap", "zero"].includes(blankMode) ||
          chartBoolean(lineFamily, "smooth", c, r, false) || chartBoolean(family, "varyColors", c, r, false) ||
          chartBoolean(lineFamily, "varyColors", c, r, false)) return { reason: "Combo profile requires default clustered columns and straight lines with gap or zero blanks." };
      }
      let doughnutOptions: { holeSize: number; firstSliceAngle: number } | undefined;
      if (chartType === "doughnut") {
        const hole = r.one(family, c, "holeSize"), angle = r.one(family, c, "firstSliceAng");
        if (!hole) return { reason: "Doughnut hole size must be explicit for projection; omitted size is not assumed." };
        // Preserve valid unsigned-byte values outside Core's authoring range as source-only.
        // PowerPoint can itself save holeSize=0 after opening an omitted hole size.
        const holeLexical = r.attr(hole, "val");
        if (holeLexical === undefined) return { reason: "Doughnut hole size has no value; source is retained without projection." };
        // OOXML also permits integer percentages such as 50%; retain their
        // original spelling in source-owned XML and project only the value.
        const percent = /^0*(?:[1-9]|[1-8][0-9]|90)%$/.test(holeLexical);
        const holeSize = percent ? Number(holeLexical.slice(0, -1)) : r.number(hole, "val", 0, 255);
        if (percent) r.note(hole, "Percentage-form hole size is preserved as source text; native application compatibility is not established.");
        if (holeSize < 1 || holeSize > 90) return { reason: "Doughnut hole size is outside Core's supported 1–90 range; source is retained without projection." };
        const angleLexical = angle ? r.attr(angle, "val") : undefined;
        const firstSliceAngle = angle && angleLexical !== undefined ? r.number(angle, "val", 0, 360) : 0;
        if (!angle) r.note(family, "Absent first slice angle projects as zero; the source remains unchanged for preservation saves.");
        else if (angleLexical === undefined) r.note(angle, "First slice angle without a value projects as zero; the source remains unchanged for preservation saves.");
        doughnutOptions = { holeSize, firstSliceAngle };
      }
      const sourceSeries = [family, ...(lineFamily ? [lineFamily] : [])].flatMap(owner =>
        children(owner).filter(n => n.namespaceUri === c && n.local === "ser").map(node => ({ node, kind: owner === lineFamily ? "line" as const : "bar" as const })));
      if (combo && (!sourceSeries.some(item => item.kind === "bar") || !sourceSeries.some(item => item.kind === "line"))) return { reason: "Combo chart requires both column and line series." };
      if (!sourceSeries.length || (chartType === "pie" && sourceSeries.length !== 1)) return { reason: "Chart requires series; Core pie supports exactly one series." };
      const seenIds = new Set<number>(), seenOrders = new Set<number>();
      const ordered = sourceSeries.map(({ node, kind }) => {
        r.known(node);
        const identity = r.number(r.one(node, c, "idx", true)!, "val", 0, 4294967295), order = r.number(r.one(node, c, "order", true)!, "val", 0, 4294967295);
        if (seenIds.has(identity) || seenOrders.has(order)) invalidStructure("Chart series identities/orders collide.", part);
        seenIds.add(identity); seenOrders.add(order); return { node, identity, order, kind };
      }).sort((a, b) => a.order - b.order);
      if (ordered.some((item, index) => item.node !== sourceSeries[index]!.node)) note("chart.series-order", family, "Series c:order differs from XML element order. Core follows c:order; PowerPoint for Mac 16.104 displayed XML element order in a controlled column-chart example.");
      if (combo) {
        let lineStarted = false;
        for (const item of ordered) {
          if (item.kind === "line") lineStarted = true;
          else if (lineStarted) return { reason: "Combo series order must place columns before lines." };
        }
      }
      const safeLabelOwner = async (target: XmlElement, node: XmlElement, owner: XmlElement) => {
        const chain = [target, node, owner, plot, chart, root]; let safe = true;
        for (let i = 0; i < chain.length - 1; i++) {
          await bindingStep();
          const child = await context.sourceElement(part, chain[i]!);
          const parent = await context.sourceElement(part, chain[i + 1]!);
          if (child.parent !== parent.node || child.node !== chain[i] || parent.node !== chain[i + 1]) safe = false;
        }
        return safe;
      };
      const collectSeriesName = async (node: XmlElement, tx: XmlElement, owner: XmlElement, seriesIndex: number, projected: string) => {
        if (!captureSeriesNameCandidates) return;
        const key = `series:${seriesIndex}`; nameKeys.push(key); await bindingStep();
        const safe = await safeLabelOwner(tx, node, owner), source = resolved.get(tx);
        const point = source?.ref.formula ? source.ref.points?.[0] : undefined, cell = source?.stored?.cells?.[0], string = cell?.cell.stringSource;
        if (!safe || !point || !cell || !string || !bookPart || cell.cell.formula) return;
        const raw = await rawLabelPoint(tx, point, source!.ref.formula!);
        if (raw) nameUses.push({ logicalKey: key, projected, cache: { part, elementIndex: raw.elementIndex, lexical: point.lexical, namespaceUri: c }, cell: { workbookPart: bookPart, worksheetPart: string.part, reference: cell.reference, elementIndex: string.elementIndex, lexical: string.lexical, namespaceUri: string.namespaceUri, preserveSpace: string.preserveSpace, ...(string.space ? { space: string.space } : {}) } });
      };
      if (captureSeriesOrder) for (const item of ordered) {
        const owner = item.kind === "line" && lineFamily ? lineFamily : family;
        const identity = r.one(item.node, c, "idx", true)!, order = r.one(item.node, c, "order", true)!;
        const rawId = await context.sourceElement(part, identity), rawOrder = await context.sourceElement(part, order);
        const simple = (node: XmlElement, expected: number) => node.namespaceUri === c && node.attributes.length === 1 && node.attributes[0]?.namespaceUri === "" && node.attributes[0]?.local === "val" && /^\d+$/u.test(node.attributes[0].value) && Number(node.attributes[0].value) === expected && node.children.every(child => child.kind === "text" && /^[\t\r\n ]*$/u.test(child.value));
        if (!await safeLabelOwner(order, item.node, owner) || rawId.node !== identity || rawId.parent !== item.node || rawOrder.node !== order || !simple(rawId.node, item.identity) || !simple(rawOrder.node, item.order)) { seriesOrderUnsafe = true; continue; }
        seriesOrderSources.push({ identity: item.identity, order: item.order, orderLexical: rawOrder.node.attributes[0]!.value, elementIndex: rawOrder.elementIndex, namespaceUri: c, family: owner.local });
      }
      if (chartType === "bubble") {
        const axes = scatterAxes(plot, family, c, r);
        if ("reason" in axes) return axes;
        const sizeMode = chartScalar(family, "sizeRepresents", c, r), scale = chartScalar(family, "bubbleScale", c, r);
        if ((sizeMode !== undefined && sizeMode !== "area") || (scale !== undefined && !/^0*100%?$/.test(scale)) ||
          chartBoolean(family, "bubble3D", c, r, false) || chartBoolean(family, "showNegBubbles", c, r, false) ||
          chartBoolean(family, "varyColors", c, r, false)) {
          return { reason: "Only 2D area-scaled bubbles at 100% with one color per series and no negative bubbles are mapped." };
        }
        if (ordered.length > 5461) return { reason: "Bubble series exceed the editable workbook column limit." };
        const series: BubbleChartSeriesInput[] = [];
        for (const [seriesIndex, { node }] of ordered.entries()) {
          await checkpoint();
          if (chartBoolean(node, "bubble3D", c, r, false)) return { reason: "3D bubble series are not mapped." };
          const tx = r.one(node, c, "tx"), x = r.one(node, c, "xVal"), y = r.one(node, c, "yVal"), size = r.one(node, c, "bubbleSize");
          if (!tx || !x || !y || !size) return { reason: "Bubble series require explicit names and independent X/Y/size sources." };
          const name = await resolve(tx, false);
          const denseBubble = (value: XmlElement) => resolve(value, true, undefined, false, true);
          const xv = await denseBubble(x), yv = await denseBubble(y), sv = await denseBubble(size);
          const padding = xv.trailingPadding ?? 0;
          if (name.values?.length !== 1 || !name.values[0] || !xv.values?.length || !yv.values || !sv.values ||
            xv.values.length + padding > 1_048_575 || xv.values.length !== yv.values.length || xv.values.length !== sv.values.length ||
            padding !== (yv.trailingPadding ?? 0) || padding !== (sv.trailingPadding ?? 0)) {
            return { reason: name.reason ?? xv.reason ?? yv.reason ?? sv.reason ?? "Bubble series require nonempty dense X/Y/size arrays of equal length." };
          }
          await collectSeriesName(node, tx, family, seriesIndex, name.values[0]);
          if (padding) note("chart.data-source", node, "Aligned trailing blank positions were omitted from all three bubble caches; source positions remain unchanged.");
          const xValues = xv.values.map(chartNumber), yValues = yv.values.map(chartNumber), sizes = sv.values.map(chartNumber);
          if (xValues.some(value => value === undefined) || yValues.some(value => value === undefined) ||
            sizes.some(value => value === undefined || value <= 0)) return { reason: "Bubble sizes must be positive finite numbers and all coordinates must be finite." };
          if (captureValues && v3Writes) {
            await bindNumericSeries(seriesIndex, tx, [["x", x, xValues], ["y", y, yValues], ["size", size, sizes]], true);
          }
          const properties = r.one(node, c, "spPr"); if (properties) r.known(properties);
          const paint = properties ? r.paint(properties) : undefined;
          if (paint && (paint.type !== "solid" || (paint.opacity !== undefined && paint.opacity !== 1))) {
            note("chart.series-style", properties!, "Bubble series preview uses a single opaque color; this fill is retained in source only.");
          }
          series.push({ name: name.values[0], xValues: xValues as number[], yValues: yValues as number[], sizes: sizes as number[],
            ...(paint?.type === "solid" ? { color: paint.color } : {}) });
        }
        const title = await chartTitle(chart, c, r, checkpoint, resolve);
        const chartProperties = r.one(root, c, "spPr"), plotProperties = r.one(plot, c, "spPr");
        if (chartProperties) r.known(chartProperties); if (plotProperties) r.known(plotProperties);
        const chartArea = chartProperties ? r.paint(chartProperties) : undefined, plotArea = plotProperties ? r.paint(plotProperties) : undefined, textStyle = chartTextProperties(root, c, r);
        note("chart.preview-layout", plot, "Bubble preview recomputes native layout and absolute diameters approximately. Source edits require independently bound X/Y/size caches and workbook cells.");
        return { fields: { type: "chart", chartType, box, series, ...axes, legend: chartLegend(chart, c, r),
          style: { ...(chartArea ? { chartArea } : {}), ...(plotArea ? { plotArea } : {}), ...(textStyle ? { textStyle } : {}) }, ...title }, previewFallback: true, ...(captureValues && v3Writes ? { valueBindings } : {}) };
      }
      if (chartType === "scatter") {
        const axes = scatterAxes(plot, family, c, r);
        if ("reason" in axes) return axes;
        const styleName = chartScalar(family, "scatterStyle", c, r);
        if (!styleName || !["line", "lineMarker", "marker"].includes(styleName)) return { reason: "Scatter style is absent or outside the supported straight-line/marker profile." };
        const scatterSeries: ScatterChartSeriesInput[] = [];
        for (const [seriesIndex, { node }] of ordered.entries()) {
          await checkpoint();
          const tx = r.one(node, c, "tx"), x = r.one(node, c, "xVal"), y = r.one(node, c, "yVal");
          if (!tx || !x || !y) return { reason: "Scatter series require explicit names and independent X/Y numeric sources." };
          if (chartBoolean(node, "smooth", c, r, false)) return { reason: "Smoothed scatter lines are outside the supported straight-line profile." };
          const name = await resolve(tx, false);
          let xs = await resolve(x, true, "declared"), ys = await resolve(y, true, "declared");
          // A missing cache may use a blank-aware workbook range only when the
          // independently parsed counterpart bounds its length. Never repair a
          // present sparse or malformed cache from workbook data.
          const missingCache = (target: XmlElement) => { const ref = resolved.get(target)?.ref; return ref?.formula !== undefined && !ref.reason && !ref.values && !ref.blankValues; };
          const yLength = (ys.values ?? ys.blankValues)?.length;
          if (!xs.values && !xs.blankValues && yLength !== undefined && missingCache(x)) xs = await resolve(x, true, yLength, true);
          const xLength = (xs.values ?? xs.blankValues)?.length;
          if (!ys.values && !ys.blankValues && xLength !== undefined && missingCache(y)) ys = await resolve(y, true, xLength, true);
          const xv = xs.values ?? xs.blankValues, yv = ys.values ?? ys.blankValues;
          if (!name.values || name.values.length !== 1 || !name.values[0] || !xv || !yv || xv.length !== yv.length) return { reason: name.reason ?? xs.reason ?? ys.reason ?? "Scatter names or coordinate counts cannot map to Core." };
          await collectSeriesName(node, tx, family, seriesIndex, name.values[0]);
          const numeric = (values: (string | null)[]) => values.map(v => v === null ? null : chartNumber(v));
          const xValues = numeric(xv), yValues = numeric(yv);
          if (xValues.some(v => v === undefined) || yValues.some(v => v === undefined)) return { reason: "Scatter coordinates contain nonfinite/error/unsupported values." };
          if (captureValues) await bindNumericSeries(seriesIndex, tx, [["x", x, xValues], ["y", y, yValues]], false);
          const properties = r.one(node, c, "spPr"); if (properties) r.known(properties);
          const stroke = properties ? r.stroke(properties) : undefined;
          const marker = chartMarker(node, c, r);
          scatterSeries.push({ name: name.values[0], xValues: xValues as (number | null)[], yValues: yValues as (number | null)[], ...(stroke ? { line: stroke.paint?.type === "none" ? false : stroke } : styleName === "marker" ? { line: false as const } : {}), ...(marker === undefined ? (styleName === "line" ? { marker: false as const } : {}) : { marker }) });
        }
        const title = await chartTitle(chart, c, r, checkpoint, resolve);
        const chartProperties = r.one(root, c, "spPr"), plotProperties = r.one(plot, c, "spPr");
        if (chartProperties) r.known(chartProperties); if (plotProperties) r.known(plotProperties);
        const chartArea = chartProperties ? r.paint(chartProperties) : undefined, plotArea = plotProperties ? r.paint(plotProperties) : undefined, textStyle = chartTextProperties(root, c, r);
        note("chart.preview-layout", plot, "Scatter preview recomputes native layout and styles approximately; numeric source edits require independently bound X/Y caches and workbook cells.");
        return { fields: { type: "chart", chartType, box, series: scatterSeries, ...axes, ...(supportedBlankMode === undefined ? {} : { blankDisplay: supportedBlankMode }), legend: chartLegend(chart, c, r), style: { ...(chartArea ? { chartArea } : {}), ...(plotArea ? { plotArea } : {}), ...(textStyle ? { textStyle } : {}) }, ...title }, previewFallback: true, ...(captureValues ? { valueBindings } : {}) };
      }
      const series: (ChartSeriesBaseInput | LineChartSeriesInput | PieChartSeriesInput | ComboChartSeriesInput)[] = [];
      let categories: string[] | undefined;
      for (const [seriesIndex, { node, kind }] of ordered.entries()) {
        await checkpoint();
        const tx = r.one(node, c, "tx"), cat = r.one(node, c, "cat"), val = r.one(node, c, "val", true)!;
        if (!tx || !cat) return { reason: "Explicit series names and categories are required; absent names/labels are not synthesized." };
        // A cache scalar is only owned by this series if every projected
        // ancestor resolves to the same direct chain in the original part.
        // Check once per series, rather than once per category point.
        const owner = kind === "line" && lineFamily ? lineFamily : family;
        const labelOwnerSafe = !captureLabelCandidates || await safeLabelOwner(cat, node, owner);
        const names = await resolve(tx, false), cats = await resolve(cat, false);
        const values = await resolve(val, true, cats.values?.length);
        const lexicalValues = values.values ?? values.blankValues;
        if (!names.values || names.values.length !== 1 || !names.values[0]?.length || !cats.values || cats.values.some(v => !v.length) || !lexicalValues) return { reason: names.reason ?? cats.reason ?? values.reason ?? "Series names/categories cannot map to nonempty Core strings." };
        const numbers = lexicalValues.map(value => value === null ? null : chartNumber(value));
        if (numbers.some(n => n === undefined || ((chartType === "pie" || chartType === "doughnut") && n !== null && n < 0))) return { reason: "Series contains nonfinite/error/unsupported values or negative pie/doughnut values; no values are coerced to zero." };
        if (numbers.length !== cats.values.length || (categories && (categories.length !== cats.values.length || categories.some((v, i) => v !== cats.values![i])))) return { reason: "Series category/value counts or category arrays disagree; source-only instead of aligning by position." };
        categories ??= cats.values;
        await collectSeriesName(node, tx, owner, seriesIndex, names.values[0]!);
        if (captureLabelCandidates) {
          if (seriesIndex === 0) for (let i = 0; i < cats.values.length; i++) { await bindingStep(); labelKeys.push(`category:${i}`); }
          const source = resolved.get(cat);
          for (let i = 0; i < cats.values.length; i++) {
            await bindingStep();
            const point = source?.ref.formula ? source.ref.points?.[i] : undefined, cell = source?.stored?.cells?.[i], string = cell?.cell.stringSource;
            if (!labelOwnerSafe || !point || !cell || !string || !bookPart || cell.cell.formula) continue;
            const raw = await rawLabelPoint(cat, point, source!.ref.formula!);
            if (!raw) continue;
            labelUses.push({ logicalKey: `category:${i}`, projected: cats.values[i]!, cache: { part, elementIndex: raw.elementIndex, lexical: point.lexical, namespaceUri: c }, cell: { workbookPart: bookPart, worksheetPart: string.part, reference: cell.reference, elementIndex: string.elementIndex, lexical: string.lexical, namespaceUri: string.namespaceUri, preserveSpace: string.preserveSpace, ...(string.space ? { space: string.space } : {}) } });
          }
        }
        if (captureValues && (!combo || v3Writes)) {
          const valuesRef = resolved.get(val), agree = (target: XmlElement, numeric: boolean) => {
            const v = resolved.get(target); return v?.ref.values && v.stored?.values && v.ref.values.length === v.stored.values.length && v.ref.values.every((text, i) => numeric ? chartNumericSourcesAgree(text, v.stored!.values![i]!) : text === v.stored!.values![i]);
          };
          const reason = unsafe ?? (!bookPart || !book ? "An unshared internal XLSX workbook is required." : !agree(tx, false) || !agree(cat, false) || !agree(val, true) ? "Complete caches and explicit local workbook name/category/value references must agree." : !valuesRef?.ref.formula || !valuesRef.ref.points ? "Existing numeric reference/cache points are required." : undefined);
          for (let pointIndex = 0; pointIndex < numbers.length; pointIndex++) {
            await bindingStep();
            const baselineValue = numbers[pointIndex];
            if (baselineValue === null || baselineValue === undefined) continue;
            const point = valuesRef?.ref.points?.[pointIndex], cell = valuesRef?.stored?.cells?.[pointIndex];
            const raw = point ? await context.sourceElement(part, point.node) : undefined;
            const scalar = raw && raw.node.children.every(n => n.kind === "text" || n.kind === "cdata");
            const blocked = reason ?? (!scalar || !cell?.cell.source || !cell.cell.numeric || cell.cell.formula ? "Cache/workbook value lacks a direct numeric scalar source." : undefined);
            valueBindings.push({ seriesIndex, pointIndex, baselineValue, ...(combo ? { v3Profile: true } : {}), source: { ...at(val), ...(raw ? { elementIndex: raw.elementIndex } : {}) },
              ...(raw ? { sourceElementIndex: raw.elementIndex } : {}), ...(point ? { cacheLexical: point.lexical, cacheNamespace: c } : {}),
              ...(bookPart && cell?.cell.source ? { workbook: { part: bookPart, worksheetPart: cell.cell.source.part, cellReference: cell.reference, sourceElementIndex: cell.cell.source.elementIndex }, workbookLexical: cell.cell.source.lexical, workbookNamespace: cell.cell.source.namespaceUri } : {}),
              status: blocked ? "unmapped" : "bound", ...(blocked ? { reason: blocked } : {}) });
          }
        }
        const properties = r.one(node, c, "spPr"); if (properties) r.known(properties);
        const paint = properties ? r.paint(properties) : undefined;
        const color = paint?.type === "solid" && paint.opacity === 1 ? paint.color : undefined;
        if (paint && color === undefined) r.note(properties!, "Series paint cannot map to Core's opaque series color.");
        const base = { name: names.values[0]!, values: numbers as (number | null)[], ...(color ? { color } : {}) };
        if (chartType === "combo") {
          if (kind === "line") {
            if (chartBoolean(node, "smooth", c, r, false)) return { reason: "Smoothed combo series are not projected." };
            const line = properties ? r.stroke(properties) : undefined, marker = chartMarker(node, c, r);
            series.push({ ...base, kind, ...(line ? { line } : {}), ...(marker === undefined ? {} : { marker }) });
          } else series.push({ ...base, kind });
        } else if (chartType === "line") {
          const line = properties ? r.stroke(properties) : undefined, marker = chartMarker(node, c, r);
          series.push({ ...base, ...(line ? { line } : {}), ...(marker === undefined ? {} : { marker }) });
        } else if (chartType === "pie" || chartType === "doughnut") {
          const indexed = new Map<number, string>();
          for (const point of children(node).filter(n => n.namespaceUri === c && n.local === "dPt")) {
            await checkpoint(); r.known(point);
            const index = r.number(r.one(point, c, "idx", true)!, "val", 0, 4294967295);
            if (index >= categories.length || indexed.has(index)) invalidStructure("Pie point identity is duplicated or outside categories.", part);
            // Track even points with unsupported colors, so duplicates cannot hide.
            indexed.set(index, "");
            const props = r.one(point, c, "spPr"); if (props) r.known(props);
            const fill = props ? r.paint(props) : undefined;
            if (fill?.type === "solid" && fill.opacity === 1 && typeof fill.color === "string") indexed.set(index, fill.color);
          }
          const complete = indexed.size === categories.length && [...indexed.values()].every(Boolean);
          if (indexed.size && !complete) r.note(node, `Partial/unsupported ${chartType} point colors remain source-only; Core palette colors are projection defaults.`);
          series.push({ ...base, ...(complete ? { pointColors: categories.map((_, i) => indexed.get(i)!) } : {}) });
        } else series.push(base);
      }
      const title = await chartTitle(chart, c, r, checkpoint, resolve);
      const chartProperties = r.one(root, c, "spPr"), plotProperties = r.one(plot, c, "spPr");
      if (chartProperties) r.known(chartProperties); if (plotProperties) r.known(plotProperties);
      const chartArea = chartProperties ? r.paint(chartProperties) : undefined, plotArea = plotProperties ? r.paint(plotProperties) : undefined;
      const textStyle = chartTextProperties(root, c, r);
      const style: ChartStyleInput = { ...(chartArea ? { chartArea } : {}), ...(plotArea ? { plotArea } : {}), ...(textStyle ? { textStyle } : {}) };
      const common = { type: "chart" as const, box, categories: categories!, ...(supportedBlankMode === undefined ? {} : { blankDisplay: supportedBlankMode }), legend: chartLegend(chart, c, r), style, ...title };
      note("chart.preview-layout", plot, "Native manual layout, inherited chart styles, number/date formatting, labels and unsupported properties are not reproduced. Core recomputes chart geometry; mapped preview is an approximation.");
      if (chartType === "pie") return { fields: { ...common, chartType, series: series as PieChartSeriesInput[] }, previewFallback: true, ...(captureValues ? { valueBindings } : {}) };
      if (chartType === "doughnut") return { fields: { ...common, chartType, ...doughnutOptions!, series: series as PieChartSeriesInput[] }, previewFallback: true, ...(captureValues ? { valueBindings } : {}) };
      if (chartType === "combo") {
        const axes = comboAxes(plot, family, lineFamily!, c, r);
        if ("reason" in axes) return axes;
        return { fields: { ...common, ...axes, chartType, series: series as ComboChartSeriesInput[] }, previewFallback: true, ...(captureValues && v3Writes ? { valueBindings } : {}) };
      }
      const axes = chartAxes(plot, family, c, r);
      if ("reason" in axes) return axes;
      if (chartType === "line") return { fields: { ...common, ...axes, chartType, series: series as LineChartSeriesInput[] }, previewFallback: true, ...(captureValues ? { valueBindings } : {}) };
      const stackOverflow = (separateSigns: boolean): boolean => {
        for (let i = 0; i < categories!.length; i++) {
          let positive = 0, negative = 0;
          for (const item of series) {
            const value = item.values[i] ?? 0;
            if (separateSigns && value < 0) negative += value; else positive += value;
            if (!Number.isFinite(positive) || !Number.isFinite(negative)) return true;
          }
        }
        return false;
      };
      if (chartType === "area") {
        if (areaGrouping === "stacked" && stackOverflow(false)) return { reason: "Area stack cumulative values exceed Core's finite range; source is retained." };
        return { fields: { ...common, ...axes, chartType, grouping: areaGrouping as "standard" | "stacked" | "percentStacked", series }, previewFallback: true, ...(captureValues ? { valueBindings } : {}) };
      }
      const direction = chartScalar(family, "barDir", c, r);
      if (direction !== "bar" && direction !== "col") return { reason: "Bar direction is absent or unsupported." };
      const grouping = chartScalar(family, "grouping", c, r) ?? "clustered";
      if (!["clustered", "stacked", "percentStacked"].includes(grouping)) return { reason: "Bar grouping is outside Core's supported groups." };
      if (grouping === "stacked" && stackOverflow(true)) return { reason: "Bar stack cumulative values exceed Core's finite range; source is retained." };
      const gapNode = r.one(family, c, "gapWidth"), overlapNode = r.one(family, c, "overlap");
      const categoryGap = gapNode ? r.number(gapNode, "val", 0, 500) : undefined;
      const overlap = overlapNode ? r.number(overlapNode, "val", -100, 100) : undefined;
      if (overlap !== undefined && ((grouping === "clustered" && overlap > 0) || (grouping !== "clustered" && overlap !== 100))) r.note(family, "Native overlap is outside Core's spacing model; projection uses Core spacing.");
      return { fields: { ...common, ...axes, chartType, series, orientation: direction === "bar" ? "horizontal" : "vertical", grouping: grouping as "clustered" | "stacked" | "percentStacked", ...(categoryGap === undefined ? {} : { categoryGap }), ...(overlap !== undefined && overlap <= 0 && grouping === "clustered" ? { seriesGap: -overlap } : {}) }, previewFallback: true, ...(captureValues ? { valueBindings } : {}) };
    };
    try {
      const result = await parse();
      if (captureValues && (formulas.get(part) ?? 0) !== interpretedFormulas) profileUnsafe = "Uninterpreted chart formula references are outside numeric preservation writes.";
      const labelCandidates = captureLabelCandidates && "fields" in result && labelKeys.length ? await chartLabelClosure(labelKeys, labelUses, cellUses, !profileUnsafe && !labelUnsafe, limits.maxChartBindingWork, bindingStep) : undefined;
      const seriesNameCandidates = captureSeriesNameCandidates && "fields" in result && nameKeys.length ? await chartLabelClosure(nameKeys, nameUses, cellUses, !profileUnsafe && !labelUnsafe, limits.maxChartBindingWork, bindingStep) : undefined;
      const seriesOrderCandidate = captureSeriesOrder && "fields" in result ? { part, sources: seriesOrderSources, complete: !profileUnsafe && !labelUnsafe && !seriesOrderUnsafe && !localDependencyReason && seriesOrderSources.length > 0, ...(legendProven ? { researchColumnLegend: true as const } : {}), ...(localDependencyReason ? { localDependencyReason } : {}) } : undefined;
      if (captureValues && "fields" in result && result.valueBindings) return { ...result, ...(labelCandidates ? { labelCandidates } : {}), valueBindings: result.valueBindings.map(b => {
        const reason = profileUnsafe ?? (b.workbook && cellUses.get(JSON.stringify([b.workbook.part, b.workbook.worksheetPart, b.workbook.cellReference])) !== 1 ? "Workbook numeric cell is shared by multiple chart references." : undefined);
        return b.status === "bound" && reason ? { ...b, status: "unmapped" as const, reason } : b;
      }), ...(seriesNameCandidates ? { seriesNameCandidates } : {}), ...(seriesOrderCandidate ? { seriesOrderCandidate } : {}) };
      return "fields" in result ? { ...result, ...(labelCandidates ? { labelCandidates } : {}), ...(seriesNameCandidates ? { seriesNameCandidates } : {}), ...(seriesOrderCandidate ? { seriesOrderCandidate } : {}) } : result;
    }
    finally { await r.audit(root, paths.get(root)!, (path, reason) => report("chart.unmapped-property", { part, path, sourceObjectId }, reason), checkpoint); }
  };
}

async function chartTitle(chart: XmlElement, c: string, r: ContentReader, checkpoint: () => Promise<void>, resolve: (node: XmlElement, numeric: boolean) => Promise<ChartReference>): Promise<{ title?: string; titleStyle?: ChartTextStyleInput }> {
  const node = r.one(chart, c, "title"); if (!node) return {};
  r.known(node);
  const tx = r.one(node, c, "tx"); if (!tx) return {};
  r.known(tx);
  const rich = r.one(tx, c, "rich");
  if (rich && children(tx).some(n => n.namespaceUri === c && ["strRef", "v"].includes(n.local))) invalidStructure("Conflicting chart title text choices.", r.part);
  if (!rich) { const value = await resolve(tx, false); return value.values?.length === 1 && value.values[0] ? { title: value.values[0] } : {}; }
  r.known(rich);
  const paragraphs: string[] = []; let titleStyle: ChartTextStyleInput | undefined;
  const styles = new Set<string>();
  for (const p of children(rich).filter(n => n.namespaceUri === r.a && n.local === "p")) {
    await checkpoint(); r.known(p); let text = "";
    for (const run of children(p)) {
      await checkpoint(); if (run.namespaceUri !== r.a) continue;
      if (run.local === "br") { r.known(run); text += "\n"; }
      if (run.local !== "r") continue;
      r.known(run); const value = r.one(run, r.a, "t", true)!; text += r.text(value);
      const style = chartTextStyle(r.one(run, r.a, "rPr"), r);
      titleStyle ??= style; styles.add(JSON.stringify(style ?? {}));
    }
    paragraphs.push(text);
  }
  if (styles.size > 1) r.note(rich, "Mixed title run formatting is flattened to one Core title style.");
  const title = paragraphs.join("\n");
  return { ...(title ? { title } : {}), ...(titleStyle ? { titleStyle } : {}) };
}
