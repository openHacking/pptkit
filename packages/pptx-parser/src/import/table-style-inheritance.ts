import type { NativeTableStylePreset, PaintInput, StrokeStyleInput, TableCellStyleInput, TableCellEdge } from "@pptkit/core";
import type { structureContext } from "../presentation/structure-context.js";
import { attribute, children, invalidStructure } from "../presentation/structure-xml.js";
import type { PptxInspection } from "../types/structure.js";
import type { PptxSourceLocation } from "../types/source.js";
import type { PptxTableStylePropertyInfo, PptxTableStyleSelection } from "../types/inheritance.js";
import { PptxImportError, type PptxSessionLimits, type PptxImportOptions } from "../types/import.js";
import type { XmlElement } from "../xml/read-xml.js";
import { contentPaths, contentReader, type ContentReader } from "./content-reader.js";
import { themeTextResolver, type ThemeTextScope } from "./theme-text-resolver.js";
import { tableStyleProperties, type TableStyleData as Data, type TableStyleDeclaration as Declaration } from "./table-style-properties.js";
import type { ContentTextInheritance } from "./content-text.js";
import { collectionTextStyles, referenceTextFonts, textFontFields, mergeTextStyles, type TextStylePatch } from "./content-text-style.js";

type Property = "cell.fill" | `cell.border.${TableCellEdge}`;
interface Region { fill?: Declaration; borders: Map<string, Declaration>; text: TextStylePatch }
const regions = ["wholeTbl", "band1H", "band2H", "band1V", "band2V", "firstCol", "lastCol", "firstRow", "lastRow", "nwCell", "neCell", "swCell", "seCell"] as const;
const sides = ["left", "right", "top", "bottom"] as const;
const guid = (value: string | undefined) => value && /^\{[\da-f]{8}-[\da-f]{4}-[\da-f]{4}-[\da-f]{4}-[\da-f]{12}\}$/i.test(value.trim()) ? value.trim().toUpperCase() : undefined;
export interface SourceTableCellStyle {
  style: TableCellStyleInput;
  text: ContentTextInheritance;
  /** Record effective final cell values after direct cell declarations. */
  record: (style: TableCellStyleInput, edges: Partial<Record<TableCellEdge, StrokeStyleInput>>) => void;
}
export interface SourceTableStyle {
  nativePreset?: NativeTableStylePreset;
  inherited: boolean;
  paint: (node: XmlElement) => Promise<PaintInput | undefined>;
  stroke: (node: XmlElement, name: string) => Promise<StrokeStyleInput | undefined>;
  copyStroke: (source: StrokeStyleInput, target: StrokeStyleInput) => void;
  blockStroke: (node: XmlElement, reason: string) => StrokeStyleInput;
  cell: (node: XmlElement, row: number, column: number, rowSpan: number, colSpan: number, rows: number, columns: number) => Promise<SourceTableCellStyle>;
  conditional: boolean;
}

/** Source identities/rules stay in Parser; Core receives only effective direct values. */
export function tableStyleInheritance(context: Awaited<ReturnType<typeof structureContext>>, inspection: PptxInspection, limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>, unmapped: (source: PptxSourceLocation, reason: string) => void, fontProfile?: PptxImportOptions["textFontProfile"]) {
  const tables: PptxTableStyleSelection[] = [], properties: PptxTableStylePropertyInfo[] = [];
  let records = 0, dependencies = 0, work = 0;
  const step = async (count = 1) => { work += count; if (work > limits.maxTableStyleWork) throw new PptxImportError("import-limit-exceeded", "Table style traversal exceeds maxTableStyleWork."); await checkpoint(); };
  const charge = (source: PptxSourceLocation, count = 0) => {
    if (++records > limits.maxTableStyleProperties || count > limits.maxTableStyleDependencies - dependencies) throw new PptxImportError("import-limit-exceeded", "Table style provenance exceeds table property/dependency limits.", source);
    dependencies += count;
  };
  const report = (info: PptxTableStylePropertyInfo) => { charge(info.target, info.resolvedFrom?.length ?? 0); properties.push(info); };
  const theme = themeTextResolver(context, inspection, step, unmapped), layouts = new Map(inspection.layouts.map(layout => [layout.part, layout]));
  const external = new Map<string, Promise<{ data: Data; entries: Map<string, XmlElement>; rawIds: () => Promise<Set<string>>; defaultId?: string }>>();
  const regionCache = new WeakMap<XmlElement, Map<ThemeTextScope, Promise<Map<string, Region>>>>();
  const { at, blocked, paint, stroke: readStroke, matrixReference } = tableStyleProperties(context.a, step, unmapped);
  const registry = (part: string) => {
    let pending = external.get(part); if (pending) return pending;
    pending = (async () => {
      const root = await context.root(part, "tblStyleLst", "presentationml.tableStyles", context.a), data = { root, part, r: contentReader(part, context.p, context.a), paths: await contentPaths(root, step) }, entries = new Map<string, XmlElement>();
      data.r.known(root); const defaultId = guid(data.r.attr(root, "def"));
      for (const node of children(root)) {
        await step(); if (node.namespaceUri !== context.a || node.local !== "tblStyle") continue;
        const id = guid(attribute(node, "styleId"));
        if (!id) { unmapped(at(data, node), "A table style definition has no supported GUID identity."); continue; }
        if (entries.has(id)) invalidStructure("Ambiguous duplicate table style identities.", part);
        entries.set(id, node);
      }
      let rawInventory: Promise<Set<string>> | undefined;
      const rawIds = (): Promise<Set<string>> => rawInventory ??= (async () => {
        const ids = new Set<string>(), rawRoot = context.rawXml.get(part);
        if (!rawRoot) throw new PptxImportError("import-invalid-package", "Table style source XML is unavailable.");
        const pending = [rawRoot];
        while (pending.length) {
          await step();
          const node = pending.pop()!;
          if (node.namespaceUri === context.a && node.local === "tblStyle") {
            const id = guid(attribute(node, "styleId")); if (id) ids.add(id);
          }
          for (const child of children(node)) pending.push(child);
        }
        return ids;
      })();
      return { data, entries, rawIds, ...(defaultId ? { defaultId } : {}) };
    })(); external.set(part, pending); return pending;
  };
  const parseRegions = (node: XmlElement, data: Data, scope: ThemeTextScope) => {
    const scopes = regionCache.get(node) ?? new Map<ThemeTextScope, Promise<Map<string, Region>>>(); regionCache.set(node, scopes);
    let pending = scopes.get(scope); if (pending) return pending;
    pending = (async () => {
      // External definitions can be evaluated under several target themes.
      // Keep diagnostic/consumption state local to that evaluation; an
      // unsupported palette in one scope must not taint another scope's audit.
      if (data.root?.local === "tblStyleLst") data = { ...data, r: contentReader(data.part, context.p, context.a) };
      const result = new Map<string, Region>(), r = data.r; r.known(node); r.attr(node, "styleId"); r.attr(node, "styleName");
      for (const name of regions) {
        await step(); const part = r.one(node, context.a, name); if (!part) continue; r.known(part);
        const cell = r.one(part, context.a, "tcStyle"), tx = r.one(part, context.a, "tcTxStyle"), region: Region = { borders: new Map(), text: new Map() };
        if (cell) {
          r.known(cell); const fill = r.one(cell, context.a, "fill"), ref = r.one(cell, context.a, "fillRef");
          if (fill && ref) invalidStructure("Conflicting table region fill choices.", data.part);
          if (fill) { r.known(fill); const decl = await paint(fill, data, scope, name); if (decl) region.fill = decl; }
          if (ref) region.fill = await matrixReference(ref, "fill", data, scope, name);
          const borders = r.one(cell, context.a, "tcBdr");
          if (borders) { r.known(borders); for (const side of [...sides, "insideH", "insideV"] as const) {
            await step(); const edge = r.one(borders, context.a, side); if (!edge) continue; r.known(edge);
            const line = r.one(edge, context.a, "ln"), ref = r.one(edge, context.a, "lnRef"); if (line && ref) invalidStructure("Conflicting table region line choices.", data.part);
            const decl = line ? await readStroke(edge, "ln", data, scope, name) : ref ? await matrixReference(ref, "line", data, scope, name) : blocked(data, edge, "Missing table line blocks an older border.", name);
            if (decl) region.borders.set(side, decl);
          } }
        }
        if (tx) {
          r.known(tx);
          const set = (key: "run.bold" | "run.italic" | "run.fontFamily" | "run.color" | "run.colorOpacity", decl: Declaration) => region.text.set(key, { source: decl.source, level: "table-style", status: decl.status, ...(decl.value !== undefined ? { value: decl.value as never } : {}), ...(decl.reason ? { reason: decl.reason } : {}), ...(decl.resolvedFrom ? { resolvedFrom: decl.resolvedFrom } : {}) });
          for (const [attr, key] of [["b", "run.bold"], ["i", "run.italic"]] as const) {
            const value = r.attr(tx, attr); if (value === undefined || value === "def") continue;
            if (value === "on" || value === "off") set(key, { source: at(data, tx, attr), status: "mapped", value: value === "on" });
            else set(key, blocked(data, tx, `Unsupported table text ${attr} blocks an older flag.`, name));
          }
          const font = r.one(tx, context.a, "font"), fontRef = r.one(tx, context.a, "fontRef"); if (font && fontRef) invalidStructure("Conflicting table font choices.", data.part);
          if (font) {
            const patch = await collectionTextStyles(font, { r, level: "table-style", ...(fontProfile ? { fontProfile } : {}), at: (n, attr) => at(data, n, attr), resolveFont: value => scope.font(value, fontProfile !== undefined, (fontProfile === "script-families-v2" || fontProfile === "script-families-v3")) }, step);
            for (const key of textFontFields) { const family = patch.get(key); if (family) region.text.set(key, family); }
          }
          if (fontRef) {
            if (fontProfile) {
              const fonts = await referenceTextFonts(fontRef, { r, level: "table-style", fontProfile, at: (n, attr) => at(data, n, attr), resolveFont: value => scope.font(value, true, (fontProfile === "script-families-v2" || fontProfile === "script-families-v3")) });
              for (const [key, decl] of fonts) region.text.set(key, decl);
            } else {
              const index = r.attr(fontRef, "idx");
              const resolved = index === "major" || index === "minor" ? await scope.font(index === "major" ? "+mj-lt" : "+mn-lt")
                : { status: "unsupported" as const, reason: "Unsupported table font reference blocks an older Latin font." };
              set("run.fontFamily", { source: at(data, fontRef, "idx"), ...resolved });
            }
          }
          // A reference's color supplies a default; the region's explicit color
          // is later, and local list/paragraph/run declarations remain later still.
          for (const owner of [fontRef, tx]) {
            if (!owner) continue;
            const colors = children(owner).filter(n => n.namespaceUri === context.a && ["srgbClr", "schemeClr", "sysClr", "scrgbClr", "hslClr", "prstClr"].includes(n.local));
            if (colors.length > 1) invalidStructure("Conflicting table text colors.", data.part);
            if (!colors[0]) continue;
            // A virtual fill only adapts the shared color evaluator; provenance
            // and audit continue to refer to the real color declaration.
            const holder: XmlElement = { ...owner, children: [{ ...owner, local: "solidFill", children: [colors[0]] }] };
            const decl = await paint(holder, data, scope, name);
            if (decl) {
              if (decl.status === "mapped") { const p = decl.value as Extract<PaintInput, { type: "solid" }>; set("run.color", { ...decl, value: p.color as string }); set("run.colorOpacity", { ...decl, value: p.opacity ?? 1 }); }
              else { set("run.color", decl); set("run.colorOpacity", decl); }
            }
          }
        }
        result.set(name, region);
      }
      await data.r.audit(node, data.paths.get(node)!, (path, reason) => unmapped({ part: data.part, path }, reason), step);
      return result;
    })(); scopes.set(scope, pending); return pending;
  };
  const forTable = async (props: XmlElement | undefined, r: ContentReader, paths: WeakMap<XmlElement, string>, source: PptxSourceLocation, layoutPart?: string): Promise<SourceTableStyle> => {
    const local: Data = { part: r.part, r, paths, ...(props ? { root: props } : {}) }, masterPart = layoutPart ? layouts.get(layoutPart)?.masterPart : undefined, scope = theme.forMaster(masterPart, layoutPart, r.part);
    const flags: Record<string, boolean> = {};
    for (const name of ["firstRow", "lastRow", "firstCol", "lastCol", "bandRow", "bandCol"]) flags[name] = props ? r.bool(props, name) : false;
    const inline = props ? r.one(props, context.a, "tableStyle") : undefined, reference = props ? r.one(props, context.a, "tableStyleId") : undefined;
    if (inline && reference) invalidStructure("Conflicting inline/reference table style choices.", r.part);
    let selected: XmlElement | undefined, data = local, reason: string | undefined, id: string | undefined;
    let rawStyleIds: (() => Promise<ReadonlySet<string>>) | undefined;
    if (inline) {
      selected = inline; id = guid(attribute(inline, "styleId"));
      if (!id) reason = "Inline table style identity is unsupported; direct declarations are projected without adopting its identity.";
    }
    else {
      id = reference ? guid(r.text(reference)) : undefined;
      const part = context.optionalTarget(inspection.mainPart, "tableStyles");
      if (part) { const list = await registry(part); data = list.data; rawStyleIds = list.rawIds; id ??= reference ? undefined : list.defaultId; if (id) selected = list.entries.get(id); }
      if (reference && !id) reason = "Unsupported table style GUID reference.";
      else if (!selected && (reference || part)) reason = "The selected/default table style definition is unavailable; no built-in style is guessed.";
    }
    if (inline || reference || selected || reason) {
      charge(source); tables.push({ source, flags: flags as PptxTableStyleSelection["flags"], status: inline ? "inline" : selected ? "resolved" : reason?.startsWith("Unsupported") ? "unsupported" : "missing", ...(id ? { styleId: id } : {}), ...(selected ? { definition: at(data, selected), ...(attribute(selected, "styleName") !== undefined ? { styleName: attribute(selected, "styleName")! } : {}) } : {}), ...(reason ? { reason } : {}) });
      if (reason) r.note(reference ?? props!, reason);
    }
    // A custom definition can use a built-in GUID. Only bare explicit references
    // are eligible for named transport; the Parser still does not infer their appearance.
    const nativePresetCandidate = reference && !selected ? id === "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}" ? "medium-2-accent-2"
      : id === "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}" ? "medium-2-accent-1" : undefined : undefined;
    const nativePreset = nativePresetCandidate && id && rawStyleIds && (await rawStyleIds()).has(id) ? undefined : nativePresetCandidate;
    const parsed = selected ? await parseRegions(selected, data, scope) : new Map<string, Region>();
    const strokeOrigins = new WeakMap<StrokeStyleInput, Declaration>(), paintOrigins = new WeakMap<XmlElement, Declaration>();
    const readPaint = async (node: XmlElement) => { const decl = await paint(node, local, scope); if (!decl) return undefined; paintOrigins.set(node, decl); return decl.status === "mapped" ? decl.value as PaintInput : { type: "none" as const }; };
    const readLine = async (node: XmlElement, name: string) => { const decl = await readStroke(node, name, local, scope); if (!decl) return undefined; const value: StrokeStyleInput = decl.status === "mapped" ? decl.value as StrokeStyleInput : { paint: { type: "none" }, width: 0 }; strokeOrigins.set(value, decl); return value; };
    return {
      ...(nativePreset === undefined ? {} : { nativePreset }),
      inherited: selected !== undefined,
      conditional: Boolean(selected && Object.values(flags).some(Boolean)), paint: readPaint, stroke: readLine,
      copyStroke(from, to) { const decl = strokeOrigins.get(from); if (decl) strokeOrigins.set(to, decl); },
      blockStroke(node, reason) { const value: StrokeStyleInput = { paint: { type: "none" }, width: 0 }; strokeOrigins.set(value, blocked(local, node, reason)); return value; },
      async cell(node, row, column, rowSpan, colSpan, rows, columns) {
        await step(); const patch = new Map<Property, Declaration>(), texts: TextStylePatch[] = [];
        const top = row === 0, bottom = row + rowSpan === rows, left = column === 0, right = column + colSpan === columns;
        const horizontal = flags.bandRow && !(flags.firstRow && top) && !(flags.lastRow && row === rows - 1), vertical = flags.bandCol && !(flags.firstCol && left) && !(flags.lastCol && column === columns - 1);
        const selected = ["wholeTbl", ...(horizontal ? [(row - (flags.firstRow ? 1 : 0)) % 2 === 0 ? "band1H" : "band2H"] : []), ...(vertical ? [(column - (flags.firstCol ? 1 : 0)) % 2 === 0 ? "band1V" : "band2V"] : []), ...(flags.firstCol && left ? ["firstCol"] : []), ...(flags.lastCol && right ? ["lastCol"] : []), ...(flags.firstRow && top ? ["firstRow"] : []), ...(flags.lastRow && bottom ? ["lastRow"] : []), ...(flags.firstRow && flags.firstCol && top && left ? ["nwCell"] : []), ...(flags.firstRow && flags.lastCol && top && right ? ["neCell"] : []), ...(flags.lastRow && flags.firstCol && bottom && left ? ["swCell"] : []), ...(flags.lastRow && flags.lastCol && bottom && right ? ["seCell"] : [])];
        for (const name of selected) {
          await step(); const region = parsed.get(name); if (!region) continue; if (region.fill) patch.set("cell.fill", region.fill); texts.push(region.text);
          for (const side of sides) {
            const rowStart = name === "firstRow" || name === "nwCell" || name === "neCell" ? 0 : name === "lastRow" || name === "swCell" || name === "seCell" ? rows - 1 : name.endsWith("H") ? row : 0;
            const rowEnd = name === "wholeTbl" || name.endsWith("V") || name === "firstCol" || name === "lastCol" ? rows : rowStart + 1;
            const columnStart = name === "firstCol" || name === "nwCell" || name === "swCell" ? 0 : name === "lastCol" || name === "neCell" || name === "seCell" ? columns - 1 : name.endsWith("V") ? column : 0;
            const columnEnd = name === "wholeTbl" || name.endsWith("H") || name === "firstRow" || name === "lastRow" ? columns : columnStart + 1;
            const exterior = side === "left" ? column === columnStart : side === "right" ? column + colSpan === columnEnd : side === "top" ? row === rowStart : row + rowSpan === rowEnd;
            const key = exterior ? side : side === "left" || side === "right" ? "insideV" : "insideH";
            const decl = region.borders.get(key); if (decl) patch.set(`cell.border.${side}`, decl);
          }
        }
        const style: TableCellStyleInput = {};
        for (const [key, decl] of patch) {
          if (key === "cell.fill") style.fill = decl.status === "mapped" ? decl.value as PaintInput : { type: "none" };
          else { style.borders ??= {}; style.borders[key.slice(12) as TableCellEdge] = decl.status === "mapped" ? decl.value as StrokeStyleInput : { paint: { type: "none" }, width: 0 }; }
        }
        const target = at(local, node), recordProperty = (property: PptxTableStylePropertyInfo["property"], decl: Declaration, to = target) => report({ ...decl, property, target: to, rowIndex: row, columnIndex: column });
        const text: ContentTextInheritance = {
          frame: new Map(), paragraph: () => mergeTextStyles(...texts),
          local: { r, level: "slide-list-style", ...(fontProfile ? { fontProfile } : {}), unmappedFontSelection: unmapped, at: (n, attr) => at(local, n, attr), resolveColor: color => scope.color(r.part, color, paths), resolveFont: value => scope.font(value, fontProfile !== undefined, (fontProfile === "script-families-v2" || fontProfile === "script-families-v3")) },
          record(n, values, prefix) { for (const [key, value] of values) if (key.startsWith(prefix)) recordProperty(key, value, at(local, n)); },
        };
        return { style, text, record(finalStyle, edges) {
          const props = r.one(node, context.a, "tcPr"), direct = props ? paintOrigins.get(props) : undefined;
          if (direct) patch.set("cell.fill", direct);
          for (const side of sides) { const edge = edges[side], decl = edge ? strokeOrigins.get(edge) : undefined; if (decl) patch.set(`cell.border.${side}`, decl); }
          for (const [key, decl] of patch) {
            const value = key === "cell.fill" ? finalStyle.fill : finalStyle.borders?.[key.slice(12) as TableCellEdge] ?? finalStyle.stroke;
            recordProperty(key, decl.status === "mapped" && value !== undefined ? { ...decl, value } : decl);
          }
        } };
      },
    };
  };
  return { forTable, report: () => ({ profile: "drawingml-direct-regions-v1" as const, tables, properties }) };
}
