import type { structureContext } from "../presentation/structure-context.js";
import type { PptxInspection } from "../types/structure.js";
import type { XmlElement } from "../xml/read-xml.js";
import { PptxImportError, type PptxSessionLimits, type PptxSourceLocation, type PptxImportOptions } from "../types/import.js";
import type { PptxPlaceholderInfo, PptxPlaceholderLevel, PptxTextStylePropertyInfo } from "../types/inheritance.js";
import { contentPaths, contentReader } from "./content-reader.js";
import type { ContentTextInheritance } from "./content-text.js";
import { frameTextStyles, listTextStyles, mergeTextStyles, type TextStyleContext, type TextStylePatch } from "./content-text-style.js";
import type { ThemeTextScope } from "./theme-text-resolver.js";
import { themeTextResolver } from "./theme-text-resolver.js";
import { notesMasterBody } from "./notes-master-body.js";
import { children } from "../presentation/structure-xml.js";

interface Ancestor {
  node: XmlElement;
  part: string;
  level: PptxPlaceholderLevel;
  paths: WeakMap<XmlElement, string>;
  source: PptxSourceLocation;
}
interface List { defaults: TextStylePatch; levels: TextStylePatch[] }
interface Body { frame: TextStylePatch; list: List }

/** Format-layer cascade; selected values are materialized as ordinary Core styles. */
export function textInheritance(
  context: Awaited<ReturnType<typeof structureContext>>, inspection: PptxInspection,
  limits: Required<PptxSessionLimits>, checkpoint: () => Promise<void>,
  sourcePaths: (part: string) => WeakMap<XmlElement, string> | undefined,
  unmapped: (source: PptxSourceLocation, reason: string) => void,
  fontProfile?: PptxImportOptions["textFontProfile"],
  notesUnmapped: (source: PptxSourceLocation, reason: string) => void = unmapped,
) {
  const properties: PptxTextStylePropertyInfo[] = [];
  let dependencies = 0;
  let notesWork = 0;
  const notesStep = async (count = 1) => {
    if (count > limits.maxNotesStyleWork - notesWork) throw new PptxImportError("import-limit-exceeded", "Notes style traversal exceeds maxNotesStyleWork.");
    notesWork += count; await checkpoint();
  };
  const reportedUnmapped = new Set<string>();
  const reportUnmapped = (source: PptxSourceLocation, reason: string) => {
    const key = JSON.stringify([source.part, source.path, source.sourceObjectId, source.relationshipId, reason]);
    if (reportedUnmapped.has(key)) return;
    reportedUnmapped.add(key);
    unmapped(source, reason);
  };
  const partCache = new Map<string, { r: ReturnType<typeof contentReader>; paths: WeakMap<XmlElement, string>; root: XmlElement }>();
  const lists = new WeakMap<XmlElement, Map<ThemeTextScope, List>>(), bodies = new WeakMap<XmlElement, Map<ThemeTextScope, Body>>();
  const layouts = new Map(inspection.layouts.map(layout => [layout.part, layout]));
  const theme = themeTextResolver(context, inspection, checkpoint, reportUnmapped);
  const reportedNotes = new Set<string>();
  const reportNotes = (source: PptxSourceLocation, reason: string) => {
    const key = JSON.stringify([source, reason]); if (reportedNotes.has(key)) return; reportedNotes.add(key); notesUnmapped(source, reason);
  };
  const notesTheme = themeTextResolver(context, inspection, notesStep, reportNotes);
  const notesBodies = new Map<string, Promise<Awaited<ReturnType<typeof notesMasterBody>>>>();
  const scopedReaders = new WeakMap<ReturnType<typeof contentReader>, WeakMap<ThemeTextScope, ReturnType<typeof contentReader>>>();
  const empty: List = { defaults: new Map(), levels: Array.from({ length: 9 }, () => new Map()) };
  const part = async (name: string, rootName: string) => {
    const cached = partCache.get(name); if (cached) return cached;
    const root = await context.root(name, rootName);
    const value = { root, r: contentReader(name, context.p, context.a), paths: sourcePaths(name) ?? await contentPaths(root, checkpoint) };
    partCache.set(name, value); return value;
  };
  const styleContext = (r: ReturnType<typeof contentReader>, paths: WeakMap<XmlElement, string>, level: TextStyleContext["level"], owner: PptxSourceLocation, scope: ThemeTextScope, isolate = true): TextStyleContext => {
    let reader = r;
    if ((fontProfile || level.startsWith("notes-master")) && isolate) {
      const scopes = scopedReaders.get(r) ?? new WeakMap<ThemeTextScope, ReturnType<typeof contentReader>>();
      reader = scopes.get(scope) ?? contentReader(r.part, r.p, r.a); scopes.set(scope, reader); scopedReaders.set(r, scopes);
    }
    return {
      r: reader, level, themeScope: scope, ...(fontProfile ? { fontProfile } : {}),
      unmappedFontSelection: level.startsWith("notes-") ? reportNotes : reportUnmapped,
      at: (node, attr) => ({ ...owner, path: `${paths.get(node)!}${attr ? `/@${attr}` : ""}` }),
      resolveColor: async node => {
        const result = await scope.color(r.part, node, paths);
        if (level.startsWith("notes-")) await notesStep(result.status === "mapped" ? Math.max(1, result.program?.modifiers.length ?? 0) : 1);
        return result;
      },
      resolveFont: async typeface => { if (level.startsWith("notes-")) await notesStep(); return scope.font(typeface, fontProfile !== undefined, (fontProfile === "script-families-v2" || fontProfile === "script-families-v3")); },
    };
  };
  const record = (local: TextStyleContext, target: XmlElement, patch: TextStylePatch, prefix: "frame." | "paragraph." | "run.") => {
    for (const [property, value] of patch) {
      if (!property.startsWith(prefix)) continue;
      if (properties.length >= limits.maxTextStyleProperties) throw new PptxImportError("import-limit-exceeded", "Text style provenance exceeds maxTextStyleProperties.", local.at(target));
      const count = value.resolvedFrom?.length ?? 0;
      if (count > limits.maxTextStyleDependencies - dependencies) throw new PptxImportError("import-limit-exceeded", "Text style provenance exceeds maxTextStyleDependencies.", local.at(target));
      dependencies += count;
      properties.push({ property, target: local.at(target), ...value });
    }
  };
  const audit = async (node: XmlElement, c: TextStyleContext) => c.r.audit(node, c.at(node).path!, (path, reason) => (c.level.startsWith("notes-") ? reportNotes : reportUnmapped)({ ...c.at(node), path }, reason), c.level.startsWith("notes-") ? notesStep : checkpoint);
  const list = async (node: XmlElement | undefined, c: TextStyleContext): Promise<List> => {
    if (!node) return empty;
    const themeScope = c.themeScope as ThemeTextScope;
    const byScope = lists.get(node), cached = byScope?.get(themeScope); if (cached) return cached;
    await (c.level.startsWith("notes-") ? notesStep : checkpoint)();
    const result = await listTextStyles(node, c);
    await audit(node, c);
    const scopes = byScope ?? new Map<ThemeTextScope, List>();
    scopes.set(themeScope, result); lists.set(node, scopes); return result;
  };
  const ancestorBody = async (ancestor: Ancestor, scope: ThemeTextScope): Promise<Body> => {
    const cached = bodies.get(ancestor.node)?.get(scope); if (cached) return cached;
    const data = await part(ancestor.part, ancestor.level === "master" ? "sldMaster" : "sldLayout");
    const c = styleContext(data.r, ancestor.paths, ancestor.level === "master" ? "master-placeholder" : "layout-placeholder", ancestor.source, scope);
    const body = data.r.one(ancestor.node, context.p, "txBody"), props = body ? data.r.one(body, context.a, "bodyPr") : undefined;
    const frame = frameTextStyles(props, c);
    if (props) await audit(props, c);
    const result = { frame, list: await list(body ? data.r.one(body, context.a, "lstStyle") : undefined, c) };
    const scopes = bodies.get(ancestor.node) ?? new Map<ThemeTextScope, Body>();
    scopes.set(scope, result); bodies.set(ancestor.node, scopes); return result;
  };
  const forDrawing = async (
    layoutPart: string | undefined, source: PptxSourceLocation,
    r: ReturnType<typeof contentReader>, paths: WeakMap<XmlElement, string>,
    placeholder: PptxPlaceholderInfo | undefined, ancestors: readonly Ancestor[],
  ): Promise<ContentTextInheritance> => {
    const layers: List[] = [];
    const masterPart = layoutPart ? layouts.get(layoutPart)?.masterPart : undefined;
    const scope = theme.forMaster(masterPart, layoutPart, source.part);
    // The master's named text style is independent of the presence of an
    // individual master placeholder. Prompt paragraphs are never style parents.
    if (masterPart && !placeholder?.grouped) {
      const data = await part(masterPart, "sldMaster");
      const txStyles = data.r.one(data.root, context.p, "txStyles");
      const type = placeholder?.type;
      const name = type === "title" || type === "ctrTitle" ? "titleStyle" : !type || ["dt", "ftr", "sldNum", "hdr", "sldImg"].includes(type) ? "otherStyle" : "bodyStyle";
      const c = styleContext(data.r, data.paths, "master-text-style", { part: masterPart }, scope);
      layers.push(await list(txStyles ? data.r.one(txStyles, context.p, name) : undefined, c));
    }
    // Plain shapes use presentation defaults before falling back to the
    // master's otherStyle. Placeholder defaults come from their master style.
    if (!placeholder) {
      const data = await part(inspection.mainPart, "presentation");
      layers.push(await list(data.r.one(data.root, context.p, "defaultTextStyle"), styleContext(data.r, data.paths, "presentation-default", { part: inspection.mainPart }, scope)));
    }
    let frame: TextStylePatch = new Map();
    for (const ancestor of ancestors) {
      await checkpoint();
      const body = await ancestorBody(ancestor, scope);
      frame = mergeTextStyles(frame, body.frame); layers.push(body.list);
    }
    const byLevel = new Map<number, TextStylePatch>();
    const local = styleContext(r, paths, "slide-list-style", source, scope, false);
    return {
      local, frame,
      paragraph(level) {
        const cached = byLevel.get(level); if (cached) return cached;
        const result = mergeTextStyles(...layers.flatMap(layer => [layer.defaults, layer.levels[level]!]));
        byLevel.set(level, result); return result;
      },
      record(target, patch, prefix) {
        record(local, target, patch, prefix);
      },
    };
  };
  const forLocal = (r: ReturnType<typeof contentReader>, paths: WeakMap<XmlElement, string>, source: PptxSourceLocation): ContentTextInheritance => {
    const local: TextStyleContext = { r, level: "notes-direct", ...(fontProfile ? { fontProfile } : {}), at: (node, attr) => ({ ...source, path: `${paths.get(node)!}${attr ? `/@${attr}` : ""}` }) };
    return { local, frame: new Map(), paragraph: () => new Map(), record(target, patch, prefix) { record(local, target, patch, prefix); } };
  };
  const forArtwork = (masterPart: string, layoutPart: string | undefined, r: ReturnType<typeof contentReader>, paths: WeakMap<XmlElement, string>, source: PptxSourceLocation): ContentTextInheritance => {
    const scope = theme.forArtwork(masterPart, layoutPart), local = styleContext(r, paths, "slide-list-style", source, scope, false);
    return { local, frame: new Map(), paragraph: () => new Map(), record(target, patch, prefix) { record(local, target, patch, prefix); } };
  };
  const forSlideArtwork = (masterPart: string | undefined, layoutPart: string | undefined, slidePart: string, r: ReturnType<typeof contentReader>, paths: WeakMap<XmlElement, string>, source: PptxSourceLocation): ContentTextInheritance => {
    const scope = theme.forMaster(masterPart, layoutPart, slidePart), local = styleContext(r, paths, "slide-list-style", source, scope, false);
    return { local, frame: new Map(), paragraph: () => new Map(), record(target, patch, prefix) { record(local, target, patch, prefix); } };
  };
  const forNotes = async (r: ReturnType<typeof contentReader>, paths: WeakMap<XmlElement, string>, source: PptxSourceLocation, placeholder: XmlElement): Promise<ContentTextInheritance> => {
    await notesStep();
    const masterPart = context.optionalTarget(source.part, "notesMaster"), scope = notesTheme.forNotes(masterPart, source.part);
    const local = styleContext(r, paths, "notes-direct", source, scope, false), layers: List[] = [];
    const matchDependencies: PptxSourceLocation[] = [];
    if (masterPart) {
      const data = await part(masterPart, "notesMaster");
      layers.push(await list(data.r.one(data.root, context.p, "notesStyle"), styleContext(data.r, data.paths, "notes-master-text-style", { part: masterPart }, scope)));
      let pending = notesBodies.get(masterPart);
      if (!pending) { pending = notesMasterBody(data.root, data.r, notesStep); notesBodies.set(masterPart, pending); }
      const { count, candidate } = await pending;
      if (count === 1 && candidate && !candidate.grouped) {
        const owner = { part: masterPart, sourceObjectId: candidate.sourceId }, c = styleContext(data.r, data.paths, "notes-master-placeholder", owner, scope);
        matchDependencies.push(local.at(placeholder, "type"), c.at(candidate.placeholder, "type"));
        const body = data.r.one(candidate.node, context.p, "txBody");
        layers.push(await list(body ? data.r.one(body, context.a, "lstStyle") : undefined, c));
        if (body) for (const node of children(body)) { await notesStep(); if (!(node.namespaceUri === context.a && node.local === "lstStyle")) reportNotes(c.at(node), "Notes-master prompt paragraphs and frame/page properties are not inherited as Core speaker-note content."); }
      } else reportNotes(source, count > 1 ? "Ambiguous notes-master body placeholder type; no placeholder-specific text defaults are selected." : candidate?.grouped ? "Grouped notes-master body placeholder has no admitted text-inheritance policy." : "Notes master has no body placeholder type match; only its explicit notesStyle can supply text defaults.");
    } else reportNotes(source, "Notes slide has no notes-master relationship; presentation/slide master themes and styles are not substituted.");
    const byLevel = new Map<number, TextStylePatch>();
    return { local, frame: new Map(), paragraph(level) {
      let patch = byLevel.get(level);
      if (!patch) {
        patch = mergeTextStyles(...layers.flatMap(layer => [layer.defaults, layer.levels[level]!]));
        if (matchDependencies.length) patch = new Map([...patch].map(([key, declaration]) => [key, declaration.level === "notes-master-placeholder" ? { ...declaration, resolvedFrom: [...(declaration.resolvedFrom ?? []), ...matchDependencies] } : declaration]));
        byLevel.set(level, patch);
      }
      return patch;
    }, record(target, patch, prefix) { record(local, target, patch, prefix); } };
  };
  return { forDrawing, forLocal, forArtwork, forSlideArtwork, forNotes, report: () => ({ properties }) };
}
