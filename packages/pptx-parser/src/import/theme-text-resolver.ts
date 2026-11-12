import { attribute, children, invalidStructure } from "../presentation/structure-xml.js";
import type { structureContext } from "../presentation/structure-context.js";
import type { PptxInspection } from "../types/structure.js";
import type { PptxSourceLocation } from "../types/source.js";
import type { XmlElement } from "../xml/read-xml.js";
import { contentPaths, contentReader } from "./content-reader.js";
import { transformTextColor } from "./color-transforms.js";
import { readSourceFontMap, type SourceFontMap } from "./font-map.js";

export type ThemeTextResolution =
  | { status: "mapped"; value: string; opacity?: number; resolvedFrom: readonly PptxSourceLocation[]; program?: { rgb: string; modifiers: readonly XmlElement[] } }
  | { status: "unsupported"; reason: string; resolvedFrom?: readonly PptxSourceLocation[] };
export type ThemeStyleEntry =
  | { status: "mapped"; part: string; node: XmlElement; paths: WeakMap<XmlElement, string>; resolvedFrom: readonly PptxSourceLocation[] }
  | { status: "unsupported"; reason: string; resolvedFrom?: readonly PptxSourceLocation[] };
export type ThemeFontResolution = ThemeTextResolution & { fontMap?: SourceFontMap };
export interface ThemeTextScope {
  matrix: (kind: "fill" | "line", index: number) => Promise<ThemeStyleEntry>;
  color: (ownerPart: string, scheme: XmlElement, ownerPaths?: WeakMap<XmlElement, string>) => Promise<ThemeTextResolution>;
  font: (typeface: string, traceSource?: boolean, scriptMap?: boolean) => Promise<ThemeFontResolution>;
}

interface ThemeColor { value: string; source: PptxSourceLocation; modifiers: readonly XmlElement[]; modifierSources: readonly PptxSourceLocation[] }
interface ThemeFormats {
  part: string;
  node: XmlElement;
  paths: WeakMap<XmlElement, string>;
  lists: Map<string, Promise<readonly XmlElement[] | undefined>>;
}
interface ThemeFonts {
  source: PptxSourceLocation;
  part: string;
  paths: WeakMap<XmlElement, string>;
  values: Map<string, { value?: string; source: PptxSourceLocation }>;
  groups: Map<string, XmlElement>;
  checkedSupplemental: Set<string>;
  maps: Map<string, Promise<SourceFontMap>>;
}
interface Theme {
  // An absent scheme inherits as a unit; a present but unsupported slot blocks
  // an older scheme instead of falling through to a different palette/font.
  formats: ThemeFormats | undefined;
  colors: Map<string, ThemeColor> | undefined;
  fonts: ThemeFonts | undefined;
}
type ColorMapResult =
  | { type: "override"; entries: Map<string, { value: string; source: PptxSourceLocation }> }
  | { type: "master" | "unsupported" };

const defaults: Readonly<Record<string, string>> = {
  bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2",
  accent1: "accent1", accent2: "accent2", accent3: "accent3",
  accent4: "accent4", accent5: "accent5", accent6: "accent6",
  hlink: "hlink", folHlink: "folHlink",
};
const slots = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"] as const;

/**
 * Resolves theme-backed text values in the target slide's verified source scope.
 * Resolved RGB/font strings are materialized in Core; package/XML provenance
 * remains in the Parser report.
 */
export function themeTextResolver(
  context: Awaited<ReturnType<typeof structureContext>>,
  inspection: PptxInspection,
  checkpoint: () => Promise<void>,
  unmapped: (source: PptxSourceLocation, reason: string) => void,
) {
  const masters = new Map(inspection.slideMasters.map(master => [master.part, master]));
  const layouts = new Map(inspection.layouts.map(layout => [layout.part, layout]));
  const trees = new Map<string, { root: XmlElement; paths: WeakMap<XmlElement, string>; r: ReturnType<typeof contentReader> }>();
  const themes = new Map<string, Promise<Theme | undefined>>();
  const colorMaps = new Map<string, Promise<ColorMapResult | undefined>>();

  const tree = async (part: string, local: string, namespace = context.p) => {
    let value = trees.get(part);
    if (!value) {
      const type = local === "theme" || local === "themeOverride" ? local : undefined;
      const root = await context.root(part, local, type, namespace);
      value = { root, paths: await contentPaths(root, checkpoint), r: contentReader(part, context.p, context.a) };
      trees.set(part, value);
    } else if (value.root.local !== local || value.root.namespaceUri !== namespace) invalidStructure(`Unexpected source part root; expected ${local}.`, part);
    return value;
  };
  const source = (part: string, paths: WeakMap<XmlElement, string>, node: XmlElement, attr?: string): PptxSourceLocation => ({
    part, path: `${paths.get(node)!}${attr ? `/@${attr}` : ""}`,
  });
  const parseColor = (part: string, paths: WeakMap<XmlElement, string>, node: XmlElement): ThemeColor | undefined => {
    const choices = children(node);
    if (choices.length !== 1 || choices[0]!.namespaceUri !== context.a) return undefined;
    const choice = choices[0]!;
    if (choice.local === "srgbClr") {
      const value = attribute(choice, "val");
      if (!value || !/^[\da-f]{6}$/i.test(value)) return undefined;
      return { value: value.toUpperCase(), source: source(part, paths, choice, "val"), modifiers: children(choice), modifierSources: children(choice).map(modifier => source(part, paths, modifier)) };
    }
    if (choice.local === "sysClr") {
      const value = attribute(choice, "lastClr");
      if (!value || !/^[\da-f]{6}$/i.test(value)) return undefined;
      return { value: value.toUpperCase(), source: source(part, paths, choice, "lastClr"), modifiers: children(choice), modifierSources: children(choice).map(modifier => source(part, paths, modifier)) };
    }
    return undefined;
  };
  const theme = (part: string, rootName = "theme"): Promise<Theme | undefined> => {
    const key = `${part}\u0000${rootName}`;
    let pending = themes.get(key);
    if (!pending) {
      pending = (async () => {
        await checkpoint();
        const data = await tree(part, rootName, context.a), r = data.r;
        const elements = data.root.local === "themeOverride" ? data.root : r.one(data.root, context.a, "themeElements");
        let colors: Theme["colors"], fonts: Theme["fonts"], formats: Theme["formats"];
        if (elements) {
          const matrix = r.one(elements, context.a, "fmtScheme");
          if (matrix) formats = { part, node: matrix, paths: data.paths, lists: new Map() };
          const scheme = r.one(elements, context.a, "clrScheme");
          if (scheme) {
            colors = new Map();
            for (const slot of slots) {
              const node = r.one(scheme, context.a, slot);
              if (node) {
                const parsed = parseColor(part, data.paths, node);
                if (parsed) colors.set(slot, parsed);
              }
            }
          }
          const fontScheme = r.one(elements, context.a, "fontScheme");
          if (fontScheme) {
            fonts = { source: source(part, data.paths, fontScheme), part, paths: data.paths, values: new Map(), groups: new Map(), checkedSupplemental: new Set(), maps: new Map() };
            for (const [name, key] of [["majorFont", "major"], ["minorFont", "minor"]] as const) {
              const group = r.one(fontScheme, context.a, name);
              if (group) fonts.groups.set(key, group);
              for (const [element, suffix] of [["latin", "lt"], ["ea", "ea"], ["cs", "cs"]] as const) {
                await checkpoint();
                const font = group ? r.one(group, context.a, element) : undefined;
                const typeface = font ? attribute(font, "typeface") : undefined;
                if (font) fonts.values.set(`${key}-${suffix}`, { ...(typeface && !typeface.startsWith("+") ? { value: typeface } : {}), source: source(part, data.paths, font, "typeface") });
              }
            }
          }
        }
        return { colors, fonts, formats };
      })();
      themes.set(key, pending);
    }
    return pending;
  };
  const colorMap = (part: string, rootName: string): Promise<ColorMapResult | undefined> => {
    const key = `${part}\u0000${rootName}`;
    let pending = colorMaps.get(key);
    if (!pending) {
      pending = (async () => {
        const data = await tree(part, rootName), r = data.r;
        const isMaster = rootName === "sldMaster" || rootName === "notesMaster";
        const override = isMaster
          ? r.one(data.root, context.p, "clrMap")
          : r.one(data.root, context.p, "clrMapOvr");
        const mappingNode = isMaster
          ? override
          : override ? r.one(override, context.a, "overrideClrMapping") : undefined;
        if (!mappingNode) {
          if (override && !isMaster) {
            const masterMapping = r.one(override, context.a, "masterClrMapping");
            if (masterMapping) { r.known(masterMapping); return { type: "master" }; }
            unmapped(source(part, data.paths, override), "Unsupported color mapping override blocks theme-color projection.");
            return { type: "unsupported" };
          }
          return undefined;
        }
        const result = new Map<string, { value: string; source: PptxSourceLocation }>();
        for (const name of Object.keys(defaults)) {
          const value = attribute(mappingNode, name);
          if (value !== undefined) {
            r.attr(mappingNode, name);
            result.set(name, { value, source: source(part, data.paths, mappingNode, name) });
          }
        }
        return { type: "override", entries: result };
      })();
      colorMaps.set(key, pending);
    }
    return pending;
  };
  const scopeCache = new Map<string, ThemeTextScope>();
  const forScope = (masterPart: string | undefined, layoutPart: string | undefined, slidePart: string | undefined, notes: boolean): ThemeTextScope => {
    const masterRoot = notes ? "notesMaster" : "sldMaster", localRoot = notes ? "notes" : "sld";
    const key = `${notes}\u0000${masterPart ?? ""}\u0000${layoutPart ?? ""}\u0000${slidePart}`;
    const cached = scopeCache.get(key);
    if (cached) return cached;
    const componentThemes = new Map<keyof Theme, Promise<Theme | undefined>>();
    const selectedTheme = (component: keyof Theme): Promise<Theme | undefined> => {
      let pending = componentThemes.get(component);
      if (!pending) {
        pending = (async () => {
          for (const owner of new Set([slidePart, layoutPart, masterPart])) {
            if (!owner) continue;
            const overridePart = context.optionalTarget(owner, "themeOverride");
            if (overridePart) {
              const candidate = await theme(overridePart, "themeOverride");
              if (candidate?.[component] !== undefined) return candidate;
            }
          }
          const themePart = masterPart ? notes ? context.optionalTarget(masterPart, "theme") : masters.get(masterPart)?.themePart : undefined;
          return themePart ? theme(themePart) : undefined;
        })();
        componentThemes.set(component, pending);
      }
      return pending;
    };
    const scope: ThemeTextScope = {
      async matrix(kind, index): Promise<ThemeStyleEntry> {
        await checkpoint();
        const selected = (await selectedTheme("formats"))?.formats;
        if (!selected) return { status: "unsupported", reason: "The source scope has no verified theme format scheme." };
        const name = kind === "line" ? "lnStyleLst" : index > 1000 ? "bgFillStyleLst" : "fillStyleLst";
        let pending = selected.lists.get(name);
        if (!pending) {
          pending = (async () => {
            const list = contentReader(selected.part, context.p, context.a).one(selected.node, context.a, name);
            if (!list) return undefined;
            const entries: XmlElement[] = [];
            for (const node of children(list)) { await checkpoint(); entries.push(node); }
            return entries;
          })(); selected.lists.set(name, pending);
        }
        const entries = await pending, offset = kind === "fill" && index > 1000 ? index - 1001 : index - 1;
        const refs = [source(selected.part, selected.paths, selected.node)];
        if (!entries || offset < 0 || offset >= entries.length) return { status: "unsupported", reason: `Theme ${name} has no entry at index ${index}.`, resolvedFrom: refs };
        const node = entries[offset]!;
        refs.push(source(selected.part, selected.paths, node));
        if (node.namespaceUri !== context.a || !(kind === "line" ? node.local === "ln" : ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(node.local))) return { status: "unsupported", reason: "The selected theme style slot has an unsupported element; its position is not skipped.", resolvedFrom: refs };
        return { status: "mapped", part: selected.part, node, paths: selected.paths, resolvedFrom: refs };
      },
      async color(ownerPart: string, scheme: XmlElement, ownerPaths?: WeakMap<XmlElement, string>): Promise<ThemeTextResolution> {
      const parsedTheme = await selectedTheme("colors");
      if (!parsedTheme) return { status: "unsupported", reason: "The source master has no verified theme relationship or theme override." };
      const token = attribute(scheme, "val");
      if (!token) return { status: "unsupported", reason: "Theme color reference is missing its slot." };
      let slot = token;
      const ownerRoot = ownerPart === context.entry.mainPart ? "presentation" : ownerPart === masterPart ? masterRoot : !notes && layouts.has(ownerPart) ? "sldLayout" : localRoot;
      const paths = ownerPaths ?? (await tree(ownerPart, ownerRoot)).paths;
      const references = [source(ownerPart, paths, scheme, "val")];
      if (token in defaults) {
        let selected: ColorMapResult | undefined;
        const slideMap = slidePart ? await colorMap(slidePart, localRoot) : undefined;
        if (slideMap?.type === "override" || slideMap?.type === "unsupported") selected = slideMap;
        else if (slideMap?.type === "master") selected = masterPart ? await colorMap(masterPart, masterRoot) : undefined;
        else if (layoutPart) {
          const layoutMap = await colorMap(layoutPart, "sldLayout");
          if (layoutMap?.type === "override" || layoutMap?.type === "unsupported") selected = layoutMap;
          else if (layoutMap?.type === "master") selected = masterPart ? await colorMap(masterPart, masterRoot) : undefined;
        }
        if (!selected && masterPart) selected = await colorMap(masterPart, masterRoot);
        if (selected?.type === "unsupported") return { status: "unsupported", reason: "Unsupported color mapping override blocks theme-color projection.", resolvedFrom: references };
        const mapped = selected?.type === "override" ? selected.entries.get(token) : undefined;
        if (selected?.type === "override" && !mapped) return { status: "unsupported", reason: `Color mapping override omits the ${token} slot.`, resolvedFrom: references };
        if (mapped) {
          slot = mapped.value;
          references.push(...(selected?.type === "override" ? [mapped.source] : []));
        } else {
          slot = defaults[token]!;
        }
      }
      if (!slots.includes(slot as typeof slots[number])) return { status: "unsupported", reason: `Unknown theme color slot ${JSON.stringify(slot)}.` };
      const color = parsedTheme.colors?.get(slot);
      if (!color) return { status: "unsupported", reason: `Theme slot ${slot} has no supported direct RGB/system fallback value.` };
      const modifiers = children(scheme);
      references.push(color.source);
      for (const modifierSource of color.modifierSources) references.push(modifierSource);
      for (const modifier of modifiers) references.push(source(ownerPart, paths, modifier));
      const transformed = transformTextColor(color.value, [...color.modifiers, ...modifiers], context.a);
      if (transformed.status === "unsupported") return { ...transformed, resolvedFrom: references };
      return { ...transformed, resolvedFrom: references, program: { rgb: color.value, modifiers: [...color.modifiers, ...modifiers] } };
    },
    async font(typeface: string, traceSource = false, scriptMap = false): Promise<ThemeFontResolution> {
      const parsedTheme = await selectedTheme("fonts");
      if (!parsedTheme) return { status: "unsupported", reason: "The source master has no verified theme relationship." };
      const match = /^\+(mj|mn)-(lt|ea|cs)$/i.exec(typeface);
      if (!match) return { status: "unsupported", reason: "Only explicit major/minor Latin, East Asian or complex-script theme font slots are supported." };
      const key = match[1]!.toLowerCase() === "mj" ? "major" : "minor", fonts = parsedTheme.fonts;
      let fontMap: SourceFontMap | undefined;
      if (scriptMap && fonts) {
        const group = fonts.groups.get(key);
        if (group) {
          let pending = fonts.maps.get(key);
          if (!pending) {
            pending = (async () => {
              const r = contentReader(fonts.part, context.p, context.a), at = (n: XmlElement, attr?: string) => source(fonts.part, fonts.paths, n, attr);
              const result = await readSourceFontMap(group, { r, at }, checkpoint);
              for (const node of children(group)) if (node.namespaceUri === context.a && node.local === "font") await r.audit(node, at(node).path!, (path, reason) => unmapped({ part: fonts.part, path }, reason), checkpoint);
              return result;
            })(); fonts.maps.set(key, pending);
          }
          fontMap = await pending;
        }
      }
      if (traceSource && !scriptMap && fonts && !fonts.checkedSupplemental.has(key)) {
        fonts.checkedSupplemental.add(key);
        const group = fonts.groups.get(key);
        if (group) for (const node of children(group)) {
          await checkpoint();
          if (node.namespaceUri === context.a && node.local === "font") unmapped(source(fonts.part, fonts.paths, node), "Theme supplemental script-font selection remains source-only in the explicit font-slot profile.");
        }
      }
      const selected = fonts?.values.get(`${key}-${match[2]!.toLowerCase()}`);
      if (!selected?.value) return { status: "unsupported", reason: "The selected theme font slot has no explicit typeface; automatic language/script substitution is not inferred.", ...(traceSource && fonts ? { resolvedFrom: [selected?.source ?? fonts.source] } : {}), ...(fontMap ? { fontMap } : {}) };
      return { status: "mapped", value: selected.value, resolvedFrom: [selected.source], ...(fontMap ? { fontMap } : {}) };
    },
    };
    scopeCache.set(key, scope);
    return scope;
  };
  return { forMaster: (masterPart: string | undefined, layoutPart: string | undefined, slidePart: string) => forScope(masterPart, layoutPart, slidePart, false), forArtwork: (masterPart: string, layoutPart?: string) => forScope(masterPart, layoutPart, undefined, false), forNotes: (masterPart: string | undefined, notesPart: string) => forScope(masterPart, undefined, notesPart, true) };
}
