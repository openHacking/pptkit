import { TEXT_FONT_SCRIPT_CODES, type TextFrameStyleInput, type TextParagraphStyleInput, type TextRunStyleInput, type TextTabStop } from "@pptkit/core";
import { attribute, children, invalidStructure } from "../presentation/structure-xml.js";
import type { PptxSourceLocation } from "../types/source.js";
import type { PptxTextStyleLevel, PptxTextStyleProperty, PptxTextStyleValue } from "../types/inheritance.js";
import type { XmlElement } from "../xml/read-xml.js";
import type { ContentReader } from "./content-reader.js";
import type { ThemeTextResolution, ThemeFontResolution } from "./theme-text-resolver.js";
import { transformTextColor } from "./color-transforms.js";
import type { PptxImportOptions } from "../types/import.js";
import { readSourceFontMap, sourceFontField, type SourceFontMap } from "./font-map.js";

export interface TextStyleContext {
  r: ContentReader;
  level: PptxTextStyleLevel;
  at: (node: XmlElement, attribute?: string) => PptxSourceLocation;
  themeScope?: object;
  resolveColor?: (node: XmlElement) => Promise<ThemeTextResolution>;
  resolveFont?: (typeface: string) => Promise<ThemeFontResolution>;
  fontProfile?: PptxImportOptions["textFontProfile"];
  unmappedFontSelection?: (source: PptxSourceLocation, reason: string) => void;
}
export interface TextStyleDeclaration {
  source: PptxSourceLocation;
  level: PptxTextStyleLevel;
  status: "mapped" | "unsupported";
  value?: PptxTextStyleValue;
  reason?: string;
  resolvedFrom?: readonly PptxSourceLocation[];
  /** Private until the effective native slot declaration has won the cascade. */
  fontMap?: SourceFontMap;
}
export type TextStylePatch = Map<PptxTextStyleProperty, TextStyleDeclaration>;
export const textFontFields = ["run.fontFamily", "run.eastAsianFontFamily", "run.complexScriptFontFamily"] as const;

async function fontValue(value: string | undefined, c: TextStyleContext): Promise<ThemeFontResolution> {
  if (!value || !value.trim() || /[\u0000-\u001F\u007F]/u.test(value)) return { status: "unsupported", reason: "Explicit source font has no supported nonempty typeface; automatic substitution is not inferred." };
  if (!value.startsWith("+")) return { status: "mapped", value, resolvedFrom: [] };
  const resolved = c.resolveFont ? await c.resolveFont(value) : { status: "unsupported" as const, reason: "Theme font resolution is unavailable in this source scope." };
  if (resolved.status === "mapped" && (!resolved.value.trim() || /[\u0000-\u001F\u007F]/u.test(resolved.value))) return { status: "unsupported", reason: "Selected theme font has no supported nonempty typeface.", resolvedFrom: resolved.resolvedFrom, ...(resolved.fontMap ? { fontMap: resolved.fontMap } : {}) };
  return resolved;
}

/** Table font-reference selection, with independent explicit script slots when opted in. */
export async function referenceTextFonts(node: XmlElement, c: TextStyleContext): Promise<TextStylePatch> {
  const patch: TextStylePatch = new Map(), w = writer(c, patch), index = c.r.attr(node, "idx");
  for (const [i, key] of (c.fontProfile ? textFontFields : [textFontFields[0]]).entries()) {
    const token = index === "major" || index === "minor" ? `+${index === "major" ? "mj" : "mn"}-${["lt", "ea", "cs"][i]}` : undefined;
    const resolved = token ? await fontValue(token, c) : { status: "unsupported" as const, reason: "Unsupported table font reference blocks an older font." };
    if (resolved.status === "mapped") w.set(key, resolved.value, node, "idx", resolved.resolvedFrom);
    else w.block(key, resolved.reason, node, "idx", resolved.resolvedFrom);
    if (resolved.fontMap) patch.get(key)!.fontMap = resolved.fontMap;
  }
  return patch;
}

/** Higher explicit declarations, including unsupported ones, replace one field. */
export function mergeTextStyles(...layers: readonly TextStylePatch[]): TextStylePatch {
  const result: TextStylePatch = new Map();
  for (const layer of layers) for (const [key, value] of layer) result.set(key, value);
  return result;
}
function writer(c: TextStyleContext, patch: TextStylePatch) {
  return {
    set(key: PptxTextStyleProperty, value: PptxTextStyleValue, node: XmlElement, attr?: string, resolvedFrom?: readonly PptxSourceLocation[]) {
      patch.set(key, { source: c.at(node, attr), level: c.level, status: "mapped", value, ...(resolvedFrom?.length ? { resolvedFrom } : {}) });
    },
    block(key: PptxTextStyleProperty, reason: string, node: XmlElement, attr?: string, resolvedFrom?: readonly PptxSourceLocation[]) {
      c.r.note(node, reason);
      patch.set(key, { source: c.at(node, attr), level: c.level, status: "unsupported", reason, ...(resolvedFrom?.length ? { resolvedFrom } : {}) });
    },
  };
}

export function frameTextStyles(node: XmlElement | undefined, c: TextStyleContext): TextStylePatch {
  const patch: TextStylePatch = new Map();
  if (!node) return patch;
  const { r } = c, w = writer(c, patch); r.known(node);
  for (const [attr, side] of [["lIns", "left"], ["rIns", "right"], ["tIns", "top"], ["bIns", "bottom"]] as const) {
    if (r.attr(node, attr) === undefined) continue;
    const value = r.number(node, attr, -2147483648, 2147483647);
    if (value >= 0) w.set(`frame.margin.${side}`, value / 12700, node, attr);
    else w.block(`frame.margin.${side}`, "Negative text insets have no current Core margin mapping.", node, attr);
  }
  for (const [attr, key, mapping] of [
    ["vert", "frame.textDirection", { horz: "horizontal", vert: "rotate90", vert270: "rotate270", eaVert: "eastAsianVertical", mongolianVert: "mongolianVertical", wordArtVert: "stackedVertical", wordArtVertRtl: "stackedVerticalRtl" }],
    ["anchor", "frame.verticalAlign", { t: "top", ctr: "middle", b: "bottom" }],
    ["wrap", "frame.wrap", { none: false, square: true }],
  ] as const) {
    const value = attribute(node, attr);
    if (value === undefined) continue;
    if (Object.hasOwn(mapping, value)) { r.attr(node, attr); w.set(key, mapping[value as keyof typeof mapping]!, node, attr); }
    else w.block(key, `Unsupported text ${attr} blocks an older inherited value.`, node, attr);
  }
  const fits = children(node).filter(n => n.namespaceUri === r.a && ["noAutofit", "normAutofit", "spAutoFit"].includes(n.local));
  if (fits.length > 1) invalidStructure("Conflicting text autofit choices.", r.part);
  const fit = fits[0];
  if (fit) {
    r.known(fit);
    w.set("frame.autoFit", fit.local === "noAutofit" ? { mode: "none" } : fit.local === "spAutoFit" ? { mode: "resize" } : { mode: "shrink", fontScale: r.number(fit, "fontScale", 1, 100000, 100000) / 100000 }, fit);
  }
  return patch;
}

export async function runTextStyles(node: XmlElement | undefined, c: TextStyleContext): Promise<TextStylePatch> {
  const patch: TextStylePatch = new Map();
  if (!node) return patch;
  const { r } = c, w = writer(c, patch); r.known(node);
  const rtl = r.one(node, r.a, "rtl");
  if (rtl) { r.known(rtl); w.set("run.rightToLeft", r.bool(rtl, "val", true), rtl, r.attr(rtl, "val") === undefined ? undefined : "val"); }
  if (r.attr(node, "sz") !== undefined) w.set("run.fontSize", r.number(node, "sz", 100, 400000) / 100, node, "sz");
  for (const [attr, key] of [["b", "run.bold"], ["i", "run.italic"]] as const) {
    if (r.attr(node, attr) !== undefined) w.set(key, r.bool(node, attr), node, attr);
  }
  for (const [attr, key, mapping] of [
    ["u", "run.underline", { none: false, sng: true }],
    ["strike", "run.strike", { noStrike: false, sngStrike: true }],
  ] as const) {
    const value = attribute(node, attr);
    if (value === undefined) continue;
    if (Object.hasOwn(mapping, value)) { r.attr(node, attr); w.set(key, mapping[value as keyof typeof mapping]!, node, attr); }
    else w.block(key, `Unsupported ${attr} blocks an older inherited character value.`, node, attr);
  }
  const language = r.attr(node, "lang");
  if (language !== undefined) w.set("run.language", language, node, "lang");
  const fills = children(node).filter(n => n.namespaceUri === r.a && ["noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"].includes(n.local));
  if (fills.length > 1) invalidStructure("Conflicting text fill choices.", r.part);
  if (fills.length) {
    const fill = fills[0]!;
    const choice = children(fill);
    const color = choice.length === 1 ? choice[0] : undefined;
    const setColor = (value: string, opacity: number, resolvedFrom?: readonly PptxSourceLocation[]) => {
      r.known(fill); r.known(color!);
      for (const modifier of children(color!)) { r.known(modifier); r.attr(modifier, "val"); }
      w.set("run.color", value, color!, color!.local === "sysClr" ? "lastClr" : "val", resolvedFrom);
      // A new fill replaces its alpha as well as RGB. Explicit opaque alpha
      // must cancel a translucent lower-level color instead of inheriting it.
      w.set("run.colorOpacity", opacity, color!, undefined, resolvedFrom);
    };
    const blockColor = (reason: string, resolvedFrom?: readonly PptxSourceLocation[]) => {
      r.known(fill);
      w.block("run.color", reason, color ?? fill, undefined, resolvedFrom);
      patch.set("run.colorOpacity", { source: c.at(color ?? fill), level: c.level, status: "unsupported", reason, ...(resolvedFrom?.length ? { resolvedFrom } : {}) });
    };
    if (fill.local === "solidFill" && color?.namespaceUri === r.a && color.local === "schemeClr" && c.resolveColor) {
      r.attr(color, "val");
      const resolved = await c.resolveColor(color);
      if (resolved.status === "mapped") {
        setColor(resolved.value, resolved.opacity ?? 1, resolved.resolvedFrom);
      }
      else blockColor(resolved.reason, resolved.resolvedFrom);
    } else if (fill.local === "solidFill" && color?.namespaceUri === r.a && ["srgbClr", "sysClr"].includes(color.local)) {
      const name = color.local === "sysClr" ? "lastClr" : "val", rgb = r.attr(color, name);
      if (rgb === undefined && color.local === "sysClr") blockColor("System text color has no portable lastClr fallback.");
      else if (!rgb || !/^[\da-f]{6}$/i.test(rgb)) invalidStructure("Invalid direct text RGB color.", r.part);
      else {
        if (color.local === "sysClr") r.attr(color, "val");
        const result = transformTextColor(rgb, children(color), r.a);
        if (result.status === "mapped") setColor(result.value, result.opacity, children(color).map(modifier => c.at(modifier)));
        else blockColor(result.reason, children(color).map(modifier => c.at(modifier)));
      }
    } else {
      blockColor("Run paint/color is outside the supported text-color profile; the explicit declaration blocks inherited color and alpha.");
    }
  }
  if (c.fontProfile) {
    for (const [i, name] of ["latin", "ea", "cs"].entries()) {
      const font = r.one(node, r.a, name); if (!font) continue;
      const token = r.attr(font, "typeface"), resolved = await fontValue(token, c), key = textFontFields[i]!;
      if (resolved.status === "mapped") w.set(key, resolved.value, font, "typeface", resolved.resolvedFrom);
      else w.block(key, resolved.reason, font, "typeface", resolved.resolvedFrom);
      // Cross-slot theme tokens explicitly name another concrete slot. They do
      // not opt this destination into that collection's script substitutions.
      if (resolved.fontMap && token?.toLowerCase().endsWith(["-lt", "-ea", "-cs"][i]!)) patch.get(key)!.fontMap = resolved.fontMap;
    }
    return patch;
  }
  const latin = r.one(node, r.a, "latin"), ea = r.one(node, r.a, "ea"), cs = r.one(node, r.a, "cs");
  let latinFamily: string | undefined;
  if (latin) {
    const font = r.attr(latin, "typeface");
    if (font && font.startsWith("+") && c.resolveFont) {
      const resolved = await c.resolveFont(font);
      if (resolved.status === "mapped") { latinFamily = resolved.value; w.set("run.fontFamily", resolved.value, latin, "typeface", resolved.resolvedFrom); }
      else w.block("run.fontFamily", resolved.reason, latin, "typeface", resolved.resolvedFrom);
    } else if (font && !font.startsWith("+")) {
      latinFamily = font;
      w.set("run.fontFamily", font, latin, "typeface");
    } else w.block("run.fontFamily", "Theme font resolution is pending or the font is empty; the explicit declaration blocks an older inherited Latin font.", latin, "typeface");
  }
  for (const alternative of [ea, cs]) {
    if (!alternative) continue;
    const font = r.attr(alternative, "typeface");
    if (font === undefined) {
      w.block("run.fontFamily", "An explicit script font lacks its required typeface and blocks older inherited font values.", alternative, "typeface");
      continue;
    }
    const resolved = font.startsWith("+") && c.resolveFont ? await c.resolveFont(font) : undefined;
    const family = resolved?.status === "mapped" ? resolved.value : font.startsWith("+") ? undefined : font;
    if (!family || !latinFamily || family !== latinFamily) {
      const reason = "Script-specific font cannot be represented by Core's single font-family value and blocks older inherited font values.";
      w.block("run.fontFamily", resolved?.status === "unsupported" ? resolved.reason : reason, alternative, "typeface", resolved?.resolvedFrom);
    } else {
      const selected = patch.get("run.fontFamily");
      if (selected?.status === "mapped") selected.resolvedFrom = [...(selected.resolvedFrom ?? []), c.at(alternative, "typeface"), ...(resolved?.resolvedFrom ?? [])];
    }
  }
  return patch;
}

/** A source table font collection may collapse only when its explicit families agree. */
export async function collectionTextStyles(node: XmlElement, c: TextStyleContext, checkpoint: () => Promise<void>): Promise<TextStylePatch> {
  const patch = await runTextStyles(node, c), w = writer(c, patch), scripts = new Set<string>();
  const fields = c.fontProfile ? textFontFields : [textFontFields[0]];
  if (c.fontProfile) for (const key of fields) if (!patch.has(key)) w.block(key, "An explicit source font collection omits this required slot; older collections are not substituted.", node);
  if (c.fontProfile === "script-families-v2" || c.fontProfile === "script-families-v3") {
    const map = await readSourceFontMap(node, c, checkpoint);
    if (map.size) for (const key of fields) patch.get(key)!.fontMap = map;
    return patch;
  }
  for (const font of children(node)) {
    await checkpoint();
    if (font.namespaceUri !== c.r.a || font.local !== "font") continue;
    c.r.known(font);
    const script = c.r.attr(font, "script"), family = c.r.attr(font, "typeface");
    if (!script || !family) { for (const key of fields) w.block(key, c.fontProfile ? "Incomplete supplemental font blocks font collection projection." : "Incomplete supplemental font blocks the single-family projection.", font); continue; }
    if (scripts.has(script)) invalidStructure("Duplicate supplemental font script identities.", c.r.part);
    scripts.add(script);
    const selected = fields.map(key => patch.get(key));
    if (selected.some(value => value?.status !== "mapped" || value.value !== family) || family.startsWith("+")) {
      for (const key of fields) w.block(key, c.fontProfile ? "Supplemental script selection is outside the explicit font-slot profile." : "Supplemental script fonts differ or cannot resolve to a single Core font family.", font, "typeface");
    } else for (const value of selected) value!.resolvedFrom = [...(value!.resolvedFrom ?? []), c.at(font, "typeface")];
  }
  if (!patch.has("run.fontFamily")) w.block("run.fontFamily", "The explicit font collection has no supported Latin family.", node);
  return patch;
}

export async function paragraphTextStyles(node: XmlElement | undefined, c: TextStyleContext): Promise<TextStylePatch> {
  if (!node) return new Map();
  const { r } = c; r.known(node);
  const patch = await runTextStyles(r.one(node, r.a, "defRPr"), c), w = writer(c, patch);
  if (r.attr(node, "rtl") !== undefined) w.set("paragraph.direction", r.bool(node, "rtl") ? "rtl" : "ltr", node, "rtl");
  const align = attribute(node, "algn");
  if (align !== undefined) {
    const value = r.enumValue(node, "algn", { l: "left", ctr: "center", r: "right", just: "justify" } as const);
    if (value) w.set("paragraph.align", value, node, "algn");
    else w.block("paragraph.align", "Unsupported paragraph alignment blocks inherited alignment.", node, "algn");
  }
  if (r.attr(node, "marL") !== undefined) w.set("paragraph.indent", r.number(node, "marL", 0, 51206400) / 12700, node, "marL");
  if (r.attr(node, "indent") !== undefined) {
    const indent = r.number(node, "indent", -51206400, 51206400);
    if (indent <= 0) w.set("paragraph.hanging", indent === 0 ? 0 : -indent / 12700, node, "indent");
    else w.block("paragraph.hanging", "Positive first-line indentation has no current Core hanging-indent mapping.", node, "indent");
  }
  for (const [name, field] of [["spcBef", "paragraph.spaceBefore"], ["spcAft", "paragraph.spaceAfter"], ["lnSpc", "paragraph.lineSpacing"]] as const) {
    const spacing = r.one(node, r.a, name);
    if (!spacing) continue;
    r.known(spacing);
    const points = r.one(spacing, r.a, "spcPts"), percent = r.one(spacing, r.a, "spcPct");
    if (points && percent) invalidStructure("Conflicting paragraph spacing choices.", r.part);
    if (name === "lnSpc" && percent) w.set(field, r.number(percent, "val", 1, 13200000) / 100000, percent, "val");
    else if (name !== "lnSpc" && points) w.set(field, r.number(points, "val", 0, 158400) / 100, points, "val");
    else w.block(field, "This paragraph spacing encoding cannot map to the current Core unit; it blocks inherited spacing.", spacing);
  }
  const bullets = children(node).filter(n => n.namespaceUri === r.a && ["buNone", "buChar", "buAutoNum", "buBlip"].includes(n.local));
  if (bullets.length > 1) invalidStructure("Conflicting paragraph bullet choices.", r.part);
  const bullet = bullets[0];
  if (bullet?.local === "buNone") { r.known(bullet); w.set("paragraph.bullet", { type: "none" }, bullet); }
  else if (bullet?.local === "buChar") {
    const character = r.attr(bullet, "char");
    if (character) w.set("paragraph.bullet", { type: "bullet", character }, bullet);
    else w.block("paragraph.bullet", "An empty bullet character blocks inherited bullets.", bullet);
  } else if (bullet?.local === "buAutoNum") {
    const numbering = r.enumValue(bullet, "type", { arabicPeriod: "arabicPeriod", arabicParenR: "arabicParen", alphaLcPeriod: "alphaLowerPeriod", alphaUcPeriod: "alphaUpperPeriod" } as const);
    if (numbering) w.set("paragraph.bullet", { type: "number", style: numbering, startAt: r.number(bullet, "startAt", 1, 32767, 1) }, bullet);
    else w.block("paragraph.bullet", "Unsupported automatic numbering blocks inherited bullets.", bullet);
  } else if (bullet) w.block("paragraph.bullet", "Picture bullets block inherited text bullets.", bullet);
  const tabList = r.one(node, r.a, "tabLst");
  if (tabList) {
    r.known(tabList);
    const nodes = children(tabList), stops: TextTabStop[] = [], sources: PptxSourceLocation[] = [];
    let reason: string | undefined, previous = -1;
    if (nodes.length > 256) reason = "Tab lists above 256 stops are not projected; the declaration blocks inherited stops.";
    else for (const tab of nodes) {
      if (tab.namespaceUri !== r.a || tab.local !== "tab") { reason = "Unrecognized tab-list content blocks inherited stops."; break; }
      r.known(tab);
      const rawPosition = r.attr(tab, "pos"), rawAlignment = r.attr(tab, "algn");
      if (rawPosition === undefined) { reason = "Application-dependent tab position or alignment blocks inherited stops."; break; }
      const position = r.number(tab, "pos", -2147483648, 2147483647);
      if (rawAlignment === undefined) { reason = "Application-dependent tab position or alignment blocks inherited stops."; break; }
      const mapping = { l: "left", ctr: "center", r: "right", dec: "decimal" } as const;
      const alignment = Object.hasOwn(mapping, rawAlignment) ? mapping[rawAlignment as keyof typeof mapping] : undefined;
      if (!alignment || position < 0 || position <= previous) { reason = "Unsupported alignment or nonincreasing tab positions block inherited stops."; break; }
      stops.push({ position: position / 12700, align: alignment }); previous = position; sources.push(c.at(tab));
    }
    if (reason) w.block("paragraph.tabStops", reason, tabList);
    else w.set("paragraph.tabStops", stops, tabList, undefined, sources);
  }
  return patch;
}

export async function listTextStyles(node: XmlElement | undefined, c: TextStyleContext): Promise<{ defaults: TextStylePatch; levels: TextStylePatch[] }> {
  if (node) c.r.known(node);
  const defaults = await paragraphTextStyles(node ? c.r.one(node, c.r.a, "defPPr") : undefined, c), levels: TextStylePatch[] = [];
  for (let index = 0; index < 9; index++) levels.push(await paragraphTextStyles(node ? c.r.one(node, c.r.a, `lvl${index + 1}pPr`) : undefined, c));
  return {
    defaults, levels,
  };
}

function values(patch: TextStylePatch, prefix: string): Record<string, PptxTextStyleValue> {
  return Object.fromEntries([...patch].filter(([key, declaration]) => key.startsWith(prefix) && declaration.status === "mapped").map(([key, declaration]) => [key.slice(prefix.length), declaration.value!]));
}
export function frameTextValues(patch: TextStylePatch): TextFrameStyleInput {
  const fields = values(patch, "frame."), margin: Record<string, number> = {};
  for (const side of ["left", "right", "top", "bottom"]) {
    const key = `margin.${side}`;
    if (typeof fields[key] === "number") margin[side] = fields[key] as number;
    delete fields[key];
  }
  return { ...fields, margin } as TextFrameStyleInput;
}
export function paragraphTextValues(patch: TextStylePatch): TextParagraphStyleInput {
  const result = values(patch, "paragraph.");
  const direction = patch.get("paragraph.direction");
  // Keep the LTR presentation default implicit only after resolving the
  // cascade: its level-specific false value may override an earlier RTL rule.
  if (direction?.level === "presentation-default" && direction.status === "mapped" && direction.value === "ltr") delete result.direction;
  return result as TextParagraphStyleInput;
}
/** Resolve source rules only after all native slot inheritance/overrides are known. */
export function effectiveTextFonts(patch: TextStylePatch, context?: Pick<TextStyleContext, "fontProfile" | "unmappedFontSelection">): TextStylePatch {
  if (![...patch.values()].some(value => value.fontMap)) return patch;
  const result: TextStylePatch = new Map([...patch].map(([key, { fontMap: _map, ...value }]) => [key, value]));
  for (const code of TEXT_FONT_SCRIPT_CODES) {
    const field = sourceFontField(code), declared = patch.get(field), selected = declared ?? patch.get("run.fontFamily"), entry = selected?.fontMap?.get(code);
    if (!entry) continue;
    const resolvedFrom = [selected!.source, ...(selected!.resolvedFrom ?? []), ...(entry.resolvedFrom ?? [])];
    if (context?.fontProfile === "script-families-v3" && !declared) {
      const { value: _value, ...source } = entry;
      const reason = `Source supplemental ${code} has no effective ${field} declaration; the Latin collection does not supply an absent native slot in script-families-v3.`;
      result.set(`run.scriptFontFamilies.${code}`, { ...source, level: selected!.level, status: "unsupported", reason, resolvedFrom });
      context.unmappedFontSelection?.(entry.source, reason);
    } else result.set(`run.scriptFontFamilies.${code}`, { ...entry, level: selected!.level, resolvedFrom });
  }
  return result;
}
export function runTextValues(patch: TextStylePatch): TextRunStyleInput {
  const fields = values(patch, "run."), scriptFontFamilies: Record<string, PptxTextStyleValue> = {};
  for (const key of Object.keys(fields)) if (key.startsWith("scriptFontFamilies.")) { scriptFontFamilies[key.slice(19)] = fields[key]!; delete fields[key]; }
  return { ...fields, ...(Object.keys(scriptFontFamilies).length ? { scriptFontFamilies } : {}) } as TextRunStyleInput;
}
