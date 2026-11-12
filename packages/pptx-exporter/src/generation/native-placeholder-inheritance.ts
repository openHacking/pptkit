import type { PresentationDocument, TextContentInput, TextStylePresetResetProperty } from "@pptkit/core";
import type { LayoutResult } from "@pptkit/layout";
import type { ExportWarning } from "../types/export.js";

import type { NativeInheritedRunProperty, NativeInheritedParagraphProperty, NativeInheritedFrameProperty, NativeInheritedProperty, NativeInheritedText, NativePlaceholderInheritance } from "../types/native-placeholder-inheritance.js";

const runSupported = new Set<TextStylePresetResetProperty>(["fontFamily", "color", "bold", "italic", "underline", "strike", "fontSize", "language", "rightToLeft"]);
const paragraphSupported = new Set<TextStylePresetResetProperty>(["paragraph.align", "paragraph.direction", "paragraph.indent", "paragraph.hanging", "paragraph.lineSpacing", "paragraph.spaceBefore", "paragraph.spaceAfter", "paragraph.bullet"]);
const frameSupported = new Set<TextStylePresetResetProperty>(["frame.textDirection", "frame.wrap", "frame.verticalAlign", "frame.autoFit", "frame.margin", "frame.margin.left", "frame.margin.right", "frame.margin.top", "frame.margin.bottom"]);
const supported = new Set([...runSupported, ...paragraphSupported, ...frameSupported]);
function localStyle(content: TextContentInput, paragraph: number, run: number) {
  return typeof content === "string" ? undefined : content[paragraph]?.runs[run]?.style;
}
/** Capture before any async asset loading; do not infer intent from equal effective values. */
export function nativePlaceholderInheritance(document: PresentationDocument, layout: LayoutResult, warnings: ExportWarning[], mode: "materialize" | "inherit"): NativePlaceholderInheritance {
  const layouts = new Map<string, Map<string, Set<NativeInheritedProperty>>>(), slides = new Map<string, Map<string, NativeInheritedText>>();
  for (const source of document.slides) {
    const slide = layout.slides.find(item => item.id === source.id);
    const template = slide && layout.layouts.find(item => item.id === slide.layoutId);
    if (!slide || !template) continue;
    for (const authored of source.elements) {
      const text = authored.type === "text" ? authored : authored.type === "shape" ? authored.text : undefined;
      if (!text?.textStylePresetReset?.length || !authored.placeholderKey) continue;
      const resetPath = authored.type === "shape" ? "text.textStylePresetReset" : "textStylePresetReset";
      if (mode === "materialize") { warnings.push({ code: "native-text-inheritance-materialized", message: "Reset text properties are exported as local effective values. Select nativePlaceholderText: inherit to opt into supported native inheritance.", slideId: source.id, elementId: authored.id, path: resetPath }); continue; }
      const placeholder = template.placeholders.find(item => item.key === authored.placeholderKey);
      const resolved = slide.elements.find(item => item.id === authored.id);
      const actual = resolved?.type === "text" ? resolved : resolved?.type === "shape" ? resolved.text : undefined;
      if (!placeholder || !actual || !resolved) continue;
      const preset = text.textStylePreset === undefined ? undefined : document.textStylePresets[text.textStylePreset]?.run;
      const masked = new Set(text.textStylePresetReset);
      let refused = text.textStylePresetReset.some(property => !supported.has(property));
      const properties = actual.content.map((paragraph, pi) => paragraph.runs.map((run, ri) => {
        const result: NativeInheritedRunProperty[] = [];
        const local = localStyle(text.content, pi, ri);
        for (const property of text.textStylePresetReset ?? []) {
          if (!runSupported.has(property) || local?.[property as NativeInheritedRunProperty] !== undefined) continue;
          const key = property as NativeInheritedRunProperty;
          const optionalFontBarrier = key === "fontFamily" && (["eastAsianFontFamily", "complexScriptFontFamily"] as const).some(slot =>
            local?.[slot] === null || (local?.[slot] === undefined && !masked.has(slot) && preset?.[slot] === null) || (run.style[slot] === undefined && placeholder.textStyle.run[slot] !== undefined));
          const explicitOpacity = local?.colorOpacity !== undefined || (!masked.has("colorOpacity") && preset?.colorOpacity !== undefined);
          if ((key === "rightToLeft" && placeholder.textStyle.run.rightToLeft === undefined) || JSON.stringify(run.style[key]) !== JSON.stringify(placeholder.textStyle.run[key]) || run.style.scriptFontFamilies || optionalFontBarrier || (key === "color" && (explicitOpacity || resolved.opacity !== 1 || (run.style.colorOpacity ?? 1) !== 1 || (placeholder.textStyle.run.colorOpacity ?? 1) !== 1))) { refused = true; continue; }
          result.push(key);
        }
        return result;
      }));
      const paragraphs = actual.content.map((paragraph, pi) => {
        const result: NativeInheritedParagraphProperty[] = [];
        const local = typeof text.content === "string" ? undefined : text.content[pi]?.style;
        for (const property of text.textStylePresetReset ?? []) {
          if (!paragraphSupported.has(property)) continue;
          const key = property.slice(10) as keyof typeof paragraph.style;
          if (local?.[key] !== undefined) continue;
          // A local/preset bullet can synthesize list indentation rather than inherit it.
          const presetBullet = text.textStylePreset === undefined ? undefined : document.textStylePresets[text.textStylePreset]?.paragraph?.bullet;
          const listIndentBarrier = (key === "indent" || key === "hanging") && ((local?.bullet !== undefined && local.bullet.type !== "none") || (!masked.has("paragraph.bullet") && presetBullet !== undefined && presetBullet.type !== "none"));
          if (listIndentBarrier || JSON.stringify(paragraph.style[key]) !== JSON.stringify(placeholder.textStyle.paragraph[key])) { refused = true; continue; }
          result.push(property as NativeInheritedParagraphProperty);
        }
        return result;
      });
      const frame: NativeInheritedFrameProperty[] = [];
      if (masked.has("frame.autoFit") && text.frame?.autoFit === undefined) {
        if (JSON.stringify(actual.frame.autoFit) !== JSON.stringify(placeholder.textStyle.frame.autoFit)) refused = true;
        else frame.push("frame.autoFit");
      }
      for (const key of ["wrap", "verticalAlign", "textDirection"] as const) {
        const property = `frame.${key}` as const;
        if (!masked.has(property) || text.frame?.[key] !== undefined) continue;
        if (actual.frame[key] !== placeholder.textStyle.frame[key]) { refused = true; continue; }
        frame.push(property);
      }
      for (const side of ["left", "right", "top", "bottom"] as const) {
        const property = `frame.margin.${side}` as const, margin = text.frame?.margin;
        if ((!masked.has("frame.margin") && !masked.has(property)) || typeof margin === "number" || margin?.[side] !== undefined) continue;
        if (actual.frame.margin[side] !== placeholder.textStyle.frame.margin[side]) { refused = true; continue; }
        frame.push(property);
      }
      if (refused) warnings.push({ code: "native-text-inheritance-materialized", message: "Some reset text properties cannot yet be represented as live PowerPoint placeholder inheritance; their effective values remain local.", slideId: source.id, elementId: authored.id, path: resetPath });
      if (!properties.some(paragraph => paragraph.some(run => run.length)) && !paragraphs.some(paragraph => paragraph.length) && !frame.length) continue;
      const perLayout = layouts.get(template.id) ?? new Map<string, Set<NativeInheritedProperty>>();
      const inherited = perLayout.get(placeholder.key) ?? new Set<NativeInheritedProperty>();
      for (const paragraph of properties) for (const run of paragraph) for (const property of run) inherited.add(property);
      for (const paragraph of paragraphs) for (const property of paragraph) inherited.add(property);
      if (inherited.size) { perLayout.set(placeholder.key, inherited); layouts.set(template.id, perLayout); }
      const entries = slides.get(source.id) ?? new Map(); entries.set(authored.id, { runs: properties, paragraphs, frame }); slides.set(source.id, entries);
    }
  }
  for (const layoutId of layouts.keys()) warnings.push({ code: "native-placeholder-metrics-may-change", message: "Native placeholder defaults can change text metrics for other shapes sharing this layout; verify the resulting presentation.", layoutId });
  return { layouts, slides };
}
