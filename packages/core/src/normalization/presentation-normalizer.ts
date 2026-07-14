import {
  DEFAULT_BACKGROUND,
  DEFAULT_LAYOUT_ID,
  DEFAULT_PRESENTATION_SIZE,
  DEFAULT_TEXT_FRAME_STYLE,
  DEFAULT_TEXT_PARAGRAPH_STYLE,
  DEFAULT_TEXT_RUN_STYLE,
  DEFAULT_THEME,
} from "../constants/presentation.js";
import type { NormalizedAsset } from "../types/asset.js";
import type {
  NormalizedElement,
  NormalizedElementBase,
  NormalizedPlaceholderDefinition,
  NormalizedTextParagraph,
  PlaceholderDefinitionInput,
  PresentationElement,
  TableCellInput,
  TextParagraphInput,
  TextContentInput,
} from "../types/element.js";
import type { Box } from "../types/geometry.js";
import type {
  NormalizedPresentation,
  NormalizedPresentationMetadata,
  NormalizedSlideLayout,
  PresentationDocument,
  PresentationSlide,
  SlideLayoutDefinition,
} from "../types/presentation.js";
import type { NormalizedPresentationTheme } from "../types/theme.js";
import type { NormalizedTextFrameStyle, NormalizedTextParagraphStyle, NormalizedTextRunStyle, TextStylePresetInput } from "../types/style.js";
import { deepClone } from "../utils/clone.js";
import { normalizeSize } from "../validation/geometry.js";
import { estimateTextHeight } from "./measure-text.js";
import {
  normalizePaint,
  normalizeShapeStyle,
  normalizeStroke,
  normalizeTableCellStyle,
  normalizeTextFrame,
  normalizeTextParagraphStyle,
  normalizeTextRunStyle,
} from "./normalize-styles.js";

function normalizeTheme(document: PresentationDocument): NormalizedPresentationTheme {
  return {
    name: document.theme.name ?? DEFAULT_THEME.name,
    colors: { ...DEFAULT_THEME.colors, ...(document.theme.colors ?? {}) },
    fonts: { ...DEFAULT_THEME.fonts, ...(document.theme.fonts ?? {}) },
  };
}

function normalizeMetadata(document: PresentationDocument): NormalizedPresentationMetadata {
  return {
    title: document.metadata.title ?? "PPTKit Presentation",
    author: document.metadata.author ?? "PPTKit",
    ...(document.metadata.company !== undefined ? { company: document.metadata.company } : {}),
    ...(document.metadata.subject !== undefined ? { subject: document.metadata.subject } : {}),
    ...(document.metadata.description !== undefined ? { description: document.metadata.description } : {}),
    language: document.metadata.language ?? "en-US",
    keywords: [...(document.metadata.keywords ?? [])],
    revision: document.metadata.revision ?? 1,
  };
}

export function normalizeTextContent(
  content: TextContentInput,
  paragraphFallback = DEFAULT_TEXT_PARAGRAPH_STYLE,
  runFallback = DEFAULT_TEXT_RUN_STYLE,
): NormalizedTextParagraph[] {
  const paragraphs: TextParagraphInput[] = typeof content === "string"
    ? content.split(/\r\n|\r|\n/).map((text) => ({ runs: [{ text }] }))
    : content;
  return paragraphs.map((paragraph) => ({
    style: normalizeTextParagraphStyle(paragraph.style, paragraphFallback),
    runs: paragraph.runs.map((run) => ({
      text: run.text,
      style: normalizeTextRunStyle(run.style, runFallback),
      ...(run.action !== undefined ? { action: deepClone(run.action) } : {}),
    })),
  }));
}

function normalizeTextFallbacks(
  preset: TextStylePresetInput | undefined,
  placeholder?: NormalizedPlaceholderDefinition,
): { frame: NormalizedTextFrameStyle; paragraph: NormalizedTextParagraphStyle; run: NormalizedTextRunStyle } {
  return {
    frame: normalizeTextFrame(preset?.frame, placeholder?.textStyle.frame ?? DEFAULT_TEXT_FRAME_STYLE),
    paragraph: normalizeTextParagraphStyle(preset?.paragraph, placeholder?.textStyle.paragraph ?? DEFAULT_TEXT_PARAGRAPH_STYLE),
    run: normalizeTextRunStyle(preset?.run, placeholder?.textStyle.run ?? DEFAULT_TEXT_RUN_STYLE),
  };
}

function normalizePlaceholder(input: PlaceholderDefinitionInput, presets: Readonly<Record<string, TextStylePresetInput>>): NormalizedPlaceholderDefinition {
  const preset = input.textStylePreset === undefined ? undefined : presets[input.textStylePreset];
  return {
    key: input.key,
    kind: input.kind,
    box: deepClone(input.box),
    textStyle: {
      frame: normalizeTextFrame(input.textStyle?.frame, normalizeTextFrame(preset?.frame)),
      paragraph: normalizeTextParagraphStyle(input.textStyle?.paragraph, normalizeTextParagraphStyle(preset?.paragraph)),
      run: normalizeTextRunStyle(input.textStyle?.run, normalizeTextRunStyle(preset?.run)),
    },
  };
}

function normalizeCellStyle(cell: TableCellInput, presets: Readonly<Record<string, TextStylePresetInput>>) {
  const preset = presets[cell.textStylePreset ?? ""];
  return normalizeTableCellStyle({
    ...cell.style,
    ...(cell.style?.margin === undefined && preset?.frame?.margin !== undefined ? { margin: preset.frame.margin } : {}),
    ...(cell.style?.verticalAlign === undefined && preset?.frame?.verticalAlign !== undefined ? { verticalAlign: preset.frame.verticalAlign } : {}),
  });
}

function boxFromElement(element: PresentationElement, placeholder?: NormalizedPlaceholderDefinition): Box {
  if (element.box !== undefined) {
    return {
      ...deepClone(element.box),
      ...(element.type === "text" && element.box.height === undefined ? { height: 0 } : {}),
    };
  }
  if (placeholder !== undefined) return deepClone(placeholder.box);
  if (element.type === "connector" && !("elementId" in element.start) && !("elementId" in element.end)) {
    const xs = [element.start.x, element.end.x, ...(element.route ?? []).map((point) => point.x)];
    const ys = [element.start.y, element.end.y, ...(element.route ?? []).map((point) => point.y)];
    const left = Math.min(...xs);
    const top = Math.min(...ys);
    return { x: left, y: top, width: Math.max(...xs) - left, height: Math.max(...ys) - top };
  }
  return { x: 0, y: 0, width: 0, height: 0 };
}

function normalizedBase(element: PresentationElement, placeholder?: NormalizedPlaceholderDefinition): NormalizedElementBase {
  return {
    id: element.id,
    name: element.name ?? element.id,
    box: boxFromElement(element, placeholder),
    transform: {
      rotation: element.transform?.rotation ?? 0,
      flipH: element.transform?.flipH ?? false,
      flipV: element.transform?.flipV ?? false,
    },
    opacity: element.opacity ?? 1,
    hidden: element.hidden ?? false,
    accessibility: {
      ...(element.accessibility?.description !== undefined ? { description: element.accessibility.description } : {}),
      decorative: element.accessibility?.decorative ?? false,
    },
    ...(element.action !== undefined ? { action: deepClone(element.action) } : {}),
    ...(element.placeholderKey !== undefined ? { placeholderKey: element.placeholderKey } : {}),
  };
}

function normalizeElement(
  element: PresentationElement,
  placeholders: ReadonlyMap<string, NormalizedPlaceholderDefinition> = new Map(),
  presets: Readonly<Record<string, TextStylePresetInput>> = {},
): NormalizedElement {
  const placeholder = element.placeholderKey === undefined ? undefined : placeholders.get(element.placeholderKey);
  const base = normalizedBase(element, placeholder);
  if (element.type === "text") {
    const fallbacks = normalizeTextFallbacks(element.textStylePreset === undefined ? undefined : presets[element.textStylePreset], placeholder);
    const paragraphFallback = fallbacks.paragraph;
    const runFallback = fallbacks.run;
    const content = normalizeTextContent(element.content, paragraphFallback, runFallback);
    const frame = normalizeTextFrame(element.frame, fallbacks.frame);
    const box = element.box?.height === undefined && element.box !== undefined && placeholder === undefined
      ? { ...base.box, height: estimateTextHeight(base.box.width, content, frame) }
      : base.box;
    return {
      ...base,
      box,
      type: "text",
      content,
      plainText: content.map((paragraph) => paragraph.runs.map((run) => run.text).join("")).join("\n"),
      frame,
    };
  }
  if (element.type === "image") {
    return {
      ...base,
      type: "image",
      assetId: element.assetId,
      fit: element.fit ?? "stretch",
      crop: { left: 0, top: 0, right: 0, bottom: 0, ...(element.crop ?? {}) },
    };
  }
  if (element.type === "shape") {
    if (element.text === undefined) return { ...base, type: "shape", shape: element.shape, style: normalizeShapeStyle(element.style) };
    const fallbacks = normalizeTextFallbacks(element.text.textStylePreset === undefined ? undefined : presets[element.text.textStylePreset], placeholder);
    const content = normalizeTextContent(element.text.content, fallbacks.paragraph, fallbacks.run);
    return {
      ...base,
      type: "shape",
      shape: element.shape,
      style: normalizeShapeStyle(element.style),
      text: {
        content,
        plainText: content.map((paragraph) => paragraph.runs.map((run) => run.text).join("")).join("\n"),
        frame: normalizeTextFrame(element.text.frame, fallbacks.frame),
      },
    };
  }
  if (element.type === "connector") {
    return {
      ...base,
      type: "connector",
      start: deepClone(element.start),
      end: deepClone(element.end),
      route: deepClone(element.route ?? []),
      style: normalizeStroke(element.style),
    };
  }
  if (element.type === "group") {
    return {
      ...base,
      type: "group",
      coordinateSize: deepClone(element.coordinateSize),
      children: element.children.map((child) => normalizeElement(child, new Map(), presets)),
    };
  }
  const box = base.box;
  return {
    ...base,
    type: "table",
    columns: [...element.columns],
    rows: element.rows.map((row) => ({
      height: row.height ?? (element.rows.length === 0 ? 0 : box.height / element.rows.length),
      cells: row.cells.map((cell) => ({
        content: normalizeTextContent(
          cell.content,
          normalizeTextParagraphStyle(presets[cell.textStylePreset ?? ""]?.paragraph),
          normalizeTextRunStyle(presets[cell.textStylePreset ?? ""]?.run),
        ),
        rowSpan: cell.rowSpan ?? 1,
        colSpan: cell.colSpan ?? 1,
        style: normalizeCellStyle(cell, presets),
      })),
    })),
  };
}

function normalizeLayout(layout: SlideLayoutDefinition, presets: Readonly<Record<string, TextStylePresetInput>>): NormalizedSlideLayout {
  const placeholders = layout.placeholders.map((placeholder) => normalizePlaceholder(placeholder, presets));
  const placeholderMap = new Map(placeholders.map((placeholder) => [placeholder.key, placeholder]));
  return {
    id: layout.id,
    name: layout.name,
    background: normalizePaint(layout.background, DEFAULT_BACKGROUND),
    elements: layout.elements.map((element) => normalizeElement(element, placeholderMap, presets)),
    placeholders,
  };
}

function defaultLayout(): NormalizedSlideLayout {
  return { id: DEFAULT_LAYOUT_ID, name: "Blank", background: deepClone(DEFAULT_BACKGROUND), elements: [], placeholders: [] };
}

function normalizeSlide(
  slide: PresentationSlide,
  layouts: ReadonlyMap<string, NormalizedSlideLayout>,
  sourceLayouts: ReadonlyMap<string, SlideLayoutDefinition>,
  presets: Readonly<Record<string, TextStylePresetInput>>,
): NormalizedPresentation["slides"][number] {
  const layoutId = slide.layoutId ?? DEFAULT_LAYOUT_ID;
  const layout = layouts.get(layoutId)!;
  const sourceLayout = sourceLayouts.get(layoutId);
  const placeholderMap = new Map(layout.placeholders.map((placeholder) => [placeholder.key, placeholder]));
  const backgroundSource = slide.background !== undefined ? "slide" : sourceLayout?.background !== undefined ? "layout" : "theme";
  return {
    id: slide.id,
    layoutId,
    background: normalizePaint(slide.background, layout.background),
    backgroundSource,
    elements: slide.elements.map((element) => normalizeElement(element, placeholderMap, presets)),
    notes: normalizeTextContent(slide.notes ?? ""),
    hidden: slide.hidden,
    ...(slide.section !== undefined ? { section: slide.section } : {}),
    tags: [...slide.tags],
    customData: deepClone(slide.customData),
  };
}

function normalizeAsset(asset: PresentationDocument["assets"][number]): NormalizedAsset {
  return {
    id: asset.id,
    kind: asset.kind,
    source: deepClone(asset.source),
    ...(asset.mimeType !== undefined ? { mimeType: asset.mimeType } : {}),
    ...(asset.width !== undefined ? { width: asset.width } : {}),
    ...(asset.height !== undefined ? { height: asset.height } : {}),
    accessibility: {
      ...(asset.accessibility?.description !== undefined ? { description: asset.accessibility.description } : {}),
      decorative: asset.accessibility?.decorative ?? false,
    },
    ...(asset.dedupeKey !== undefined ? { dedupeKey: asset.dedupeKey } : {}),
  };
}

export class PresentationNormalizer {
  normalize(document: PresentationDocument): NormalizedPresentation {
    const normalizedLayouts = document.layouts.map((layout) => normalizeLayout(layout, document.textStylePresets));
    if (!normalizedLayouts.some((layout) => layout.id === DEFAULT_LAYOUT_ID)) normalizedLayouts.unshift(defaultLayout());
    const layouts = new Map(normalizedLayouts.map((layout) => [layout.id, layout]));
    const sourceLayouts = new Map(document.layouts.map((layout) => [layout.id, layout]));
    return {
      irVersion: 1,
      id: document.id,
      metadata: normalizeMetadata(document),
      size: normalizeSize(document.size, DEFAULT_PRESENTATION_SIZE),
      theme: normalizeTheme(document),
      assets: document.assets.map(normalizeAsset),
      layouts: normalizedLayouts,
      slides: document.slides.map((slide) => normalizeSlide(slide, layouts, sourceLayouts, document.textStylePresets)),
    };
  }
}
