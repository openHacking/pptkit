import type { EastAsianVerticalMeasurer, FontRequest, MeasurementRun, TextMeasurementOptions, TextMeasurementRequest, VerticalTextDirection, VerticalTextMeasurer } from "../types/analysis.js";
import { formatListMarker } from "@pptkit/core";

/** Loaded fonts and explicit glyph/style coverage supplied by the application. */
export interface BrowserMeasurementFont {
  family: string;
  faces: readonly FontFace[];
  /** Must verify glyph and requested style coverage from the application's font inventory. */
  covers: (request: FontRequest) => boolean;
}

export interface BrowserTextMeasurementOptions {
  document: Document;
  fonts: readonly BrowserMeasurementFont[];
  /** Ordered, explicitly permitted fallback families from fonts. Default: none. */
  fallbackFamilies?: readonly string[];
  /** Choose a bounded browser estimate, a PowerPoint-observed profile, or rejection. Default: approximate. */
  tabWrap?: "approximate" | "powerpoint" | "reject";
}

export interface BrowserTextMeasurement extends TextMeasurementOptions {
  measureEastAsianVertical: EastAsianVerticalMeasurer;
  measureVertical: VerticalTextMeasurer;
  /** Remove the private measurement host. Further use throws. Idempotent. */
  dispose(): void;
}

function familyName(value: string): string {
  return value.replace(/^(['"])(.*)\1$/, "$2");
}

function cssFamily(value: string): string {
  return `"${value.replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/[\n\r\f]/g, " ")}"`;
}

function fontRequest(run: MeasurementRun, text = run.text): FontRequest {
  return { family: run.style.fontFamily, fontSize: run.style.fontSize, bold: run.style.bold, italic: run.style.italic, language: run.style.language, text };
}

/** Only scripts with verified left-to-right ordering belong in this tab profile. */
function supportedTabText(text: string): boolean {
  for (const character of text) {
    const code = character.codePointAt(0)!;
    if (/^\p{Mark}$/u.test(character)) return false;
    if (/^[\x20-\x7e\t\r\n]$/u.test(character) || /^\p{Script=Han}$/u.test(character) ||
        (code >= 0x3041 && code <= 0x3096) || (code >= 0x30a1 && code <= 0x30fa) || (code >= 0xac00 && code <= 0xd7a3)) continue;
    return false;
  }
  return true;
}

/** Browser CSS typography analysis, not a native Office line-layout engine. */
export function createBrowserTextMeasurement(options: BrowserTextMeasurementOptions): BrowserTextMeasurement {
  const document = options.document;
  if (!document?.documentElement || !document.defaultView || !document.fonts) throw new TypeError("A live browser document with Font Loading support is required");
  if (options.tabWrap !== undefined && options.tabWrap !== "approximate" && options.tabWrap !== "powerpoint" && options.tabWrap !== "reject") throw new TypeError("tabWrap must be approximate, powerpoint or reject");
  const tabWrap = options.tabWrap ?? "approximate";
  const fonts = new Map<string, BrowserMeasurementFont>();
  for (const font of options.fonts) {
    if (!font.family.trim() || typeof font.covers !== "function" || font.faces.length === 0 || fonts.has(font.family)) throw new TypeError("Supply unique font families, loaded faces and an explicit coverage function");
    if (font.faces.some((face) => familyName(face.family) !== font.family)) throw new TypeError("FontFace family must match its registered family");
    fonts.set(font.family, { ...font, faces: [...font.faces] });
  }
  const fallbacks = [...(options.fallbackFamilies ?? [])];
  if (fallbacks.some((family) => !fonts.has(family))) throw new TypeError("Fallback families must belong to the font inventory");
  const host = document.createElement("div");
  host.setAttribute("aria-hidden", "true");
  host.setAttribute("data-pptkit-measurement", "");
  host.style.cssText = "all:initial!important;position:fixed!important;left:0!important;top:0!important;width:0!important;height:0!important;visibility:hidden!important;pointer-events:none!important;overflow:hidden!important;contain:layout style paint!important;transform:none!important;zoom:1!important";
  const shadow = host.attachShadow({ mode: "closed" });
  document.documentElement.appendChild(host);
  let disposed = false;
  function active(): void {
    if (disposed || !host.isConnected) throw new Error("Browser text measurement has been disposed or detached");
    const style = document.defaultView!.getComputedStyle(document.documentElement);
    if (style.transform !== "none" || (style.zoom !== "normal" && Number(style.zoom) !== 1)) throw new Error("Measurement document root must not be transformed or zoomed");
  }
  function available(family: string, request: FontRequest): boolean {
    const font = fonts.get(family);
    if (!font || font.faces.some((face) => face.status !== "loaded" || familyName(face.family) !== family || !document.fonts.has(face))) return false;
    const candidate = Object.freeze({ ...request, family });
    const covered: unknown = font.covers(candidate);
    if (typeof covered !== "boolean") {
      // JavaScript callers can bypass the synchronous coverage contract.
      void Promise.resolve(covered).catch(() => undefined);
      throw new TypeError("Font coverage must return a synchronous boolean");
    }
    if (!covered) return false;
    return document.fonts.check(`${request.italic ? "italic" : "normal"} ${request.bold ? 700 : 400} ${request.fontSize}px ${cssFamily(family)}`, request.text);
  }
  function styleRun(span: HTMLElement, run: MeasurementRun): void {
    span.style.fontFamily = cssFamily(run.style.fontFamily);
    span.style.fontSize = `${run.style.fontSize}px`;
    span.style.fontWeight = run.style.bold ? "700" : "400";
    span.style.fontStyle = run.style.italic ? "italic" : "normal";
    span.style.textDecoration = [run.style.underline ? "underline" : "", run.style.strike ? "line-through" : ""].filter(Boolean).join(" ") || "none";
    span.lang = run.style.language;
  }
  function measure(request: TextMeasurementRequest, vertical?: VerticalTextDirection) {
    active();
    if (!Number.isFinite(request.availableWidth) || request.availableWidth < 0) throw new RangeError("Measurement width must be finite and nonnegative");
    const rightToLeft = vertical === "eastAsianVertical" || vertical === "stackedVerticalRtl";
    const writingMode = rightToLeft ? "vertical-rl" : "vertical-lr";
    const orientation = vertical === "stackedVertical" || vertical === "stackedVerticalRtl" ? "upright" : "mixed";
    const root = document.createElement("div");
    root.style.cssText = `all:initial;display:flow-root;box-sizing:content-box;${vertical ? `height:${request.availableWidth}px;width:max-content;writing-mode:${writingMode};text-orientation:${orientation}` : `width:${request.availableWidth}px;writing-mode:horizontal-tb`};font-size:0;line-height:0;direction:ltr;font-synthesis:none;white-space:${request.wrap ? "pre-wrap" : "pre"};overflow-wrap:${request.wrap ? "anywhere" : "normal"};`;
    const spans: HTMLElement[] = [];
    const tabbed: { block: HTMLElement; groups: HTMLElement[]; stopOrigin: number; stops: NonNullable<TextMeasurementRequest["paragraphs"][number]["style"]["tabStops"]>; softWrapEligible: boolean }[] = [];
    shadow.appendChild(root);
    try {
      for (const paragraph of request.paragraphs) {
        if (paragraph.runs.some(run => run.style.rightToLeft === true)) throw new Error("Browser measurement cannot reproduce DrawingML run-level RTL character overrides.");
        const first = paragraph.runs[0];
        if (!first) continue;
        const bullet = formatListMarker(paragraph.style.bullet);
        const block = document.createElement("div");
        styleRun(block, first);
        const style = paragraph.style;
        block.style.cssText += `;display:block;box-sizing:border-box;margin:0;padding:${style.spaceBefore}px 0 ${style.spaceAfter}px ${style.indent}px;text-indent:${-style.hanging}px;line-height:${style.lineSpacing};text-align:${style.align};`;
        block.style.direction = style.direction ?? "ltr";
        const source = paragraph.runs.map(run => run.text).join("");
        const hasTabs = source.includes("\t");
        if (vertical) {
          if (hasTabs || style.bullet.type !== "none" || style.direction === "rtl")
            throw new Error("Vertical measurement does not support tabs, list markers, or RTL paragraph direction.");
          block.style.padding = "0";
          block.style.paddingInlineStart = `${style.indent}px`;
          block.style.paddingBlockStart = `${style.spaceBefore}px`;
          block.style.paddingBlockEnd = `${style.spaceAfter}px`;
          if (style.align === "left" || style.align === "right") block.style.textAlign = style.align === "left" ? "start" : "end";
        }
        const supportedList = style.bullet.type === "none" || (first.text.length > 0 && (style.bullet.type === "bullet" ? style.bullet.character === "•" : style.bullet.type === "number" && style.bullet.style === "arabicPeriod"));
        if (hasTabs && (style.direction === "rtl" || style.align !== "left" || style.hanging > style.indent || !supportedList || !style.tabStops?.length || style.tabStops.length > 256 || /\t(?:\r|\n|$)/u.test(source) || !supportedTabText(source))) {
          throw new Error("Browser tab-stop measurement requires LTR direction, 1–256 explicit stops, printable ASCII, Han, precomposed kana or Hangul, left alignment, hanging no greater than indent, no line-ending tab and a supported list marker.");
        }
        // Position explicit lines first; supported automatic breaks are added
        // after the tab-aligned fragments have been measured.
        if (hasTabs) block.style.whiteSpace = "pre";
        const entries = hasTabs ? paragraph.runs : bullet ? [{ ...first, text: bullet }, ...paragraph.runs] : paragraph.runs;
        if (hasTabs) {
          let count = 0, controls = 0;
          for (const char of source) {
            if (char === "\t") { if (++count > style.tabStops!.length) throw new Error("Browser tab-stop measurement requires an explicit stop for every tab on each line."); controls++; }
            else if (char === "\r" || char === "\n") { count = 0; controls++; }
            if (controls > 4096) throw new Error("Browser tab-stop measurement supports at most 4096 tabs and hard-break characters per paragraph.");
          }
        }
        let groups: HTMLElement[] = [];
        const lines: HTMLElement[][] = [groups];
        const addGroup = () => { const group = document.createElement("span"); group.style.cssText = "display:inline-block;white-space:pre;vertical-align:baseline;text-indent:0"; block.appendChild(group); groups.push(group); return group; };
        let group = hasTabs ? addGroup() : undefined;
        if (hasTabs && bullet) {
          const marker = bullet.slice(0, -1), markerRun = { ...first, text: marker };
          if (!available(markerRun.style.fontFamily, fontRequest(markerRun))) throw new Error(`Font coverage or loading is unavailable for "${markerRun.style.fontFamily}"`);
          const span = document.createElement("span"); styleRun(span, markerRun); span.style.display = "inline-block"; span.style.minWidth = `${style.hanging}px`; span.textContent = marker;
          group!.appendChild(span); spans.push(span);
        }
        let previousCR = false;
        for (const run of entries) {
          if (!available(run.style.fontFamily, fontRequest(run))) throw new Error(`Font coverage or loading is unavailable for "${run.style.fontFamily}"`);
          const append = (text: string) => {
            const span = document.createElement("span"); styleRun(span, run); span.textContent = text;
            (group ?? block).appendChild(span); spans.push(span);
          };
          if (!hasTabs) { append(run.text); continue; }
          if (run.text === "") { append(""); continue; }
          const parts = run.text.split(/(\r\n|\r|\n|\t)/u);
          for (const part of parts) {
            if (part === "") continue;
            if (part === "\n" && previousCR) { previousCR = false; continue; }
            previousCR = false;
            if (part === "\t") { append("\u200b"); group = addGroup(); }
            else if (part === "\r" || part === "\n" || part === "\r\n") {
              // A break can have its own font size, including on an empty line.
              // Keep that line-box contribution when replacing the character with <br>.
              append("\u200b");
              const lineBreak = document.createElement("br"); styleRun(lineBreak, run); block.appendChild(lineBreak);
              groups = []; lines.push(groups); group = addGroup(); previousCR = part === "\r";
            } else append(part);
          }
        }
        // Preserve empty paragraphs and the final empty line after an explicit break.
        const text = entries.map((run) => run.text).join("");
        if (text === "" || /[\n\r]$/.test(text)) {
          const emptyLine = document.createElement("span");
          styleRun(emptyLine, entries.at(-1) ?? first);
          emptyLine.textContent = "\u200b";
          block.appendChild(emptyLine);
          spans.push(emptyLine);
        }
        root.appendChild(block);
        if (hasTabs) for (const line of lines) tabbed.push({ block, groups: line, stopOrigin: bullet ? 0 : style.indent, stops: style.tabStops!, softWrapEligible: !bullet && style.indent === 0 && style.hanging === 0 });
      }
      if (root.getClientRects().length === 0) throw new Error("Measurement document has no active CSS layout");
      // Read all fragment metrics before writing margins to avoid forcing a
      // full document reflow for every tab in large bounded paragraphs.
      const placements: { group: HTMLElement; margin: number }[] = [];
      const groupAlignments = new Map<HTMLElement, string>();
      const overlapping = new Set<HTMLElement>();
      for (const { block, groups, stopOrigin, stops, softWrapEligible } of tabbed) {
        let cursor = groups[0]!.getBoundingClientRect().right - block.getBoundingClientRect().left;
        for (const group of groups.slice(1)) {
          const stop = stops.find(candidate => stopOrigin + candidate.position > cursor);
          if (!stop) throw new Error("Browser tab-stop measurement requires an explicit stop beyond each preceding text segment.");
          groupAlignments.set(group, stop.align);
          const width = group.getBoundingClientRect().width;
          let offset = stop.align === "right" ? width : stop.align === "center" ? width / 2 : 0;
          if (stop.align === "decimal") {
            const text = group.textContent ?? "", decimal = text.indexOf(".");
            if (decimal < 0) throw new Error("Browser decimal tab-stop measurement requires an explicit period in each decimal-aligned segment.");
            let remaining = decimal; const range = document.createRange(); range.setStart(group, 0);
            for (const node of Array.from(group.childNodes)) {
              const value = node.textContent ?? "";
              if (remaining <= value.length) { range.setEnd(node.firstChild ?? node, remaining); break; }
              remaining -= value.length;
            }
            offset = decimal === 0 ? 0 : range.getBoundingClientRect().width;
          }
          const margin = stopOrigin + stop.position - offset - cursor;
          if (margin < -0.001) {
            if (tabWrap !== "powerpoint" || !request.wrap || !softWrapEligible || groups.length !== 2 || stop.align === "left") throw new Error("Browser tab-stop measurement cannot position a segment that overlaps its preceding text.");
            overlapping.add(group);
          }
          const advance = overlapping.has(group) ? margin : Math.max(0, margin);
          placements.push({ group, margin: advance });
          cursor += advance + width;
        }
      }
      for (const { group, margin } of placements) group.style.marginLeft = `${margin}px`;
      if (request.wrap) {
        const frame = root.getBoundingClientRect();
        const wraps: { group: HTMLElement; breaks: number; margin: number | undefined; anywhere: boolean }[] = [];
        for (const { groups, softWrapEligible } of tabbed) if (overlapping.has(groups.at(-1)!) || groups.at(-1)!.getBoundingClientRect().right > frame.right + 0.001) {
          if (tabWrap === "reject") throw new Error("Browser tab-stop measurement rejects automatic wrapping under the selected tabWrap policy.");
          const prefix = (groups[0]!.textContent ?? "").replace(/\u200b/g, "");
          const suffix = (groups.at(-1)!.textContent ?? "").replace(/\u200b/g, "");
          if (!softWrapEligible || groups.length !== 2 || !/^[A-Za-z0-9]+$/u.test(prefix) || !/^[A-Za-z0-9]+(?:\.[0-9]+)?(?: [A-Za-z0-9]+(?:\.[0-9]+)?)*$/u.test(suffix) || groups[0]!.getBoundingClientRect().right > frame.right + 0.001) {
            throw new Error("Browser tab-stop measurement cannot yet resolve automatic wrapping of this tabbed line.");
          }
          const group = groups[1]!, nodes: Text[] = [], walker = document.createTreeWalker(group, NodeFilter.SHOW_TEXT);
          while (walker.nextNode()) nodes.push(walker.currentNode as Text);
          const locate = (offset: number): [Text, number] => {
            for (const node of nodes) { if (offset <= node.length) return [node, offset]; offset -= node.length; }
            throw new Error("Browser tab-stop measurement cannot locate a word boundary.");
          };
          const words = Array.from(suffix.matchAll(/[^ ]+/gu));
          if (words.length > 256) throw new Error("Browser tab-stop measurement supports at most 256 words on a wrapped tabbed line.");
          const alignment = groupAlignments.get(group);
          const powerpointTextSupported = ((alignment === "left" || alignment === "center" || alignment === "right") && /^[A-Za-z]+(?: [A-Za-z]+)+$/u.test(suffix)) || (alignment === "decimal" && /^[0-9]+\.[0-9]+(?: [A-Za-z]+)+$/u.test(suffix));
          if (tabWrap === "powerpoint" && !powerpointTextSupported) {
            throw new Error("Browser tab-stop measurement cannot yet resolve automatic wrapping under the PowerPoint profile for this line.");
          }
          let firstWordOverflows = false, firstWordWidth = 0;
          for (const [index, word] of words.entries()) {
            const range = document.createRange(), [startNode, startOffset] = locate(word.index!), [endNode, endOffset] = locate(word.index! + word[0].length);
            range.setStart(startNode, startOffset); range.setEnd(endNode, endOffset);
            const width = range.getBoundingClientRect().width;
            if (width > frame.width + 0.001) throw new Error("Browser tab-stop measurement cannot yet resolve automatic wrapping of this tabbed line.");
            if (index === 0) { firstWordWidth = width; firstWordOverflows = group.getBoundingClientRect().left + width > frame.right + 0.001; }
          }
          if (tabWrap === "powerpoint") {
            if (alignment === "center" || alignment === "right" || alignment === "decimal") {
              const groupRect = group.getBoundingClientRect(), prefixRight = groups[0]!.getBoundingClientRect().right;
              if (prefixRight + firstWordWidth > frame.right + 0.001) throw new Error("Browser tab-stop measurement cannot position a segment that overlaps its preceding text.");
              let lineWidth = firstWordWidth;
              for (const word of words.slice(1)) {
                const range = document.createRange(), [startNode, startOffset] = locate(0), [endNode, endOffset] = locate(word.index! + word[0].length);
                range.setStart(startNode, startOffset); range.setEnd(endNode, endOffset);
                const candidateWidth = range.getBoundingClientRect().width;
                if (prefixRight + candidateWidth > frame.right + 0.001) break;
                lineWidth = candidateWidth;
              }
              const alignedStart = alignment === "center" ? groupRect.left + (groupRect.width - lineWidth) / 2 : alignment === "right" ? groupRect.left + groupRect.width - lineWidth : groupRect.left;
              const desiredStart = Math.max(prefixRight, Math.min(alignedStart, frame.right - lineWidth));
              wraps.push({ group, breaks: 0, margin: Math.max(0, desiredStart - prefixRight), anywhere: false });
              continue;
            }
            let breaks = 0, margin: number | undefined;
            if (firstWordOverflows) {
              const range = document.createRange(), [startNode, startOffset] = locate(0), [endNode, endOffset] = locate(1);
              range.setStart(startNode, startOffset); range.setEnd(endNode, endOffset);
              const firstCharacterOverflows = group.getBoundingClientRect().left + range.getBoundingClientRect().width > frame.right + 0.001;
              breaks = firstCharacterOverflows ? 2 : 1;
              margin = firstCharacterOverflows ? 0 : group.getBoundingClientRect().left - frame.left;
            }
            wraps.push({ group, breaks, margin, anywhere: true });
          } else wraps.push({ group, breaks: firstWordOverflows ? 1 : 0, margin: firstWordOverflows ? 0 : undefined, anywhere: false });
        }
        for (const { group, breaks, margin, anywhere } of wraps) {
          group.style.display = "inline";
          group.style.whiteSpace = "normal";
          group.style.overflowWrap = anywhere ? "anywhere" : "normal";
          if (margin !== undefined) group.style.marginLeft = `${margin}px`;
          for (let index = 0; index < breaks; index++) group.parentElement!.insertBefore(document.createElement("br"), group);
        }
      }
      const origin = root.getBoundingClientRect();
      if (vertical) {
        const computed = document.defaultView!.getComputedStyle(root);
        if (computed.writingMode !== writingMode || computed.textOrientation !== orientation)
          throw new Error("This browser does not support the requested vertical CSS layout.");
        let inlineStart = 0, inlineEnd = 0, blockStart = 0, blockEnd = origin.width;
        for (const span of spans) {
          const range = document.createRange();
          range.selectNodeContents(span);
          for (const rect of Array.from(range.getClientRects())) {
            inlineStart = Math.min(inlineStart, rect.top - origin.top);
            inlineEnd = Math.max(inlineEnd, rect.bottom - origin.top);
            blockStart = Math.min(blockStart, rightToLeft ? origin.right - rect.right : rect.left - origin.left);
            blockEnd = Math.max(blockEnd, rightToLeft ? origin.right - rect.left : rect.right - origin.left);
          }
        }
        return { bounds: { x: inlineStart, y: blockStart, width: inlineEnd - inlineStart, height: blockEnd - blockStart } };
      }
      let left = 0;
      let top = 0;
      let right = 0;
      let bottom = origin.height;
      for (const span of spans) {
        const range = document.createRange();
        range.selectNodeContents(span);
        for (const rect of Array.from(range.getClientRects())) {
          left = Math.min(left, rect.left - origin.left);
          top = Math.min(top, rect.top - origin.top);
          right = Math.max(right, rect.right - origin.left);
          bottom = Math.max(bottom, rect.bottom - origin.top);
        }
      }
      return { bounds: { x: left, y: top, width: right - left, height: bottom - top } };
    } finally { root.remove(); }
  }
  return {
    resolveFont(request) {
      active();
      for (const family of [request.family, ...fallbacks]) if (available(family, request)) return { family };
      return undefined;
    },
    measure: request => measure(request),
    measureEastAsianVertical(request) {
      if (!Number.isFinite(request.availableBlockSize) || request.availableBlockSize < 0)
        throw new RangeError("Measurement block size must be finite and nonnegative");
      return measure({ paragraphs: request.paragraphs, availableWidth: request.availableInlineSize, wrap: request.wrap }, "eastAsianVertical");
    },
    measureVertical(request) {
      if (!Number.isFinite(request.availableBlockSize) || request.availableBlockSize < 0)
        throw new RangeError("Measurement block size must be finite and nonnegative");
      if (request.textDirection !== "eastAsianVertical" && request.textDirection !== "mongolianVertical" && request.textDirection !== "stackedVertical" && request.textDirection !== "stackedVerticalRtl")
        throw new TypeError("Unsupported vertical writing mode");
      return measure({ paragraphs: request.paragraphs, availableWidth: request.availableInlineSize, wrap: request.wrap }, request.textDirection);
    },
    dispose() { if (!disposed) { disposed = true; host.remove(); } },
  };
}
