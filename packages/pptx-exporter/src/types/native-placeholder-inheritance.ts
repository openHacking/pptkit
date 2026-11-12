/** Internal, captured authoring intent used by package writers. */
export type NativeInheritedRunProperty = "fontFamily" | "color" | "bold" | "italic" | "underline" | "strike" | "fontSize" | "language" | "rightToLeft";
export type NativeInheritedParagraphProperty = "paragraph.align" | "paragraph.direction" | "paragraph.indent" | "paragraph.hanging" | "paragraph.lineSpacing" | "paragraph.spaceBefore" | "paragraph.spaceAfter" | "paragraph.bullet";
export type NativeInheritedFrameProperty = "frame.textDirection" | "frame.wrap" | "frame.verticalAlign" | "frame.autoFit" | "frame.margin.left" | "frame.margin.right" | "frame.margin.top" | "frame.margin.bottom";
export type NativeInheritedProperty = NativeInheritedRunProperty | NativeInheritedParagraphProperty;
export interface NativeInheritedText {
  readonly runs: readonly (readonly (readonly NativeInheritedRunProperty[])[])[];
  readonly paragraphs: readonly (readonly NativeInheritedParagraphProperty[])[];
  readonly frame: readonly NativeInheritedFrameProperty[];
}
export type NativeInheritedRuns = ReadonlyMap<string, NativeInheritedText>;

export type NativeInheritedPlaceholderDefaults = ReadonlyMap<string, ReadonlySet<NativeInheritedProperty>>;
export interface NativePlaceholderInheritance {
  readonly layouts: ReadonlyMap<string, NativeInheritedPlaceholderDefaults>;
  readonly slides: ReadonlyMap<string, NativeInheritedRuns>;
}
