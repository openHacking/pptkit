import type { LayoutRebindApplyResult, LayoutRebindPlan } from "./layout-rebind.js";
import type { EventTimelinePlan, NormalizedSlideEventTimeline, SlideEventTimelineInput } from "./event-timeline.js";
import type { AssetReference, AssetRegistrationOptions, AssetRegistrationRequest, AssetCopyOptions, AssetCopyMapping, AssetRetentionOptions, PresentationAssetReplacement, NormalizedAsset, PresentationAsset, PresentationAssetInput } from "./asset.js";
import type {
  NormalizedElement,
  NormalizedPlaceholderDefinition,
  NormalizedTextParagraph,
  PlaceholderDefinitionInput,
  PresentationElement,
  PresentationElementInput,
  TextContentInput,
} from "./element.js";
import type { PresentationSize } from "./geometry.js";
import type { BackgroundPaintInput, NormalizedBackgroundPaint, TextStylePresetMap } from "./style.js";
import type { NormalizedPresentationTheme, PresentationThemeInput } from "./theme.js";

export type JsonValue = null | boolean | number | string | JsonValue[] | { [key: string]: JsonValue };

export type CustomDocumentProperty =
  | { name: string; type: "string"; value: string }
  | { name: string; type: "boolean"; value: boolean }
  | { name: string; type: "integer" | "real"; value: number }
  | { name: string; type: "date"; value: string }
  | { name: string; type: "decimal"; value: string };

export interface PresentationMetadataInput {
  title?: string;
  author?: string;
  company?: string;
  subject?: string;
  description?: string;
  language?: string;
  keywords?: string[];
  revision?: number;
  createdAt?: string;
  modifiedAt?: string;
  customProperties?: readonly CustomDocumentProperty[];
}

export interface NormalizedPresentationMetadata {
  title: string;
  author: string;
  company?: string;
  subject?: string;
  description?: string;
  language: string;
  keywords: string[];
  revision: number;
  createdAt?: string;
  modifiedAt?: string;
  customProperties?: readonly CustomDocumentProperty[];
}

export type SlideTransitionInput = {
  /** Visual transition duration in milliseconds. */
  readonly durationMs: number;
  /** Whether a click may advance from this slide. */
  readonly onClick?: boolean;
  /** Automatic advance delay in milliseconds. */
  readonly afterMs?: number;
} & ({ readonly type: "fade" } | { readonly type: "push" | "wipe"; readonly direction: "left" | "right" | "up" | "down" });

/** Independent reading order for slide roots or the direct children of a local group. */
export interface SlideReadingOrder {
  readonly groupId?: string;
  readonly elementIds: readonly string[];
}

export interface PresentationSlideInput {
  /** False hides ordinary layout and master artwork; background and placeholders remain. */
  showInheritedShapes?: boolean;
  id?: string;
  layoutId?: string;
  background?: BackgroundPaintInput;
  elements?: PresentationElementInput[];
  notes?: TextContentInput;
  transition?: SlideTransitionInput;
  eventTimeline?: SlideEventTimelineInput;
  readingOrder?: readonly SlideReadingOrder[];
  hidden?: boolean;
  section?: string;
  tags?: string[];
  customData?: Record<string, JsonValue>;
}

/** Explicit reconciliation choices for a detached cross-document slide copy. */
export interface SlideCopyOptions {
  index?: number;
  /** Target layout identity, or null for the destination's default layout. */
  layoutId: string | null;
  theme: "destination";
  /** Preserve authored coordinates, or uniformly fit explicit page-coordinate geometry. */
  geometry: "preserve" | "fit-explicit-geometry";
  /** Preserve opaque slide metadata unchanged, or reject pages carrying any custom data. */
  customData: "preserve" | "reject";
  assets: AssetCopyOptions;
  /** Source slide IDs other than the copied page mapped to existing destination IDs. */
  slideLinks?: Readonly<Record<string, string>>;
  /** External connector source identities mapped to objects in the destination layout/master scope. */
  connectorTargets?: Readonly<Record<string, string>>;
  /** Require matching roles, optionally also requiring equal explicit numeric indexes. */
  placeholderRoles?: "require-match" | "require-match-and-index";
  /** Resolve unmapped placeholder keys by explicit index; requires strict role-and-index matching. */
  placeholderMatching?: "explicit-index";
  placeholders?: Readonly<Record<string, string>>;
  textPresets?: Readonly<Record<string, string>>;
}

export interface SlideCopyResult {
  readonly slide: ManagedPresentationSlide;
  readonly elementIds: Readonly<Record<string, string>>;
  readonly assets: readonly AssetCopyMapping[];
}

export interface SlidesCopyOptions extends SlideCopyOptions {
  /** Optional target layout per source slide; absent entries use layoutId. */
  layouts?: Readonly<Record<string, string | null>>;
  /** Per-page mapping entries override the corresponding shared entries. */
  perSlide?: Readonly<Record<string, {
    readonly connectorTargets?: Readonly<Record<string, string>>;
    readonly placeholders?: Readonly<Record<string, string>>;
    readonly textPresets?: Readonly<Record<string, string>>;
  }>>;
}

export interface SlidesCopyResult {
  readonly slides: readonly ManagedPresentationSlide[];
  readonly slideIds: Readonly<Record<string, string>>;
  readonly elementIds: Readonly<Record<string, Readonly<Record<string, string>>>>;
  readonly assets: readonly AssetCopyMapping[];
}

export interface LayoutCopyOptions {
  id: string;
  /** Reuse the first exactly equivalent existing layout when all resources are already present. */
  reuseEquivalent?: boolean;
  /** Destination master identity, or null to use the presentation theme directly. */
  masterId: string | null;
  theme: "destination";
  assets: AssetCopyOptions;
  slideLinks?: Readonly<Record<string, string>>;
  /** External connector source identities mapped to objects in the destination master scope. */
  connectorTargets?: Readonly<Record<string, string>>;
  textPresets?: Readonly<Record<string, string>>;
}

export interface LayoutCopyResult {
  readonly layout: SlideLayoutDefinition;
  readonly elementIds: Readonly<Record<string, string>>;
  readonly assets: readonly AssetCopyMapping[];
}

export interface MasterCopyOptions {
  id: string;
  /** Reuse the first exactly equivalent existing master when all resources are already present. */
  reuseEquivalent?: boolean;
  /** Copy local declarations, materialize the effective source theme, or inherit the destination theme. */
  theme: "source-declarations" | "source-effective" | "destination";
  assets: AssetCopyOptions;
  slideLinks?: Readonly<Record<string, string>>;
  textPresets?: Readonly<Record<string, string>>;
}

export interface MasterCopyResult {
  readonly master: SlideMasterDefinition;
  readonly elementIds: Readonly<Record<string, string>>;
  readonly assets: readonly AssetCopyMapping[];
}

export interface PresentationSlide {
  readonly showInheritedShapes?: boolean | undefined;
  readonly id: string;
  readonly layoutId: string | undefined;
  readonly background: BackgroundPaintInput | undefined;
  readonly elements: readonly PresentationElement[];
  readonly notes: TextContentInput | undefined;
  readonly transition: SlideTransitionInput | undefined;
  readonly eventTimeline: SlideEventTimelineInput | undefined;
  readonly readingOrder?: readonly SlideReadingOrder[] | undefined;
  readonly hidden: boolean;
  /** Set slide visibility when supported by the owning document. */
  setHidden?(hidden: boolean): void;
  readonly section: string | undefined;
  readonly tags: readonly string[];
  readonly customData: Readonly<Record<string, JsonValue>>;
  /** Replace or clear speaker notes with an owned immutable snapshot. */
  setNotes(notes?: TextContentInput): void;
  /** Replace or clear the slide transition with an owned immutable snapshot. */
  setTransition(transition?: SlideTransitionInput): void;
  /** Replace or clear a whole-object foreground timing program. */
  setEventTimeline(timeline?: SlideEventTimelineInput): void;
  /** Compile against the current owned object inventory. */
  getEventTimelinePlan(): EventTimelinePlan | undefined;
  /** Replace a top-level object atomically while retaining its root identity. */
  replaceElement(elementId: string, input: PresentationElementInput): PresentationElement;
  addElement(input: PresentationElementInput): PresentationElement;
  insertElement(index: number, input: PresentationElementInput): PresentationElement;
  moveElement(elementId: string, toIndex: number): void;
  removeElement(elementId: string): PresentationElement;
  duplicateElement(elementId: string, toIndex?: number): PresentationElement;
}

/** Editing capabilities on factory-owned slides, without extending legacy custom requirements. */
export interface ManagedPresentationSlide extends PresentationSlide {
  setReadingOrder(order?: readonly SlideReadingOrder[]): void;
  /** Set whether the slide is hidden during slideshow playback. */
  setHidden(hidden: boolean): void;
  setShowInheritedShapes(value?: boolean): void;
  /** Atomically reorder every top-level object by ID, changing its drawing order. */
  reorderElements(elementIds: readonly string[]): void;
  /** Atomically reorder the direct children of an owned group by ID. */
  reorderGroupElements(groupId: string, elementIds: readonly string[]): void;
  /** Replace or clear the local background with an owned immutable snapshot. */
  setBackground(value?: BackgroundPaintInput): void;
}

/** A shared theme/background and ordinary-artwork owner. */
export interface SlideMasterInput {
  elements?: PresentationElementInput[];
  id: string;
  name: string;
  theme?: PresentationThemeInput;
  background?: BackgroundPaintInput;
}
export interface SlideMasterDefinition {
  readonly elements?: readonly PresentationElement[];
  readonly id: string;
  readonly name: string;
  readonly theme?: Readonly<PresentationThemeInput>;
  readonly background?: BackgroundPaintInput;
}
export interface NormalizedSlideMaster {
  elements?: NormalizedElement[];
  id: string;
  name: string;
  theme: NormalizedPresentationTheme;
  background?: NormalizedBackgroundPaint;
}

export interface SlideLayoutInput {
  /** False hides ordinary master artwork beneath this layout. */
  showMasterShapes?: boolean;
  masterId?: string;
  id: string;
  name: string;
  background?: BackgroundPaintInput;
  elements?: PresentationElementInput[];
  placeholders?: PlaceholderDefinitionInput[];
}

export interface SlideLayoutDefinition {
  readonly showMasterShapes?: boolean;
  readonly masterId?: string;
  readonly id: string;
  readonly name: string;
  readonly background?: BackgroundPaintInput;
  readonly elements: readonly PresentationElement[];
  readonly placeholders: readonly PlaceholderDefinitionInput[];
}

/** Ordered, format-independent grouping of slides. IDs are Core identities, not native GUIDs. */
export interface PresentationSection {
  readonly id: string;
  readonly name?: string;
  readonly slideIds: readonly string[];
}

export interface PresentationInit {
  masters?: readonly SlideMasterInput[];
  /** Optional nonnegative starting slide number; undefined leaves the host default. */
  firstSlideNumber?: number;
  id?: string;
  metadata?: PresentationMetadataInput;
  size?: Partial<PresentationSize>;
  theme?: PresentationThemeInput;
  textStylePresets?: TextStylePresetMap;
}

export interface PresentationDocument {
  /** Optional ordered section collection; legacy documents may omit it. */
  readonly sections?: readonly PresentationSection[] | undefined;
  /** Atomically replace the explicit section collection when supported. */
  setSections?(sections?: readonly PresentationSection[]): void;
  readonly firstSlideNumber?: number | undefined;
  /** Available on factory-managed documents; legacy custom documents may omit it. */
  setFirstSlideNumber?(value?: number): void;
  /** Absent in legacy custom implementations; equivalent to no explicit masters. */
  readonly masters?: readonly SlideMasterDefinition[];
  readonly id: string;
  readonly metadata: Readonly<PresentationMetadataInput>;
  /** Replace metadata when supported by the owning document. */
  setMetadata?(metadata: PresentationMetadataInput): void;
  readonly size: Readonly<PresentationSize>;
  readonly theme: Readonly<PresentationThemeInput>;
  readonly textStylePresets: TextStylePresetMap;
  readonly slides: readonly PresentationSlide[];
  readonly assets: readonly PresentationAsset[];
  readonly layouts: readonly SlideLayoutDefinition[];
  addSlide(input?: PresentationSlideInput): PresentationSlide;
  insertSlide(index: number, input?: PresentationSlideInput): PresentationSlide;
  moveSlide(slideId: string, toIndex: number): void;
  removeSlide(slideId: string): PresentationSlide;
  duplicateSlide(slideId: string, toIndex?: number): PresentationSlide;
  registerAsset(input: PresentationAssetInput, options?: AssetRegistrationOptions): PresentationAsset;
  replaceAsset(assetId: string, replacement: PresentationAssetReplacement): PresentationAsset;
  getAssetReferences(assetId: string): readonly AssetReference[];
  removeAsset(assetId: string, options?: AssetRetentionOptions): PresentationAsset;
  removeUnusedAssets(options?: AssetRetentionOptions): readonly PresentationAsset[];
  getAsset(assetId: string): PresentationAsset | undefined;
  defineSlideLayout(input: SlideLayoutInput): SlideLayoutDefinition;
}

/** Editing capabilities of factory-created documents; legacy custom documents remain valid. */
export interface ManagedPresentationDocument extends PresentationDocument {
  setSections(sections?: readonly PresentationSection[]): void;
  /** Replace the complete slide order and section collection in one transaction. */
  setSectionStructure(slideOrder: readonly string[], sections?: readonly PresentationSection[]): void;
  /** Atomically apply a serializable, value-checked edit batch and return its inverse. */
  applyEditOperation(operation: import("./edit-operation.js").PresentationEditOperation): import("./edit-operation.js").PresentationEditResult;
  /** Replace the complete metadata value with a validated, detached snapshot. */
  setMetadata(metadata: PresentationMetadataInput): void;
  /** Register a batch atomically; failures do not commit batch additions. */
  registerAssets(requests: readonly AssetRegistrationRequest[]): readonly PresentationAsset[];
  /** Copy explicitly selected assets into this document as one atomic batch. */
  copyAssetsFrom(source: PresentationDocument, assetIds: readonly string[], options: AssetCopyOptions): readonly AssetCopyMapping[];
  /** Copy one slide and its image resources with explicit destination policies. */
  copySlideFrom(source: PresentationDocument, slideId: string, options: SlideCopyOptions): SlideCopyResult;
  /** Copy selected slides and their resources as one transaction, remapping links within the selection. */
  copySlidesFrom(source: PresentationDocument, slideIds: readonly string[], options: SlidesCopyOptions): SlidesCopyResult;
  /** Copy a layout into this document under an explicit identity and master policy. */
  copySlideLayoutFrom(source: PresentationDocument, layoutId: string, options: LayoutCopyOptions): LayoutCopyResult;
  /** Copy a master with explicit theme and resource policies. */
  copySlideMasterFrom(source: PresentationDocument, masterId: string, options: MasterCopyOptions): MasterCopyResult;
  /** Set or clear the authored starting slide number. */
  setFirstSlideNumber(value?: number): void;
  /** Recheck a reviewed value plan and commit all selected slides atomically. */
  applyLayoutRebind(plan: LayoutRebindPlan): LayoutRebindApplyResult;
  readonly slides: readonly ManagedPresentationSlide[];
  addSlide(input?: PresentationSlideInput): ManagedPresentationSlide;
  insertSlide(index: number, input?: PresentationSlideInput): ManagedPresentationSlide;
  duplicateSlide(slideId: string, toIndex?: number): ManagedPresentationSlide;
  readonly masters: readonly SlideMasterDefinition[];
  defineSlideMaster(input: SlideMasterInput): SlideMasterDefinition;
  /** Full replacement, retaining identity. Omitted declarations resume inheritance. */
  replaceSlideMaster(masterId: string, input: Omit<SlideMasterInput, "id"> & { id?: string }): SlideMasterDefinition;
  /** Full replacement. Invalid definitions or affected slide dependencies abort without changing ownership. */
  replaceSlideLayout(layoutId: string, input: Omit<SlideLayoutInput, "id"> & { id?: string }): SlideLayoutDefinition;
}

export interface NormalizedSlideLayout {
  showMasterShapes?: boolean;
  backgroundSource?: "layout" | "master" | "theme";
  masterId?: string;
  /** Effective owning theme, present only for explicitly master-bound layouts. */
  theme?: NormalizedPresentationTheme;
  id: string;
  name: string;
  background: NormalizedBackgroundPaint;
  elements: NormalizedElement[];
  placeholders: NormalizedPlaceholderDefinition[];
}

export interface NormalizedSlide {
  showInheritedShapes?: boolean;
  id: string;
  layoutId: string;
  background: NormalizedBackgroundPaint;
  backgroundSource: "slide" | "layout" | "master" | "theme";
  elements: NormalizedElement[];
  notes: NormalizedTextParagraph[];
  transition?: SlideTransitionInput;
  eventTimeline?: NormalizedSlideEventTimeline;
  readingOrder?: readonly SlideReadingOrder[];
  hidden: boolean;
  section?: string;
  tags: string[];
  customData: Record<string, JsonValue>;
}

export interface NormalizedPresentation {
  sections?: PresentationSection[];
  firstSlideNumber?: number;
  masters?: NormalizedSlideMaster[];
  irVersion: 2;
  id: string;
  metadata: NormalizedPresentationMetadata;
  size: PresentationSize;
  theme: NormalizedPresentationTheme;
  slides: NormalizedSlide[];
  assets: NormalizedAsset[];
  layouts: NormalizedSlideLayout[];
}

/** Read-only authoring data consumed by pure validation/normalization and edit planning. */
export type PresentationSlideData = Pick<PresentationSlide, "id" | "showInheritedShapes" | "layoutId" | "background" | "elements" | "notes" | "transition" | "eventTimeline" | "readingOrder" | "hidden" | "section" | "tags" | "customData">;
export type PresentationData = Pick<PresentationDocument, "id" | "metadata" | "size" | "theme" | "textStylePresets" | "assets" | "layouts" | "masters" | "firstSlideNumber" | "sections"> & { readonly slides: readonly PresentationSlideData[] };
