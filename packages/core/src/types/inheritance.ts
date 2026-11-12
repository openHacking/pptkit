import type { ThemeColorRole } from "./style.js";
import type { NormalizedPresentation } from "./presentation.js";

export interface PresentationInheritanceTarget {
  readonly kind: "document" | "master" | "layout" | "slide";
  readonly id: string;
}

export type PresentationInheritanceTextTarget =
  | { readonly kind: "element"; readonly id: string }
  | { readonly kind: "placeholder"; readonly id: string; readonly key: string }
  | { readonly kind: "notes"; readonly id: string };

export type PresentationInheritanceSource =
  | { readonly kind: "default"; readonly path: string }
  | { readonly kind: "preset"; readonly id: string; readonly path: string }
  | ((PresentationInheritanceTarget | PresentationInheritanceTextTarget) & { readonly path: string });

/** Logical effective property; slide/layout theme entries describe their owning theme. */
export type PresentationInheritanceProperty =
  | "theme.name" | `theme.colors.${ThemeColorRole}` | "theme.fonts.heading" | "theme.fonts.body"
  | `background.stops.${number}.${"position" | "color" | "opacity"}`
  | "background" | "background.type" | "background.color" | "background.opacity" | "background.angle" | "background.scaled" | "background.stops"
  | `frame.${string}` | `content.${number}.style.${string}` | `content.${number}.runs.${number}.style.${string}`
  | `text.frame.${string}` | `text.content.${number}.style.${string}` | `text.content.${number}.runs.${number}.style.${string}`
  | `rows.${number}.cells.${number}.style.${string}`
  | `rows.${number}.cells.${number}.content.${number}.style.${string}`
  | `rows.${number}.cells.${number}.content.${number}.runs.${number}.style.${string}`
  | `notes.${number}.style.${string}` | `notes.${number}.runs.${number}.style.${string}`
  | `style.fill.${string}` | `style.stroke.${string}`
  | `style.textStyle.${string}` | `titleStyle.${string}` | `legend.textStyle.${string}`
  | `axes.category.labelStyle.${string}` | `axes.value.labelStyle.${string}`
  | `textStyle.frame.${string}` | `textStyle.paragraph.${string}` | `textStyle.run.${string}`;

export interface PresentationInheritanceOrigin {
  readonly target: PresentationInheritanceTarget | PresentationInheritanceTextTarget;
  readonly property: PresentationInheritanceProperty;
  readonly source: PresentationInheritanceSource;
  /** An inherited optional font was explicitly removed, rather than never declared. */
  readonly state?: "cleared";
}

export interface PresentationInheritanceResolution {
  /** Ordinary detached IR. Editing this snapshot does not recompute its origin report. */
  readonly presentation: NormalizedPresentation;
  /** Frozen snapshot of supported authoring declaration origins, not source-format import provenance. */
  readonly origins: readonly PresentationInheritanceOrigin[];
}
