import type { NormalizedPresentationTheme } from "@pptkit/core";

/** Capture the already-resolved Core themes; never solve inheritance here. */
export function ownerTheme(document: { masters?: readonly { id: string; theme: NormalizedPresentationTheme }[]; theme: NormalizedPresentationTheme; layouts: readonly { id: string; theme?: NormalizedPresentationTheme }[]; slides: readonly { id: string; layoutId: string }[] }) {
  const masters = new Map((document.masters ?? []).map(master => [master.id, master.theme]));
  const themes = new Map(document.layouts.map(layout => [layout.id, layout.theme ?? document.theme]));
  const slides = new Map(document.slides.map(slide => [slide.id, slide.layoutId]));
  return (owner: { slideId?: string; layoutId?: string; masterId?: string }) => {
    if (owner.masterId !== undefined) return masters.get(owner.masterId) ?? document.theme;
    const layoutId = owner.layoutId ?? (owner.slideId === undefined ? undefined : slides.get(owner.slideId));
    return layoutId === undefined ? document.theme : themes.get(layoutId) ?? document.theme;
  };
}
