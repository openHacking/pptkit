import { createPresentation, resolvePresentationInheritance, type ManagedPresentationSlide } from "@pptkit/core";
import { resolveLayout, type LayoutSlideMaster } from "@pptkit/layout";

export function sharedMasterArtwork() {
  const document = createPresentation({ masters: [{ id: "brand", name: "Brand", elements: [{ id: "brand-rule", type: "shape", shape: "rect", box: { x: 20, y: 20, width: 300, height: 8 }, style: { fill: { type: "solid", color: { theme: "accent1" } } } }] }] });
  document.defineSlideLayout({ id: "body", name: "Body", masterId: "brand", showMasterShapes: true });
  const slide: ManagedPresentationSlide = document.addSlide({ layoutId: "body", showInheritedShapes: false });
  slide.setShowInheritedShapes();
  const layout = resolveLayout(document);
  const master: LayoutSlideMaster | undefined = layout.masters?.[0];
  return { document, master, resolution: resolvePresentationInheritance(document) };
}
