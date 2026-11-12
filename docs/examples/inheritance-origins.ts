import { createPresentation, resolvePresentationInheritance, type PresentationInheritanceOrigin } from "@pptkit/core";
import { resolveNormalizedLayout } from "@pptkit/layout";

export function authoringOrigins() {
  const document = createPresentation({ masters: [{ id: "brand", name: "Brand", theme: { colors: { accent1: "118855" } } }] });
  document.defineSlideLayout({ id: "body", name: "Body", masterId: "brand", background: { type: "solid", color: { theme: "accent1" }, opacity: 0 } });
  document.addSlide({ id: "one", layoutId: "body" });
  const result = resolvePresentationInheritance(document);
  const accent: PresentationInheritanceOrigin | undefined = result.origins.find(row => row.target.kind === "slide" && row.target.id === "one" && row.property === "theme.colors.accent1");
  return { result, accent, layout: resolveNormalizedLayout(result.presentation) };
}
