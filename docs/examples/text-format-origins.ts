import { createPresentation, resolvePresentationInheritance, type PresentationInheritanceTextTarget } from "@pptkit/core";

export function formatOrigins() {
  const document = createPresentation({ textStylePresets: { "brand.body": { run: { bold: true, eastAsianFontFamily: "Arial" } } } });
  document.defineSlideLayout({ id: "body", name: "Body", placeholders: [{ key: "body", kind: "body", box: { x: 10, y: 20, width: 300, height: 80 }, textStylePreset: "brand.body" }] });
  document.addSlide({ layoutId: "body", elements: [{ id: "text", type: "text", placeholderKey: "body", content: [{ runs: [{ text: "", style: { bold: false, eastAsianFontFamily: null } }] }] }] });
  document.slides[0]!.setNotes([{ runs: [{ text: "", style: { complexScriptFontFamily: null } }] }]);
  const result = resolvePresentationInheritance(document);
  const target: PresentationInheritanceTextTarget = { kind: "element", id: "text" };
  return { target, result, notes: result.origins.filter(row => row.target.kind === "notes"), clears: result.origins.filter(row => row.target.kind === "element" && row.target.id === "text" && row.state === "cleared") };
}
