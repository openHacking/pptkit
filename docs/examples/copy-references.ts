import { createPresentation, validatePresentation } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";

export async function copyConnectedSlide() {
  const document = createPresentation();
  const slide = document.addSlide({ id: "diagram" });
  slide.addElement({ type: "shape", shape: "rect", id: "left",
    box: { x: 20, y: 20, width: 100, height: 60 }, action: { type: "slide", slideId: slide.id } });
  slide.addElement({ type: "shape", shape: "rect", id: "right",
    box: { x: 240, y: 20, width: 100, height: 60 } });
  slide.addElement({ type: "connector", start: { elementId: "left", anchor: "right" },
    end: { elementId: "right", anchor: "left" } });
  const copy = document.duplicateSlide(slide.id);
  document.removeSlide(slide.id);
  // The copy's connection endpoints and self-link now use its own identities.
  return { copy, diagnostics: validatePresentation(document), output: await generatePptx(document) };
}

export async function copySharedTemplate() {
  const source = createPresentation(), destination = createPresentation();
  source.defineSlideMaster({ id: "source-master", name: "Master" });
  source.defineSlideLayout({ id: "source-layout", name: "Layout", masterId: "source-master" });
  source.addSlide({ id: "source-slide", layoutId: "source-layout", elements: [
    { type: "text", id: "label", box: { x: 20, y: 20, width: 200, height: 40 }, content: "Copied page" },
  ] });
  const assets = { deduplicate: "content", external: "reject" } as const;
  for (const suffix of ["one", "two"]) {
    const master = destination.copySlideMasterFrom(source, "source-master", {
      id: `master-${suffix}`, theme: "source-effective", assets, reuseEquivalent: true,
    });
    const layout = destination.copySlideLayoutFrom(source, "source-layout", {
      id: `layout-${suffix}`, masterId: master.master.id, theme: "destination", assets, reuseEquivalent: true,
    });
    destination.copySlideFrom(source, "source-slide", {
      layoutId: layout.layout.id, theme: "destination", geometry: "preserve", customData: "reject", assets,
    });
  }
  return { document: destination, diagnostics: validatePresentation(destination), output: await generatePptx(destination) };
}
