import { createPresentation, normalizePresentation, type SlideMasterInput } from "@pptkit/core";
import { generatePptx } from "@pptkit/pptx-exporter";
import { renderPresentationToSvg } from "@pptkit/svg-renderer";
export async function multipleMasters() {
  const masters: SlideMasterInput[] = [
    { id: "warm", name: "Warm", theme: { colors: { accent1: "CC2200" } } },
    { id: "cool", name: "Cool", theme: { colors: { accent1: "0044CC" } } },
  ];
  const document = createPresentation({ masters });
  for (const master of masters) {
    document.defineSlideLayout({ id: master.id + "-layout", name: master.name, masterId: master.id });
    document.addSlide({ layoutId: master.id + "-layout" }).addElement({ type: "text", box: { x: 20, y: 20, width: 300, height: 60 }, content: [{ runs: [{ text: master.name, style: { color: { theme: "accent1" } } }] }] });
  }
  document.replaceSlideMaster("warm", { name: "Warm updated", theme: { colors: { accent1: "118855" } } });
  document.defineSlideMaster({ id: "alternate", name: "Alternate" });
  document.replaceSlideLayout("cool-layout", { name: "Cool now inherits", masterId: "alternate" });
  return { normalized: normalizePresentation(document), pptx: await generatePptx(document), svg: await renderPresentationToSvg(document) };
}
