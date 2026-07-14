import { createPresentation, type NormalizedAsset } from "@pptkit/core";
import { renderPresentationToSvg } from "@pptkit/svg-renderer";

function readAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(String(reader.result)));
    reader.addEventListener("error", () => reject(reader.error ?? new Error("Could not read image.")));
    reader.readAsDataURL(file);
  });
}

export async function previewImage(file: File, container: HTMLElement) {
  const presentation = createPresentation({ metadata: { title: "Browser preview" } });
  const image = presentation.registerAsset({
    kind: "image",
    source: { type: "path", value: file.name },
    mimeType: file.type,
    accessibility: { description: file.name },
  });
  presentation.addSlide({
    elements: [{
      type: "image",
      assetId: image.id,
      fit: "contain",
      box: { x: 80, y: 60, width: 800, height: 420 },
    }],
  });

  const resolveAsset = async (asset: NormalizedAsset) =>
    asset.id === image.id ? readAsDataUrl(file) : undefined;
  const result = await renderPresentationToSvg(presentation, { resolveAsset });
  container.innerHTML = result.slides[0]?.svg ?? "";
  return result;
}
