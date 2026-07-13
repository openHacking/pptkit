import type { ElementAction, NormalizedPlaceholderDefinition } from "@pptkit/core";
import type { LayoutElement, LayoutResult } from "@pptkit/layout";
import { encodeUtf8 } from "../binary/encode.js";
import { REL } from "../constants/ooxml.js";
import { elementXml, type ElementXmlContext } from "../ooxml/elements.js";
import {
  appPropertiesXml,
  contentTypesXml,
  corePropertiesXml,
  notesMasterXml,
  notesSlideXml,
  presentationPropertiesXml,
  presentationXml,
  slideLayoutXml,
  slideMasterXml,
  slideXml,
  tableStylesXml,
  themeXml,
  viewPropertiesXml,
} from "../ooxml/package-parts.js";
import { relationshipsXml, rootRelationshipsXml } from "../ooxml/relationships.js";
import type { ExportWarning } from "../types/export.js";
import type { AssetLoader, PackagedMedia, Relationship, ZipPart } from "../types/internal.js";

interface PartContextResult {
  context: ElementXmlContext;
  relationships: Relationship[];
}

function mediaFilename(value: string): string {
  return value.slice(value.lastIndexOf("/") + 1);
}

function createPartContext(options: {
  baseRelationships: Relationship[];
  media: ReadonlyMap<string, PackagedMedia>;
  slideIndexes: ReadonlyMap<string, number>;
  placeholders?: NormalizedPlaceholderDefinition[];
  warnings: ExportWarning[];
  slideId?: string;
}): PartContextResult {
  const relationships = [...options.baseRelationships];
  let nextRelationship = relationships.length + 1;
  let nextObject = 2;
  const context: ElementXmlContext = {
    ...(options.slideId !== undefined ? { slideId: options.slideId } : {}),
    warnings: options.warnings,
    placeholders: new Map((options.placeholders ?? []).map((placeholder) => [placeholder.key, placeholder])),
    nextObjectId: () => nextObject++,
    imageRelationship(assetId) {
      const packaged = options.media.get(assetId);
      if (packaged === undefined) {
        options.warnings.push({ code: "image-omitted", message: `Image asset "${assetId}" was omitted from the part.`, ...(options.slideId !== undefined ? { slideId: options.slideId } : {}), assetId });
        return undefined;
      }
      const id = `rId${nextRelationship++}`;
      relationships.push({ id, type: REL.image, target: `../media/${mediaFilename(packaged.name)}` });
      return id;
    },
    actionRelationship(action: ElementAction) {
      const id = `rId${nextRelationship++}`;
      if (action.type === "url") {
        relationships.push({ id, type: REL.hyperlink, target: action.url, targetMode: "External" });
        return id;
      }
      const index = options.slideIndexes.get(action.slideId);
      if (index === undefined) return undefined;
      relationships.push({ id, type: REL.slide, target: `../slides/slide${index + 1}.xml` });
      return id;
    },
  };
  return { context, relationships };
}

function elementsXml(elements: readonly LayoutElement[], context: ElementXmlContext): string[] {
  return elements.map((element) => elementXml(element, context)).filter((value) => value !== "");
}

export async function buildPackage(
  layout: LayoutResult,
  warnings: ExportWarning[],
  loadAsset: AssetLoader,
): Promise<ZipPart[]> {
  const parts: ZipPart[] = [];
  const media = new Map<string, PackagedMedia>();
  for (const asset of layout.assets) {
    try {
      const loaded = await loadAsset(asset.source, asset.mimeType);
      const name = `ppt/media/${asset.id}${loaded.extension}`;
      media.set(asset.id, { name, loaded });
      parts.push({ name, data: loaded.data });
    } catch (error) {
      warnings.push({ code: "asset-read-failed", message: `Could not package image asset "${asset.id}": ${error instanceof Error ? error.message : String(error)}`, assetId: asset.id });
    }
  }

  const slideIndexes = new Map(layout.slides.map((slide, index) => [slide.id, index]));
  const layoutIndexes = new Map(layout.layouts.map((item, index) => [item.id, index]));

  for (let index = 0; index < layout.layouts.length; index += 1) {
    const item = layout.layouts[index]!;
    const partContext = createPartContext({
      baseRelationships: [{ id: "rId1", type: REL.slideMaster, target: "../slideMasters/slideMaster1.xml" }],
      media,
      slideIndexes,
      placeholders: item.placeholders,
      warnings,
    });
    const xml = elementsXml(item.elements, partContext.context);
    parts.push({ name: `ppt/slideLayouts/slideLayout${index + 1}.xml`, data: encodeUtf8(slideLayoutXml(item, xml, partContext.context)) });
    parts.push({ name: `ppt/slideLayouts/_rels/slideLayout${index + 1}.xml.rels`, data: encodeUtf8(relationshipsXml(partContext.relationships)) });
  }

  for (let index = 0; index < layout.slides.length; index += 1) {
    const slide = layout.slides[index]!;
    const layoutIndex = layoutIndexes.get(slide.layoutId)!;
    const slideLayout = layout.layouts[layoutIndex]!;
    const partContext = createPartContext({
      baseRelationships: [
        { id: "rId1", type: REL.slideLayout, target: `../slideLayouts/slideLayout${layoutIndex + 1}.xml` },
        { id: "rId2", type: REL.notesSlide, target: `../notesSlides/notesSlide${index + 1}.xml` },
      ],
      media,
      slideIndexes,
      placeholders: slideLayout.placeholders,
      warnings,
      slideId: slide.id,
    });
    const xml = elementsXml(slide.elements, partContext.context);
    parts.push({
      name: `ppt/slides/slide${index + 1}.xml`,
      data: encodeUtf8(slideXml(xml, slide.backgroundSource === "slide" ? slide.background : undefined, slide.hidden, {
        ...(slide.section !== undefined ? { section: slide.section } : {}),
        tags: slide.tags,
        customData: slide.customData,
      })),
    });
    parts.push({ name: `ppt/slides/_rels/slide${index + 1}.xml.rels`, data: encodeUtf8(relationshipsXml(partContext.relationships)) });

    const notesContext = createPartContext({
      baseRelationships: [
        { id: "rId1", type: REL.notesMaster, target: "../notesMasters/notesMaster1.xml" },
        { id: "rId2", type: REL.slide, target: `../slides/slide${index + 1}.xml` },
      ],
      media,
      slideIndexes,
      warnings,
      slideId: slide.id,
    });
    parts.push({ name: `ppt/notesSlides/notesSlide${index + 1}.xml`, data: encodeUtf8(notesSlideXml(slide.notes, notesContext.context)) });
    parts.push({ name: `ppt/notesSlides/_rels/notesSlide${index + 1}.xml.rels`, data: encodeUtf8(relationshipsXml(notesContext.relationships)) });
  }

  const presentationRelationships: Relationship[] = [
    { id: "rId1", type: REL.slideMaster, target: "slideMasters/slideMaster1.xml" },
  ];
  layout.slides.forEach((_, index) => presentationRelationships.push({ id: `rId${index + 2}`, type: REL.slide, target: `slides/slide${index + 1}.xml` }));
  presentationRelationships.push(
    { id: `rId${layout.slideCount + 2}`, type: REL.notesMaster, target: "notesMasters/notesMaster1.xml" },
    { id: `rId${layout.slideCount + 3}`, type: REL.presProps, target: "presProps.xml" },
    { id: `rId${layout.slideCount + 4}`, type: REL.viewProps, target: "viewProps.xml" },
    { id: `rId${layout.slideCount + 5}`, type: REL.theme, target: "theme/theme1.xml" },
    { id: `rId${layout.slideCount + 6}`, type: REL.tableStyles, target: "tableStyles.xml" },
  );
  const masterRelationships: Relationship[] = layout.layouts.map((_, index) => ({ id: `rId${index + 1}`, type: REL.slideLayout, target: `../slideLayouts/slideLayout${index + 1}.xml` }));
  masterRelationships.push({ id: `rId${layout.layouts.length + 1}`, type: REL.theme, target: "../theme/theme1.xml" });

  parts.push(
    { name: "[Content_Types].xml", data: encodeUtf8(contentTypesXml(media, layout.slideCount, layout.layouts.length)) },
    { name: "_rels/.rels", data: encodeUtf8(rootRelationshipsXml()) },
    { name: "ppt/presentation.xml", data: encodeUtf8(presentationXml(layout)) },
    { name: "ppt/_rels/presentation.xml.rels", data: encodeUtf8(relationshipsXml(presentationRelationships)) },
    { name: "ppt/presProps.xml", data: encodeUtf8(presentationPropertiesXml()) },
    { name: "ppt/viewProps.xml", data: encodeUtf8(viewPropertiesXml()) },
    { name: "ppt/tableStyles.xml", data: encodeUtf8(tableStylesXml()) },
    { name: "ppt/slideMasters/slideMaster1.xml", data: encodeUtf8(slideMasterXml(layout.layouts)) },
    { name: "ppt/slideMasters/_rels/slideMaster1.xml.rels", data: encodeUtf8(relationshipsXml(masterRelationships)) },
    { name: "ppt/theme/theme1.xml", data: encodeUtf8(themeXml(layout.theme)) },
    { name: "ppt/theme/theme2.xml", data: encodeUtf8(themeXml(layout.theme)) },
    { name: "ppt/notesMasters/notesMaster1.xml", data: encodeUtf8(notesMasterXml()) },
    { name: "ppt/notesMasters/_rels/notesMaster1.xml.rels", data: encodeUtf8(relationshipsXml([
      { id: "rId1", type: REL.theme, target: "../theme/theme2.xml" },
    ])) },
    { name: "docProps/core.xml", data: encodeUtf8(corePropertiesXml(layout.metadata)) },
    { name: "docProps/app.xml", data: encodeUtf8(appPropertiesXml(layout.slideCount, layout.metadata)) },
  );
  return parts;
}
