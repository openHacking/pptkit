import type { NormalizedAsset } from "@pptkit/core";
import type { LayoutResult } from "@pptkit/layout";
import { encodeUtf8 } from "../binary/encode.js";
import { REL } from "../constants/ooxml.js";
import { elementXml, imageXml } from "../ooxml/elements.js";
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

export async function buildPackage(
  title: string | undefined,
  layout: LayoutResult,
  assets: NormalizedAsset[],
  warnings: ExportWarning[],
  loadAsset: AssetLoader,
): Promise<ZipPart[]> {
  const parts: ZipPart[] = [];
  const media = new Map<string, PackagedMedia>();
  for (const asset of assets) {
    try {
      const loaded = await loadAsset(asset.source, asset.mimeType);
      const name = `ppt/media/${asset.id}${loaded.extension}`;
      media.set(asset.id, { name, loaded });
      parts.push({ name, data: loaded.data });
    } catch (error) {
      warnings.push({ code: "asset-read-failed", message: `Could not package image asset "${asset.id}": ${error instanceof Error ? error.message : String(error)}`, assetId: asset.id });
    }
  }

  for (let index = 0; index < layout.slides.length; index += 1) {
    const slide = layout.slides[index]!;
    const relationships: Relationship[] = [
      { id: "rId1", type: REL.slideLayout, target: "../slideLayouts/slideLayout1.xml" },
      { id: "rId2", type: REL.notesSlide, target: `../notesSlides/notesSlide${index + 1}.xml` },
    ];
    let nextRelationship = 3;
    const elements: string[] = [];
    slide.elements.forEach((element, elementIndex) => {
      if (element.type === "image") {
        const packaged = media.get(element.assetId);
        if (!packaged) {
          warnings.push({ code: "image-omitted", message: `Image asset "${element.assetId}" was omitted from the slide.`, slideId: slide.id, elementIndex, assetId: element.assetId });
          return;
        }
        const relationshipId = `rId${nextRelationship++}`;
        relationships.push({ id: relationshipId, type: REL.image, target: `../media/${mediaFilename(packaged.name)}` });
        elements.push(imageXml(element, relationshipId, elementIndex));
        return;
      }
      elements.push(elementXml(element, elementIndex, warnings, slide.id));
    });
    parts.push({ name: `ppt/slides/slide${index + 1}.xml`, data: encodeUtf8(slideXml(elements, slide.background)) });
    parts.push({ name: `ppt/slides/_rels/slide${index + 1}.xml.rels`, data: encodeUtf8(relationshipsXml(relationships)) });
    parts.push({ name: `ppt/notesSlides/notesSlide${index + 1}.xml`, data: encodeUtf8(notesSlideXml()) });
    parts.push({ name: `ppt/notesSlides/_rels/notesSlide${index + 1}.xml.rels`, data: encodeUtf8(relationshipsXml([
      { id: "rId1", type: REL.notesMaster, target: "../notesMasters/notesMaster1.xml" },
      { id: "rId2", type: REL.slide, target: `../slides/slide${index + 1}.xml` },
    ])) });
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
  parts.push(
    { name: "[Content_Types].xml", data: encodeUtf8(contentTypesXml(media, layout.slideCount)) },
    { name: "_rels/.rels", data: encodeUtf8(rootRelationshipsXml()) },
    { name: "ppt/presentation.xml", data: encodeUtf8(presentationXml(layout)) },
    { name: "ppt/_rels/presentation.xml.rels", data: encodeUtf8(relationshipsXml(presentationRelationships)) },
    { name: "ppt/presProps.xml", data: encodeUtf8(presentationPropertiesXml()) },
    { name: "ppt/viewProps.xml", data: encodeUtf8(viewPropertiesXml()) },
    { name: "ppt/tableStyles.xml", data: encodeUtf8(tableStylesXml()) },
    { name: "ppt/slideMasters/slideMaster1.xml", data: encodeUtf8(slideMasterXml()) },
    { name: "ppt/slideMasters/_rels/slideMaster1.xml.rels", data: encodeUtf8(relationshipsXml([
      { id: "rId1", type: REL.slideLayout, target: "../slideLayouts/slideLayout1.xml" },
      { id: "rId2", type: REL.theme, target: "../theme/theme1.xml" },
    ])) },
    { name: "ppt/slideLayouts/slideLayout1.xml", data: encodeUtf8(slideLayoutXml()) },
    { name: "ppt/slideLayouts/_rels/slideLayout1.xml.rels", data: encodeUtf8(relationshipsXml([
      { id: "rId1", type: REL.slideMaster, target: "../slideMasters/slideMaster1.xml" },
    ])) },
    { name: "ppt/theme/theme1.xml", data: encodeUtf8(themeXml()) },
    { name: "ppt/theme/theme2.xml", data: encodeUtf8(themeXml()) },
    { name: "ppt/notesMasters/notesMaster1.xml", data: encodeUtf8(notesMasterXml()) },
    { name: "ppt/notesMasters/_rels/notesMaster1.xml.rels", data: encodeUtf8(relationshipsXml([
      { id: "rId1", type: REL.theme, target: "../theme/theme2.xml" },
    ])) },
    { name: "docProps/core.xml", data: encodeUtf8(corePropertiesXml(title)) },
    { name: "docProps/app.xml", data: encodeUtf8(appPropertiesXml(layout.slideCount)) },
  );
  return parts;
}

function mediaFilename(value: string): string {
  return value.slice(value.lastIndexOf("/") + 1);
}
