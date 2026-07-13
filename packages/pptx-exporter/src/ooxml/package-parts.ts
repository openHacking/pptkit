import type { NormalizedPaint, NormalizedPresentationMetadata, NormalizedPresentationTheme, NormalizedTextParagraph } from "@pptkit/core";
import type { LayoutResult, LayoutSlideLayout } from "@pptkit/layout";
import { CONTENT_TYPES } from "../constants/ooxml.js";
import type { PackagedMedia } from "../types/internal.js";
import type { ElementXmlContext } from "./elements.js";
import { placeholderShapeXml, textParagraphsXml } from "./elements.js";
import { colorValue, emu, escapeXml, groupTransformXml, paintXml } from "./xml.js";

function backgroundXml(background: NormalizedPaint | undefined): string {
  return background === undefined ? "" : `<p:bg><p:bgPr>${paintXml(background)}<a:effectLst/></p:bgPr></p:bg>`;
}

export function slideXml(elements: string[], background: NormalizedPaint | undefined, hidden: boolean, metadata?: { section?: string; tags: string[]; customData: Record<string, unknown> }): string {
  const hasMetadata = metadata !== undefined && (metadata.section !== undefined || metadata.tags.length > 0 || Object.keys(metadata.customData).length > 0);
  const extension = hasMetadata ? `<p:extLst><p:ext uri="{7F4F1C93-4B18-4D30-91CE-PPTKIT000001}"><pptkit:slideData>${escapeXml(JSON.stringify(metadata))}</pptkit:slideData></p:ext></p:extLst>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:pptkit="https://pptkit.dev/ooxml/2026/main"${hidden ? ' show="0"' : ""}><p:cSld>${backgroundXml(background)}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml()}</p:grpSpPr>${elements.join("")}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${extension}</p:sld>`;
}

export function presentationXml(layout: LayoutResult): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1" autoCompressPictures="0"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId${layout.slideCount + 2}"/></p:notesMasterIdLst><p:sldIdLst>${layout.slides.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`).join("")}</p:sldIdLst><p:sldSz cx="${emu(layout.size.width)}" cy="${emu(layout.size.height)}"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="${escapeXml(layout.metadata.language)}"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800"/></a:lvl1pPr></p:defaultTextStyle></p:presentation>`;
}

export function contentTypesXml(media: ReadonlyMap<string, PackagedMedia>, slideCount: number, layoutCount: number): string {
  const defaults = [
    `<Default Extension="rels" ContentType="${CONTENT_TYPES.rels}"/>`,
    `<Default Extension="xml" ContentType="application/xml"/>`,
  ];
  const seen = new Set(["rels", "xml"]);
  for (const item of media.values()) {
    const extension = item.loaded.extension.slice(1);
    if (!seen.has(extension)) {
      seen.add(extension);
      defaults.push(`<Default Extension="${extension}" ContentType="${item.loaded.mimeType}"/>`);
    }
  }
  const overrides = [
    `<Override PartName="/ppt/presentation.xml" ContentType="${CONTENT_TYPES.presentation}"/>`,
    `<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="${CONTENT_TYPES.notesMaster}"/>`,
    `<Override PartName="/ppt/presProps.xml" ContentType="${CONTENT_TYPES.presProps}"/>`,
    `<Override PartName="/ppt/viewProps.xml" ContentType="${CONTENT_TYPES.viewProps}"/>`,
    `<Override PartName="/ppt/tableStyles.xml" ContentType="${CONTENT_TYPES.tableStyles}"/>`,
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${CONTENT_TYPES.slideMaster}"/>`,
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="${CONTENT_TYPES.theme}"/>`,
    `<Override PartName="/ppt/theme/theme2.xml" ContentType="${CONTENT_TYPES.theme}"/>`,
    `<Override PartName="/docProps/core.xml" ContentType="${CONTENT_TYPES.core}"/>`,
    `<Override PartName="/docProps/app.xml" ContentType="${CONTENT_TYPES.app}"/>`,
  ];
  for (let index = 1; index <= layoutCount; index += 1) overrides.push(`<Override PartName="/ppt/slideLayouts/slideLayout${index}.xml" ContentType="${CONTENT_TYPES.slideLayout}"/>`);
  for (let index = 1; index <= slideCount; index += 1) {
    overrides.push(`<Override PartName="/ppt/slides/slide${index}.xml" ContentType="${CONTENT_TYPES.slide}"/>`);
    overrides.push(`<Override PartName="/ppt/notesSlides/notesSlide${index}.xml" ContentType="${CONTENT_TYPES.notesSlide}"/>`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults.join("")}${overrides.join("")}</Types>`;
}

export function slideMasterXml(layouts: LayoutSlideLayout[]): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="Master"><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml()}</p:grpSpPr></p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:sldLayoutIdLst>${layouts.map((_, index) => `<p:sldLayoutId id="${2147483649 + index}" r:id="rId${index + 1}"/>`).join("")}</p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr lang="en-US" sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr lang="en-US" sz="2800"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:otherStyle></p:txStyles></p:sldMaster>`;
}

export function notesMasterXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml()}</p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr><a:defRPr lang="en-US"/></a:lvl1pPr></p:notesStyle></p:notesMaster>`;
}

export function notesSlideXml(notes: NormalizedTextParagraph[], context: ElementXmlContext): string {
  const hasContent = notes.some((paragraph) => paragraph.runs.some((run) => run.text.length > 0));
  const body = hasContent ? `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Text"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${textParagraphsXml(notes, context)}</p:txBody></p:sp>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml()}</p:grpSpPr>${body}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
}

export function presentationPropertiesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`;
}

export function viewPropertiesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr horzBarState="maximized"><p:restoredLeft sz="15611"/><p:restoredTop sz="94610"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0" snapToObjects="1"><p:cViewPr varScale="1"><p:scale><a:sx n="136" d="100"/><a:sy n="136" d="100"/></p:scale><p:origin x="216" y="312"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`;
}

export function tableStylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
}

export function slideLayoutXml(layout: LayoutSlideLayout, elements: string[], context: ElementXmlContext): string {
  const placeholders = layout.placeholders.map((placeholder, index) => placeholderShapeXml(placeholder, index, context)).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"><p:cSld name="${escapeXml(layout.name)}">${backgroundXml(layout.background)}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml()}</p:grpSpPr>${elements.join("")}${placeholders}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function themeEntry(tag: string, color: string): string {
  return `<a:${tag}><a:srgbClr val="${colorValue(color)}"/></a:${tag}>`;
}

export function themeXml(theme: NormalizedPresentationTheme): string {
  const c = theme.colors;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${escapeXml(theme.name)}"><a:themeElements><a:clrScheme name="${escapeXml(theme.name)}">${themeEntry("dk1", c.text1)}${themeEntry("lt1", c.background1)}${themeEntry("dk2", c.text2)}${themeEntry("lt2", c.background2)}${themeEntry("accent1", c.accent1)}${themeEntry("accent2", c.accent2)}${themeEntry("accent3", c.accent3)}${themeEntry("accent4", c.accent4)}${themeEntry("accent5", c.accent5)}${themeEntry("accent6", c.accent6)}${themeEntry("hlink", c.hyperlink)}${themeEntry("folHlink", c.followedHyperlink)}</a:clrScheme><a:fontScheme name="${escapeXml(theme.name)}"><a:majorFont><a:latin typeface="${escapeXml(theme.fonts.heading)}"/><a:ea typeface="${escapeXml(theme.fonts.heading)}"/><a:cs typeface="${escapeXml(theme.fonts.heading)}"/></a:majorFont><a:minorFont><a:latin typeface="${escapeXml(theme.fonts.body)}"/><a:ea typeface="${escapeXml(theme.fonts.body)}"/><a:cs typeface="${escapeXml(theme.fonts.body)}"/></a:minorFont></a:fontScheme><a:fmtScheme name="${escapeXml(theme.name)}"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="28575"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

export function corePropertiesXml(metadata: NormalizedPresentationMetadata): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:cpTerms="http://purl.org/dc/terms/"><dc:title>${escapeXml(metadata.title)}</dc:title><dc:creator>${escapeXml(metadata.author)}</dc:creator>${metadata.subject === undefined ? "" : `<dc:subject>${escapeXml(metadata.subject)}</dc:subject>`}${metadata.description === undefined ? "" : `<dc:description>${escapeXml(metadata.description)}</dc:description>`}<cp:keywords>${escapeXml(metadata.keywords.join(", "))}</cp:keywords><cp:revision>${metadata.revision}</cp:revision><dc:language>${escapeXml(metadata.language)}</dc:language></cp:coreProperties>`;
}

export function appPropertiesXml(slideCount: number, metadata: NormalizedPresentationMetadata): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>PPTKit</Application>${metadata.company === undefined ? "" : `<Company>${escapeXml(metadata.company)}</Company>`}<Slides>${slideCount}</Slides><PresentationFormat>On-screen Show (16:9)</PresentationFormat></Properties>`;
}
