import type { PresentationSize, NormalizedBackgroundPaint, NormalizedPresentationMetadata, NormalizedPresentationTheme, NormalizedTextParagraph, SlideTransitionInput } from "@pptkit/core";
import type { LayoutResult, LayoutSlideLayout } from "@pptkit/layout";
import { CONTENT_TYPES } from "../constants/ooxml.js";
import type { PackagedMedia } from "../types/internal.js";
import type { ElementXmlContext } from "./elements.js";
import { placeholderShapeXml, textParagraphsXml } from "./elements.js";
import { backgroundPaintXml, colorValue, emu, escapeXml, groupTransformXml } from "./xml.js";

function backgroundXml(background: NormalizedBackgroundPaint | undefined): string {
  return background === undefined ? "" : `<p:bg><p:bgPr>${backgroundPaintXml(background)}<a:effectLst/></p:bgPr></p:bg>`;
}

function transitionXml(transition: SlideTransitionInput | undefined): string {
  if (transition === undefined) return "";
  const speed = transition.durationMs <= 750 ? "fast" : transition.durationMs >= 1500 ? "slow" : "med";
  const attributes = ` spd="${speed}"${transition.onClick === false ? ' advClick="0"' : ""}${transition.afterMs !== undefined ? ` advTm="${transition.afterMs}"` : ""}`;
  const effect = transition.type === "fade" ? "<p:fade/>" : `<p:${transition.type} dir="${{ left: "l", right: "r", up: "u", down: "d" }[transition.direction]}"/>`;
  return `<p:transition xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" mc:Ignorable="p14" p14:dur="${transition.durationMs}"${attributes}>${effect}</p:transition>`;
}

type MediaTiming = { kind: "audio" | "video"; objectId: number; durationMs: number; playbackStart?: "click-sequence" | "object-click" | "automatic"; playbackVolume?: "low" | "medium" | "high" | "mute"; loop?: boolean; rewindAfterPlayback?: boolean };

function audioVolumeAttributes(volume: MediaTiming["playbackVolume"]): string {
  if (volume === "low") return ' vol="20000"';
  if (volume === "medium") return "";
  if (volume === "mute") return ' mute="1"';
  return ' vol="80000"';
}

function audioTimeAttributes(loop: boolean | undefined, rewind: boolean | undefined): string {
  return `${loop ? ' repeatCount="indefinite"' : ""} fill="${rewind ? "remove" : "hold"}" display="0"`;
}

/** Native mixed timeline with independently triggered audio objects. */
function objectClickMediaTimingXml(timings: readonly MediaTiming[]): string {
  const mainItems = timings.filter(item => item.playbackStart !== "object-click");
  let nextId = mainItems.length ? 3 : 2;
  const mainDefault = () => mainItems.map(({ objectId, durationMs }) => {
    const first = nextId; nextId += 4;
    return `<p:par><p:cTn id="${first}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${first + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${first + 2}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${first + 3}" dur="${durationMs}" fill="hold"/><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
  }).join("");
  const mainAutomatic = () => {
    const groups: MediaTiming[][] = [];
    for (const item of mainItems) {
      if (item.playbackStart === "automatic" && groups.length) groups[groups.length - 1]!.push(item);
      else groups.push([item]);
    }
    return groups.map((group, groupIndex) => {
      const first = group[0]!, groupId = nextId++, triggerId = nextId++;
      const automaticFirst = first.playbackStart === "automatic";
      let elapsed = 0;
      const effects = group.map((item, index) => {
        if (index && group[index - 1]!.loop) throw new TypeError("PPTX audio export: automatic audio cannot follow looping audio in the main sequence.");
        if (elapsed > 2147483647) throw new RangeError("PPTX media automatic sequence delay exceeds 2147483647 milliseconds.");
        const wrapper = index ? nextId++ : undefined, effectId = nextId++, commandId = nextId++;
        const effect = `<p:par><p:cTn id="${effectId}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="${item.playbackStart === "automatic" ? "afterEffect" : "clickEffect"}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${commandId}" dur="${item.durationMs}" fill="hold"/><p:tgtEl><p:spTgt spid="${item.objectId}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par>`;
        const inner = index ? `<p:par><p:cTn id="${wrapper}" fill="hold"><p:stCondLst><p:cond delay="${elapsed}"/></p:stCondLst><p:childTnLst>${effect}</p:childTnLst></p:cTn></p:par>` : `<p:par><p:cTn id="${triggerId}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effect}</p:childTnLst></p:cTn></p:par>`;
        elapsed += item.durationMs;
        return inner;
      }).join("");
      const start = automaticFirst && groupIndex === 0 ? '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' : "";
      return `<p:par><p:cTn id="${groupId}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>${start}</p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par>`;
    }).join("");
  };
  const main = mainItems.some(item => item.playbackStart === "automatic") ? mainAutomatic() : mainDefault();
  const interaction = (item: MediaTiming, pause: boolean) => {
    const first = nextId; nextId += 5;
    const { objectId, durationMs } = item;
    return `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${first}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cond></p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst><p:par><p:cTn id="${first + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${first + 2}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${first + 3}" presetID="${pause ? 2 : 1}" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="${pause ? "togglePause" : "playFrom(0.0)"}"><p:cBhvr><p:cTn id="${first + 4}" dur="${pause ? 1 : durationMs}" fill="hold"/><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`;
  };
  const media = (item: MediaTiming) => {
    const end = item.kind === "audio" ? '<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>' : "";
    return `<p:${item.kind}><p:cMediaNode${item.kind === "audio" ? audioVolumeAttributes(item.playbackVolume) : ' vol="80000"'}><p:cTn id="${nextId++}"${item.kind === "audio" ? audioTimeAttributes(item.loop, item.rewindAfterPlayback) : ' fill="hold" display="0"'}><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>${end}</p:cTn><p:tgtEl><p:spTgt spid="${item.objectId}"/></p:tgtEl></p:cMediaNode></p:${item.kind}>`;
  };
  const videos = timings.filter(item => item.kind === "video"), audios = timings.filter(item => item.kind === "audio");
  let tail = videos.length ? media(videos[0]!) : "";
  for (const video of videos) tail += interaction(video, true);
  for (const video of videos.slice(1)) tail += media(video);
  for (const audio of audios) { if (audio.playbackStart === "object-click") tail += interaction(audio, false); tail += media(audio); }
  const mainSequence = mainItems.length ? `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${main}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>` : "";
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>${mainSequence}${tail}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
}

function audioTimingXml(timings: readonly { objectId: number; durationMs: number; playbackVolume?: MediaTiming["playbackVolume"]; loop?: boolean; rewindAfterPlayback?: boolean }[] | undefined): string {
  if (!timings?.length) return "";
  const effects = timings.map(({ objectId, durationMs }, index) => {
    const first = 3 + index * 4;
    return `<p:par><p:cTn id="${first}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${first + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${first + 2}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${first + 3}" dur="${durationMs}" fill="hold"/><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
  }).join("");
  const media = timings.map(({ objectId, playbackVolume, loop, rewindAfterPlayback }, index) => `<p:audio><p:cMediaNode${audioVolumeAttributes(playbackVolume)}><p:cTn id="${3 + timings.length * 4 + index}"${audioTimeAttributes(loop, rewindAfterPlayback)}><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cMediaNode></p:audio>`).join("");
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${effects}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>${media}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
}

function videoTimingXml(timings: readonly { objectId: number; durationMs: number; kind?: "audio" | "video"; playbackVolume?: MediaTiming["playbackVolume"]; loop?: boolean; rewindAfterPlayback?: boolean }[] | undefined): string {
  if (!timings?.length) return "";
  const count = timings.length;
  const videos = timings.filter(item => item.kind !== "audio"), audios = timings.filter(item => item.kind === "audio");
  const effects = timings.map(({ objectId, durationMs }, index) => {
    const first = 3 + index * 4;
    return `<p:par><p:cTn id="${first}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${first + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${first + 2}" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="${first + 3}" dur="${durationMs}" fill="hold"/><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
  }).join("");
  const video = (objectId: number, id: number) => `<p:video><p:cMediaNode vol="80000"><p:cTn id="${id}" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cMediaNode></p:video>`;
  const interactions = videos.map(({ objectId }, index) => {
    const first = 4 + count * 4 + index * 5;
    return `<p:seq concurrent="1" nextAc="seek"><p:cTn id="${first}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cond></p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst><p:par><p:cTn id="${first + 1}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${first + 2}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${first + 3}" presetID="2" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst><p:cmd type="call" cmd="togglePause"><p:cBhvr><p:cTn id="${first + 4}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:nextCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`;
  }).join("");
  const remainingVideos = videos.slice(1).map(({ objectId }, index) => video(objectId, 4 + count * 4 + videos.length * 5 + index)).join("");
  const audio = audios.map(({ objectId, playbackVolume, loop, rewindAfterPlayback }, index) => `<p:audio><p:cMediaNode${audioVolumeAttributes(playbackVolume)}><p:cTn id="${3 + count * 4 + videos.length * 6 + index}"${audioTimeAttributes(loop, rewindAfterPlayback)}><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn><p:tgtEl><p:spTgt spid="${objectId}"/></p:tgtEl></p:cMediaNode></p:audio>`).join("");
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${effects}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>${video(videos[0]!.objectId, 3 + count * 4)}${interactions}${remainingVideos}${audio}</p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
}

export function slideXml(elements: string[], background: NormalizedBackgroundPaint | undefined, hidden: boolean, metadata?: { section?: string; tags: string[]; customData: Record<string, unknown>; showInheritedShapes?: boolean; transition?: SlideTransitionInput; audioTiming?: readonly { objectId: number; durationMs: number; playbackStart?: "click-sequence" | "object-click" | "automatic"; playbackVolume?: "low" | "medium" | "high" | "mute"; loop?: boolean; rewindAfterPlayback?: boolean }[]; videoTiming?: readonly { objectId: number; durationMs: number }[]; mixedMediaTiming?: readonly MediaTiming[] }): string {
  const hasMetadata = metadata !== undefined && (metadata.section !== undefined || metadata.tags.length > 0 || Object.keys(metadata.customData).length > 0);
  const extension = hasMetadata ? `<p:extLst><p:ext uri="{7F4F1C93-4B18-4D30-91CE-PPTKIT000001}"><pptkit:slideData>${escapeXml(JSON.stringify({ ...(metadata.section !== undefined ? { section: metadata.section } : {}), tags: metadata.tags, customData: metadata.customData }))}</pptkit:slideData></p:ext></p:extLst>` : "";
  const allMedia = metadata?.mixedMediaTiming ?? metadata?.audioTiming?.map(item => ({ ...item, kind: "audio" as const }));
  const timing = allMedia?.some(item => item.playbackStart === "object-click" || item.playbackStart === "automatic") ? objectClickMediaTimingXml(allMedia) : audioTimingXml(metadata?.audioTiming) + videoTimingXml(metadata?.mixedMediaTiming ?? metadata?.videoTiming);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:pptkit="https://pptkit.dev/ooxml/2026/main"${hidden ? ' show="0"' : ""}${metadata?.showInheritedShapes === undefined ? "" : ` showMasterSp="${metadata.showInheritedShapes ? 1 : 0}"`}><p:cSld>${backgroundXml(background)}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml()}</p:grpSpPr>${elements.join("")}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${transitionXml(metadata?.transition)}${timing}${extension}</p:sld>`;
}

export function presentationXml(layout: LayoutResult, firstSlideNumber?: number): string {
  // PowerPoint repairs cross-kind master/layout ID collisions. Keep the legacy
  // default master, then place explicit masters beyond the whole layout range.
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1" autoCompressPictures="0"${firstSlideNumber === undefined ? "" : ` firstSlideNum="${firstSlideNumber}"`}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/>${(layout.masters ?? []).map((_, i) => `<p:sldMasterId id="${2147483649 + layout.layouts.length + i}" r:id="rId${layout.slideCount + 7 + i}"/>`).join("")}</p:sldMasterIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId${layout.slideCount + 2}"/></p:notesMasterIdLst><p:sldIdLst>${layout.slides.map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`).join("")}</p:sldIdLst><p:sldSz cx="${emu(layout.size.width)}" cy="${emu(layout.size.height)}"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="${escapeXml(layout.metadata.language)}"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800"/></a:lvl1pPr></p:defaultTextStyle></p:presentation>`;
}

export function contentTypesXml(media: ReadonlyMap<string, PackagedMedia>, slideCount: number, layoutCount: number, chartCount = 0, extraMasters = 0, customProperties = false): string {
  const defaults = [
    `<Default Extension="rels" ContentType="${CONTENT_TYPES.rels}"/>`,
    `<Default Extension="xml" ContentType="application/xml"/>`,
  ];
  const seen = new Set(["rels", "xml"]);
  if (chartCount > 0) {
    seen.add("xlsx");
    defaults.push('<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>');
  }
  for (const item of media.values()) {
    if (item.svgName && !seen.has("svg")) {
      seen.add("svg");
      defaults.push('<Default Extension="svg" ContentType="image/svg+xml"/>');
    }
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
  if (customProperties) overrides.push('<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>');
  for (let index = 0; index < extraMasters; index++) {
    overrides.push(`<Override PartName="/ppt/slideMasters/slideMaster${index + 2}.xml" ContentType="${CONTENT_TYPES.slideMaster}"/>`);
    overrides.push(`<Override PartName="/ppt/theme/theme${index + 3}.xml" ContentType="${CONTENT_TYPES.theme}"/>`);
  }
  for (let index = 1; index <= layoutCount; index += 1) overrides.push(`<Override PartName="/ppt/slideLayouts/slideLayout${index}.xml" ContentType="${CONTENT_TYPES.slideLayout}"/>`);
  for (let index = 1; index <= slideCount; index += 1) {
    overrides.push(`<Override PartName="/ppt/slides/slide${index}.xml" ContentType="${CONTENT_TYPES.slide}"/>`);
    overrides.push(`<Override PartName="/ppt/notesSlides/notesSlide${index}.xml" ContentType="${CONTENT_TYPES.notesSlide}"/>`);
  }
  for (let index = 1; index <= chartCount; index += 1) overrides.push(`<Override PartName="/ppt/charts/chart${index}.xml" ContentType="${CONTENT_TYPES.chart}"/>`);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults.join("")}${overrides.join("")}</Types>`;
}

export function slideMasterXml(layouts: LayoutSlideLayout[], master?: { name?: string; background?: NormalizedBackgroundPaint; layoutIndexes: number[]; preserve?: boolean }, elements: string[] = []): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${master?.preserve ? ' preserve="1"' : ""}><p:cSld name="${escapeXml(master?.name ?? "Master")}">${master?.background === undefined ? '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' : backgroundXml(master.background)}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml()}</p:grpSpPr>${elements.join("")}</p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/>${layouts.length ? `<p:sldLayoutIdLst>${layouts.map((_, index) => `<p:sldLayoutId id="${2147483649 + (master?.layoutIndexes[index] ?? index)}" r:id="rId${index + 1}"/>`).join("")}</p:sldLayoutIdLst>` : ""}<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr lang="en-US" sz="4400"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr lang="en-US" sz="2800"/></a:lvl1pPr></p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:otherStyle></p:txStyles></p:sldMaster>`;
}

function notesImagePlaceholderXml(id: number, geometry = ""): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Slide Image"/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr>${geometry}</p:spPr></p:sp>`;
}

function notesFrameXml(x: number, y: number, width: number, height: number): string {
  return `<a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(width)}" cy="${emu(height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`;
}

export function notesMasterXml(slideSize: { width: number; height: number }): string {
  // The fixed 540x720pt notes page is separate from the slide. Fit the actual
  // slide aspect into its upper frame; the lower body never overlaps it.
  // Use the same rounded dimensions written to p:sldSz, so a value near an
  // EMU boundary does not produce a different notes-thumbnail aspect ratio.
  const slideWidth = emu(slideSize.width), slideHeight = emu(slideSize.height);
  const scale = Math.min(468 / slideWidth, 288 / slideHeight);
  const width = slideWidth * scale, height = slideHeight * scale;
  const image = notesImagePlaceholderXml(2, notesFrameXml(36 + (468 - width) / 2, 36 + (288 - height) / 2, width, height));
  const body = `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr>${notesFrameXml(36, 360, 468, 324)}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml()}</p:grpSpPr>${image}${body}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:notesStyle><a:lvl1pPr><a:defRPr lang="en-US"/></a:lvl1pPr></p:notesStyle></p:notesMaster>`;
}

export function notesSlideXml(notes: NormalizedTextParagraph[], context: ElementXmlContext): string {
  // Keep empty paragraphs when they carry authored formatting or actions, while
  // omitting the plain default paragraph synthesized during normalization.
  const hasAuthoredContent = notes.some((paragraph) => {
    const paragraphStyle = paragraph.style;
    const paragraphIsDefault = paragraphStyle.align === "left"
      && paragraphStyle.indent === 0
      && paragraphStyle.hanging === 0
      && paragraphStyle.lineSpacing === 1
      && paragraphStyle.spaceBefore === 0
      && paragraphStyle.spaceAfter === 0
      && paragraphStyle.bullet.type === "none";
    return !paragraphIsDefault || paragraph.runs.some((run) => {
      const style = run.style;
      const runIsDefault = typeof style.fontFamily === "object"
        && "theme" in style.fontFamily
        && style.fontFamily.theme === "body"
        && style.fontSize === 18
        && style.bold === false
        && style.italic === false
        && style.underline === false
        && style.strike === false
        && typeof style.color === "object"
        && "theme" in style.color
        && style.color.theme === "text1"
        && (style.colorOpacity ?? 1) === 1
        && style.language === "en-US";
      return run.text.length > 0 || run.action !== undefined || !runIsDefault;
    });
  });
  const body = hasAuthoredContent ? `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Text"/><p:cNvSpPr txBox="1"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${textParagraphsXml(notes, context)}</p:txBody></p:sp>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml()}</p:grpSpPr>${body}${notesImagePlaceholderXml(3)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`;
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
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1"${layout.showMasterShapes === undefined ? "" : ` showMasterSp="${layout.showMasterShapes ? 1 : 0}"`}><p:cSld name="${escapeXml(layout.name)}">${backgroundXml(layout.backgroundSource === "master" ? undefined : layout.background)}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${groupTransformXml()}</p:grpSpPr>${elements.join("")}${placeholders}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function themeEntry(tag: string, color: string): string {
  return `<a:${tag}><a:srgbClr val="${colorValue(color)}"/></a:${tag}>`;
}

export function themeXml(theme: NormalizedPresentationTheme): string {
  const c = theme.colors;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${escapeXml(theme.name)}"><a:themeElements><a:clrScheme name="${escapeXml(theme.name)}">${themeEntry("dk1", c.text1)}${themeEntry("lt1", c.background1)}${themeEntry("dk2", c.text2)}${themeEntry("lt2", c.background2)}${themeEntry("accent1", c.accent1)}${themeEntry("accent2", c.accent2)}${themeEntry("accent3", c.accent3)}${themeEntry("accent4", c.accent4)}${themeEntry("accent5", c.accent5)}${themeEntry("accent6", c.accent6)}${themeEntry("hlink", c.hyperlink)}${themeEntry("folHlink", c.followedHyperlink)}</a:clrScheme><a:fontScheme name="${escapeXml(theme.name)}"><a:majorFont><a:latin typeface="${escapeXml(theme.fonts.heading)}"/><a:ea typeface="${escapeXml(theme.fonts.heading)}"/><a:cs typeface="${escapeXml(theme.fonts.heading)}"/></a:majorFont><a:minorFont><a:latin typeface="${escapeXml(theme.fonts.body)}"/><a:ea typeface="${escapeXml(theme.fonts.body)}"/><a:cs typeface="${escapeXml(theme.fonts.body)}"/></a:minorFont></a:fontScheme><a:fmtScheme name="${escapeXml(theme.name)}"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="28575"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

export function customPropertiesXml(metadata: NormalizedPresentationMetadata): string {
  const properties = metadata.customProperties ?? [];
  const body = properties.map((property, index) => {
    const overlappingEscape = /_x[0-9a-fA-F]{4}_x[0-9a-fA-F]{4}_/;
    if (overlappingEscape.test(property.name) || property.type === "string" && overlappingEscape.test(property.value)) throw new RangeError("Overlapping OOXML escape-like sequences are not supported for custom properties.");
    const type = { string: "lpwstr", boolean: "bool", integer: "i4", real: "r8", date: "filetime" }[property.type];
    const value = property.type === "boolean" ? (property.value ? "true" : "false") : property.type === "real" && Object.is(property.value, -0) ? "-0" : String(property.value);
    const protect = (text: string) => text.replace(/_x[0-9a-fA-F]{4}_/g, match => `_x005F_${match.slice(1)}`);
    return `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${index + 2}" name="${escapeXml(protect(property.name)).replace(/\t/g, "&#9;").replace(/\n/g, "&#10;").replace(/\r/g, "&#13;")}"><vt:${type}>${escapeXml(property.type === "string" ? protect(value) : value).replace(/\r/g, "&#13;")}</vt:${type}></property>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${body}</Properties>`;
}

export function corePropertiesXml(metadata: NormalizedPresentationMetadata): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:cpTerms="http://purl.org/dc/terms/"${metadata.createdAt === undefined && metadata.modifiedAt === undefined ? "" : ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'}><dc:title>${escapeXml(metadata.title)}</dc:title><dc:creator>${escapeXml(metadata.author)}</dc:creator>${metadata.subject === undefined ? "" : `<dc:subject>${escapeXml(metadata.subject)}</dc:subject>`}${metadata.description === undefined ? "" : `<dc:description>${escapeXml(metadata.description)}</dc:description>`}<cp:keywords>${escapeXml(metadata.keywords.join(", "))}</cp:keywords><cp:revision>${metadata.revision}</cp:revision><dc:language>${escapeXml(metadata.language)}</dc:language>${metadata.createdAt === undefined ? "" : `<cpTerms:created xsi:type="cpTerms:W3CDTF">${escapeXml(metadata.createdAt)}</cpTerms:created>`}${metadata.modifiedAt === undefined ? "" : `<cpTerms:modified xsi:type="cpTerms:W3CDTF">${escapeXml(metadata.modifiedAt)}</cpTerms:modified>`}</cp:coreProperties>`;
}

/** Describe the serialized frame, without guessing a native size preset. */
function presentationFormat(size: PresentationSize): string {
  const width = emu(size.width), height = emu(size.height);
  if (!Number.isSafeInteger(width) || !Number.isSafeInteger(height) || width <= 0 || height <= 0) return "Custom";
  const x = BigInt(width), y = BigInt(height);
  for (const [w, h] of [[16n, 9n], [4n, 3n], [16n, 10n]] as const) {
    if (x * h === y * w) return `On-screen Show (${w}:${h})`;
  }
  return "Custom";
}

export function appPropertiesXml(slideCount: number, metadata: NormalizedPresentationMetadata, size: PresentationSize): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>PPTKit</Application>${metadata.company === undefined ? "" : `<Company>${escapeXml(metadata.company)}</Company>`}<Slides>${slideCount}</Slides><PresentationFormat>${presentationFormat(size)}</PresentationFormat></Properties>`;
}
