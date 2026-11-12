import { isTextStylePresetResetProperty } from "../text/run-properties.js";
import { resetFrameMarginSide } from "../text/reset-frame-margin.js";
import { resetScriptFont } from "../text/reset-script-font.js";
import type { TextFontScriptCode } from "../text/font-scripts.js";
import type { Insets } from "../types/geometry.js";
import type { TextStylePresetResetProperty } from "../types/style.js";
import { DEFAULT_LAYOUT_ID } from "../constants/presentation.js";
import { estimateTextHeight } from "../normalization/measure-text.js";
import { normalizePresentation } from "../normalization/normalize-presentation.js";
import type { NormalizedElement, PresentationElement, PresentationElementInput, TextContentInput } from "../types/element.js";
import type { JsonValue, PresentationData, PresentationSlideData } from "../types/presentation.js";
import type { NormalizedPresentationTheme } from "../types/theme.js";
import { LayoutRebindError } from "../types/layout-rebind.js";
import type { LayoutRebindFinding, LayoutRebindMatch, LayoutRebindOptions, LayoutRebindPlan, LayoutRebindPropertyChange, LayoutRebindRequest } from "../types/layout-rebind.js";
import { deepClone, deepFreeze } from "../utils/clone.js";
import { validatePresentation } from "../validation/validate-presentation.js";
import { captureDefinitionInput } from "./definition-input.js";

export interface StagedSlideRebind { readonly slideId: string; readonly layoutId: string | undefined; readonly background: PresentationSlideData["background"]; readonly elements: readonly PresentationElementInput[]; }

function fields(value: object, allowed: readonly string[]): void {
  if (!value || typeof value !== "object" || Array.isArray(value) || Object.keys(value).some(key => !allowed.includes(key))) throw new TypeError("Unsupported layout rebind fields.");
}
export function slideRebindData(s:PresentationSlideData):PresentationSlideData {return { id:s.id, layoutId:s.layoutId, background:s.background, elements:s.elements, notes:s.notes, transition:s.transition, eventTimeline:s.eventTimeline,readingOrder:s.readingOrder, hidden:s.hidden, section:s.section, tags:s.tags, customData:s.customData, showInheritedShapes:s.showInheritedShapes };}
function captureData(document: PresentationData): PresentationData {
  const slides = document.slides.map(slideRebindData);
  return captureDefinitionInput({ id:document.id, metadata:document.metadata, size:document.size, theme:document.theme, textStylePresets:document.textStylePresets, assets:document.assets, layouts:document.layouts, masters:document.masters ?? [], slides }, ["id","metadata","size","theme","textStylePresets","assets","layouts","masters","slides"]);
}
function resolveTokens<T>(value: T, theme: NormalizedPresentationTheme): T {
  function visit(v: unknown): unknown {
    if (v === null || typeof v !== "object") return v;
    if (Array.isArray(v)) return v.map(visit);
    const obj = v as Record<string, unknown>;
    if (Object.keys(obj).length === 1 && typeof obj.theme === "string") {
      if (Object.hasOwn(theme.colors,obj.theme)) return theme.colors[obj.theme as keyof typeof theme.colors];
      if (Object.hasOwn(theme.fonts,obj.theme)) return theme.fonts[obj.theme as keyof typeof theme.fonts];
    }
    return Object.fromEntries(Object.entries(obj).map(([k,v]) => [k,visit(v)]));
  }
  return visit(value) as T;
}
function materialize(element: NormalizedElement, theme: NormalizedPresentationTheme): PresentationElementInput {
  const n = resolveTokens(element,theme);
  if (n.type === "text") { const { plainText: _plain, ...value } = n; return value; }
  if (n.type === "shape") {
    if (!n.text) return n;
    const { plainText: _plain, ...text } = n.text; return { ...n, text };
  }
  if (n.type === "group") return { ...n, children:n.children.map(child => materialize(child,theme)) };
  if (n.type === "table") { const { styleSource: _source, ...table } = n; return table; }
  if (n.type === "chart" && n.chartType === "line") return { ...n, series:n.series.map(series => ({ ...series, line:series.line === false ? { paint:{ type:"none" as const }, width:0 } : series.line })) };
  return n;
}
function cloneInput(element:PresentationElement):PresentationElementInput {
  return element.type === "group" ? { ...deepClone(element), children:element.children.map(cloneInput) } : deepClone(element);
}
function clearContentStyles(content: TextContentInput): TextContentInput {
  return typeof content === "string" ? content : content.map(p => ({ runs:p.runs.map(r => ({ text:r.text, ...(r.action === undefined ? {} : { action:deepClone(r.action) }) })) }));
}
function resetText(input: PresentationElementInput): void {
  if (input.type === "text") { delete input.frame; delete input.textStylePreset; delete input.textStylePresetReset; input.content=clearContentStyles(input.content); }
  else if (input.type === "shape" && input.text) { delete input.text.frame; delete input.text.textStylePreset; delete input.text.textStylePresetReset; input.text.content=clearContentStyles(input.text.content); }
}
function resetTextProperties(input: PresentationElementInput, properties: readonly TextStylePresetResetProperty[]): void {
  const text = input.type === "text" ? input : input.type === "shape" ? input.text : undefined;
  if (!text || !properties.length) return;
  text.textStylePresetReset = [...new Set([...(text.textStylePresetReset ?? []), ...properties])];
  for (const property of properties) {
    if (property.startsWith("frame.margin.")) resetFrameMarginSide(text.frame, property.slice(13) as keyof Insets);
    else if (property.startsWith("frame.")) { if (text.frame) delete text.frame[property.slice(6) as keyof typeof text.frame]; }
    else if (typeof text.content !== "string") for (const paragraph of text.content) {
      if (property.startsWith("paragraph.")) { if (paragraph.style) delete paragraph.style[property.slice(10) as keyof typeof paragraph.style]; }
      else if (property.startsWith("scriptFontFamilies.")) for (const run of paragraph.runs) resetScriptFont(run.style, property.slice(19) as TextFontScriptCode);
      else for (const run of paragraph.runs) if (run.style) delete run.style[property as keyof typeof run.style];
    }
  }
}
function inventory(elements: readonly PresentationElement[]): Set<string> {
  const ids=new Set<string>(), stack=[...elements]; while(stack.length) { const e=stack.pop()!; ids.add(e.id); if(e.type==="group") stack.push(...e.children); } return ids;
}
function changes(before: unknown, after: unknown, slideId: string, path: string, out: LayoutRebindPropertyChange[]): void {
  if (JSON.stringify(before) === JSON.stringify(after)) return;
  if (before && after && typeof before === "object" && typeof after === "object" && Array.isArray(before) === Array.isArray(after)) {
    const x=before as Record<string,unknown>,y=after as Record<string,unknown>;
    const keys=Array.isArray(before) ? Array.from({length:Math.max(before.length,(after as unknown[]).length)},(_,i)=>String(i)) : [...new Set([...Object.keys(x),...Object.keys(y)])].sort();
    for(const key of keys) changes(x[key],y[key],slideId,path ? `${path}.${key}` : key,out);
  } else out.push({slideId,path,...(before===undefined?{}:{before:JSON.parse(JSON.stringify(before)) as JsonValue}),...(after===undefined?{}:{after:JSON.parse(JSON.stringify(after)) as JsonValue})});
}
function fontFamilies(value:unknown,out=new Set<string>()):Set<string> {
  if (!value || typeof value !== "object") return out;
  for (const [key,item] of Object.entries(value)) {
    if (["fontFamily","eastAsianFontFamily","complexScriptFontFamily"].includes(key) && typeof item === "string") out.add(item);
    else if (key === "scriptFontFamilies" && item && typeof item === "object") {
      for (const family of Object.values(item)) if (typeof family === "string") out.add(family);
    } else fontFamilies(item,out);
  }
  return out;
}
function overflowRisk(elements:readonly NormalizedElement[]):boolean {
  for(const e of elements) {
    if(e.type==="group" && overflowRisk(e.children))return true;
    const text=e.type==="text"?e:e.type==="shape"?e.text:undefined;
    if(text && estimateTextHeight(e.box.width,text.content,text.frame)>e.box.height)return true;
  }
  return false;
}
function textRisk(elements: readonly NormalizedElement[]): boolean {
  for(const e of elements) { if(e.type==="group" && textRisk(e.children)) return true; const text=e.type==="text"?e:e.type==="shape"?e.text:undefined; if(text?.frame.autoFit.mode==="shrink") return true; } return false;
}

/** Plan from immutable values only; apply recomputes this stage instead of trusting candidate state. */
export function stageLayoutRebind(document: PresentationData, requestsInput: readonly LayoutRebindRequest[], optionsInput: LayoutRebindOptions, expectedSourceSnapshot?:string): { plan:LayoutRebindPlan; slides:readonly StagedSlideRebind[] } {
  const captured=captureDefinitionInput({requests:requestsInput,options:optionsInput},["requests","options"]), {options}=captured;
  if(!Array.isArray(captured.requests)) throw new TypeError("Layout rebind requests must be an array.");
  fields(options,["localAppearance","unmatched","reset"]);
  if(!["preserve","adopt"].includes(options.localAppearance) || (options.unmatched!==undefined && !["block","detach"].includes(options.unmatched))) throw new TypeError("Explicit local appearance and a supported unmatched policy are required.");
  const reset=options.reset ?? [];
  if(!Array.isArray(reset) || new Set(reset).size!==reset.length || reset.some(x=>!["geometry","text","background"].includes(x) && !(typeof x === "string" && x.startsWith("text.") && isTextStylePresetResetProperty(x.slice(5)))) || (options.localAppearance==="preserve" && reset.length>0)) throw new TypeError("Reset must be a unique supported list in adopt mode only.");
  let data:PresentationData;
  try { data=captureData(document); }
  catch(error) { if(expectedSourceSnapshot!==undefined)throw new LayoutRebindError("stale-plan","Current authoring data no longer matches the reviewed finite snapshot.");throw error; }
  const sourceSnapshot=JSON.stringify(data);
  if(expectedSourceSnapshot!==undefined && sourceSnapshot!==expectedSourceSnapshot)throw new LayoutRebindError("stale-plan","Authoring data changed after the layout rebind was prepared.");
  const before=normalizePresentation(data),bySlide=new Map(data.slides.map((s,i)=>[s.id,i])),requested=new Map<string,LayoutRebindRequest>();
  for(const request of captured.requests) {
    fields(request,["slideId","layoutId","bindings","references"]);
    if(typeof request.slideId!=="string" || !bySlide.has(request.slideId) || requested.has(request.slideId) || typeof request.layoutId!=="string") throw new TypeError("Requests need unique existing slide IDs and target layout IDs.");
    if(!before.layouts.some(l=>l.id===request.layoutId)) throw new TypeError(`Unknown target layout "${request.layoutId}".`);
    for(const key of ["bindings","references"] as const) if(request[key]!==undefined) {
      const values=request[key]!; if(!values || typeof values!=="object" || Array.isArray(values) || Object.entries(values).some(([k,v])=>!k || (typeof v!=="string" && !(key==="bindings" && v===null)))) throw new TypeError("Binding/reference maps need logical IDs and string keys (or null detach).");
    }
    requested.set(request.slideId,request);
  }
  const findings:LayoutRebindFinding[]=[], matches:LayoutRebindMatch[]=[], slides:StagedSlideRebind[]=[];
  function finding(slideId:string,code:string,severity:LayoutRebindFinding["severity"],message:string,elementId?:string,path?:string):void { findings.push({slideId,code,severity,message,...(elementId===undefined?{}:{elementId}),...(path===undefined?{}:{path})}); }
  for(const[sIndex,slide]of data.slides.entries()) {
    const request=requested.get(slide.id); if(!request) continue;
    const old=before.slides[sIndex]!, sourceLayout=before.layouts.find(l=>l.id===old.layoutId)!, target=before.layouts.find(l=>l.id===request.layoutId)!, oldTheme=sourceLayout.theme ?? before.theme;
    const sourceDefinition=data.layouts.find(l=>l.id===sourceLayout.id), targetDefinition=data.layouts.find(l=>l.id===target.id);
    const sourceInherited=inventory([...(sourceDefinition?.elements??[]),...(data.masters?.find(m=>m.id===sourceDefinition?.masterId)?.elements??[])]), targetInherited=inventory([...(targetDefinition?.elements??[]),...(data.masters?.find(m=>m.id===targetDefinition?.masterId)?.elements??[])]);
    const targetInheritedRoots=new Set([...(targetDefinition?.elements??[]),...(data.masters?.find(m=>m.id===targetDefinition?.masterId)?.elements??[])].map(e=>e.id));
    const topIds=new Set(slide.elements.map(e=>e.id));
    for(const id of Object.keys(request.bindings??{})) if(!topIds.has(id)) throw new TypeError(`Unknown top-level binding element "${id}".`);
    for(const[from,to]of Object.entries(request.references??{})) if(!sourceInherited.has(from) || !targetInherited.has(to)) throw new TypeError("Reference maps must link source and target inherited objects.");
    const usedReferences=new Set<string>();
    const sameLayout=sourceLayout.id===target.id && reset.length===0 && Object.keys(request.bindings??{}).length===0 && Object.keys(request.references??{}).length===0;
    const elements=slide.elements.map((element,index):PresentationElementInput=>{
      const normalized=old.elements[index]!, source=sourceLayout.placeholders.find(p=>p.key===element.placeholderKey), explicit=Object.hasOwn(request.bindings??{},element.id);
      let targetKey:string|undefined,method:LayoutRebindMatch["method"]="ordinary",candidates:string[]=[];
      if(explicit) {
        method="explicit"; const key=request.bindings![element.id];
        if(key!==null) { const p=target.placeholders.find(p=>p.key===key); if(!p) finding(slide.id,"missing-explicit-placeholder","error",`No target placeholder "${key}".`,element.id); else {targetKey=p.key;candidates=[p.key];if(source && source.kind!==p.kind) finding(slide.id,"explicit-role-change","warning","The explicit binding changes the placeholder role.",element.id);} }
      } else if(source) {
        const identity=target.placeholders.find(p=>p.key===source.key);
        if(identity && identity.kind===source.kind) {targetKey=identity.key;method="key";candidates=[identity.key];}
        else if(identity) { method="unmatched";candidates=[identity.key];finding(slide.id,"placeholder-kind-conflict","error","The same logical key has a different target role; choose explicitly.",element.id); }
        else {
          const indexed=source.index===undefined?undefined:target.placeholders.find(p=>p.index===source.index);
          if(indexed && indexed.kind===source.kind) {targetKey=indexed.key;method="index";candidates=[indexed.key];}
          else if(indexed) {method="unmatched";candidates=[indexed.key];finding(slide.id,"placeholder-index-kind-conflict","error","The same explicit index has a different target role; choose explicitly.",element.id);}
          else {
            candidates=target.placeholders.filter(p=>p.kind===source.kind).map(p=>p.key).sort();
            if(candidates.length===1) {targetKey=candidates[0];method="kind";}
            else {method="unmatched";finding(slide.id,candidates.length?"ambiguous-placeholder":"missing-placeholder",options.unmatched==="detach"&&!candidates.length?"warning":"error",candidates.length?"Multiple target placeholders share this role; choose a key explicitly.":"No target placeholder has this role.",element.id);}
          }
        }
      }
      const selected=target.placeholders.find(p=>p.key===targetKey);
      if(source?.index!==undefined && selected?.index!==undefined && source.index!==selected.index) finding(slide.id,"placeholder-index-changed","info","The selected target has a different explicit placeholder index.",element.id);
      matches.push({slideId:slide.id,elementId:element.id,...(element.placeholderKey===undefined?{}:{sourceKey:element.placeholderKey}),...(targetKey===undefined?{}:{targetKey}),...(source?.index===undefined?{}:{sourceIndex:source.index}),...(selected?.index===undefined?{}:{targetIndex:selected.index}),method,candidates});
      if(sameLayout) return cloneInput(element);
      let input=options.localAppearance==="preserve"?materialize(normalized,oldTheme):cloneInput(element);
      if(targetKey===undefined) {delete input.placeholderKey;if(element.placeholderKey!==undefined || explicit) input.box=deepClone(normalized.box);}
      else input.placeholderKey=targetKey;
      if(options.localAppearance==="adopt") { if(reset.includes("geometry") && targetKey!==undefined) delete input.box;if(reset.includes("text") && targetKey!==undefined) resetText(input); else if(targetKey!==undefined) resetTextProperties(input, reset.flatMap(item => item.startsWith("text.") && isTextStylePresetResetProperty(item.slice(5)) ? [item.slice(5) as TextStylePresetResetProperty] : [])); }
      function references(e:PresentationElementInput):void {
        if(e.type==="connector") for(const key of ["start","end"] as const) {
          const endpoint=e[key];if(!("elementId"in endpoint))continue;
          if(Object.hasOwn(request!.references??{},endpoint.elementId)) { const id=endpoint.elementId;usedReferences.add(id);e[key]={...endpoint,elementId:request!.references![id]!};if(options.localAppearance==="preserve" && e[key].elementId!==id) finding(slide.id,"inherited-connector-geometry","error","Preserving a connector needs unchanged inherited targets; adopt explicitly to follow a new target.",e.id); }
          if("elementId" in e[key] && targetInherited.has(e[key].elementId) && !targetInheritedRoots.has(e[key].elementId)) finding(slide.id,"inherited-reference-scope","error","A slide-level connector needs a target in the inherited root coordinate scope, not a grouped descendant.",e.id);
          if(!Object.hasOwn(request!.references??{},endpoint.elementId) && sourceInherited.has(endpoint.elementId) && !targetInherited.has(endpoint.elementId)) finding(slide.id,"missing-inherited-reference","error","An inherited connector target needs an explicit target object mapping.",e.id,`elements.${index}.${key}`);
        }
        if(e.type==="group")for(const child of e.children)references(child);
      }
      references(input);
      return input;
    });
    for(const from of Object.keys(request.references??{}))if(!usedReferences.has(from) && !sameLayout) finding(slide.id,"unused-reference-mapping","error",`No selected connector uses inherited object "${from}".`);
    if(!sameLayout) {
      finding(slide.id,"inherited-template-changed","info","Shared master/layout artwork always comes from the target template; it is not copied locally.");
      const stack=[...slide.elements];while(stack.length){const e=stack.pop()!;if(options.localAppearance==="preserve" && e.type==="table" && e.style)finding(slide.id,"table-style-materialized","warning","Preserving effective table appearance materializes regional rules into local cells.",e.id);if(e.type==="group")stack.push(...e.children);}
    }
    const background=sameLayout?slide.background:options.localAppearance==="preserve"?resolveTokens(old.background,oldTheme):reset.includes("background")?undefined:slide.background;
    slides.push({slideId:slide.id,layoutId:sameLayout?slide.layoutId:target.id===DEFAULT_LAYOUT_ID&&!targetDefinition?undefined:target.id,background,elements});
  }
  const staged=new Map(slides.map(s=>[s.slideId,s])), candidate:PresentationData={...data,slides:data.slides.map(s=>{const item=staged.get(s.id);return item?{...s,layoutId:item.layoutId,background:item.background,elements:item.elements as readonly PresentationElement[]}:s;})};
  const errors=validatePresentation(candidate).filter(d=>d.severity==="error");
  for(const d of errors) finding(d.slideId??"",`candidate-${d.code}`,"error",d.message,d.elementId,d.path);
  const differences:LayoutRebindPropertyChange[]=[],changedSlideIds:string[]=[];
  if(!errors.length && !findings.some(f=>f.severity==="error")) {
    const after=normalizePresentation(candidate);
    for(const item of slides) {
      const i=bySlide.get(item.slideId)!,old=data.slides[i]!,next=candidate.slides[i]!;
      if(JSON.stringify({layoutId:old.layoutId,background:old.background,elements:old.elements})!==JSON.stringify({layoutId:next.layoutId,background:next.background,elements:next.elements}))changedSlideIds.push(item.slideId);
      const x=before.slides[i]!,y=after.slides[i]!,xt=before.layouts.find(l=>l.id===x.layoutId)?.theme??before.theme,yt=after.layouts.find(l=>l.id===y.layoutId)?.theme??after.theme;
      changes(resolveTokens(x,xt),resolveTokens(y,yt),item.slideId,"",differences);
      const oldFonts=fontFamilies(resolveTokens(x.elements,xt)),newFonts=fontFamilies(resolveTokens(y.elements,yt));
      if([...newFonts].some(f=>!oldFonts.has(f)))finding(item.slideId,"font-availability-unverified","warning","Effective font families changed; Core does not inspect host font availability or shaping.");
      if(overflowRisk(y.elements)&&JSON.stringify(x.elements)!==JSON.stringify(y.elements))finding(item.slideId,"text-overflow-estimate","warning","Deterministic Core text estimation exceeds a target frame; inspect measured Layout/native output. No font size was changed to fit.");
      if(textRisk(y.elements)&&JSON.stringify(x.elements)!==JSON.stringify(y.elements))finding(item.slideId,"target-auto-fit-shrink","warning","The target retains an authored shrink policy; no additional scaling was inferred by rebinding.");
    }
  }
  const requests=data.slides.filter(s=>requested.has(s.id)).map(s=>requested.get(s.id)!);
  const plan:LayoutRebindPlan=deepFreeze({status:findings.some(f=>f.severity==="error")?"needs-review":"ready",sourceSnapshot,requests,options,matches,findings,changes:differences,changedSlideIds});
  return {plan,slides:deepFreeze(slides)};
}
export function prepareLayoutRebind(document:PresentationData,requests:readonly LayoutRebindRequest[],options:LayoutRebindOptions):LayoutRebindPlan {return stageLayoutRebind(document,requests,options).plan;}
