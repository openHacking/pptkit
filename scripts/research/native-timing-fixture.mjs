// Hand-authored research specimens, not an admitted animation serializer.
export const animationSpecimens = Object.freeze([
  { name: "fade-in", tag: "animEffect", attributes: { transition: "in", filter: "fade" }, category: "entr" },
  { name: "spin", tag: "animRot", attributes: { by: "21600000" }, category: "emph" },
  { name: "fade-out", tag: "animEffect", attributes: { transition: "out", filter: "fade" }, category: "exit" },
  { name: "line-path", tag: "animMotion", attributes: { origin: "layout", path: "M 0 0 L 0.25 0 E", pathEditMode: "relative" }, category: "path" },
].map(x => Object.freeze({ ...x, attributes: Object.freeze(x.attributes) })));

export function nativeTimingSpecimen(name, targetId) {
  const spec = animationSpecimens.find(x => x.name === name);
  if (!spec || !Number.isSafeInteger(targetId) || targetId < 1 || targetId > 4294967295) throw new TypeError("Unknown specimen or invalid target.");
  const attributes = Object.entries(spec.attributes).map(([key,value]) => ` ${key}="${value}"`).join("");
  const attrNames = spec.tag === "animRot" ? "<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst>" : spec.tag === "animMotion" ? "<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst>" : "";
  const behavior = `<p:${spec.tag}${attributes}><p:cBhvr><p:cTn id="5" dur="500" fill="hold"/><p:tgtEl><p:spTgt spid="${targetId}"/></p:tgtEl>${attrNames}</p:cBhvr></p:${spec.tag}>`;
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold" presetClass="${spec.category}" nodeType="clickEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${behavior}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
}
