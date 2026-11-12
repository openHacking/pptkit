import { tableStyleFixture } from "./table-style-fixture.mjs";
export const matrixFill = (token="phClr", modifiers="", type="schemeClr") => `<a:solidFill><a:${type} val="${token}">${modifiers}</a:${type}></a:solidFill>`;
export const matrixLine = (fill=matrixFill(), width=12700, dash="solid") => `<a:ln w="${width}" cap="flat" cmpd="sng" algn="ctr">${fill}<a:prstDash val="${dash}"/><a:miter lim="800000"/></a:ln>`;
export function formatScheme({ fills=[matrixFill(),matrixFill("123456","","srgbClr"),"<a:noFill/>"], lines=[matrixLine(),matrixLine("<a:noFill/>",25400),matrixLine(matrixFill("accent2"),38100,"dot")], backgrounds=[matrixFill("ABCDEF","","srgbClr"),matrixFill("phClr",'<a:alpha val="25000"/>'),"<a:noFill/>"] }={}) {
  return `<a:fmtScheme name="Matrix fixture"><a:fillStyleLst>${fills.join("")}</a:fillStyleLst><a:lnStyleLst>${lines.join("")}</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${backgrounds.join("")}</a:bgFillStyleLst></a:fmtScheme>`;
}
export const fillReference = (index=1,color='<a:schemeClr val="accent1"/>') => `<a:fillRef idx="${index}">${color}</a:fillRef>`;
export const lineReference = (index=1,color='<a:schemeClr val="accent2"/>') => `<a:lnRef idx="${index}">${color}</a:lnRef>`;
export function matrixFixture({strict=false,fill=fillReference(),border=lineReference(),scheme=formatScheme(),rules}={}){
  const f=tableStyleFixture({strict,flags:"",rows:2,columns:2,rules:rules??`<a:wholeTbl><a:tcStyle>${fill}<a:tcBdr>${["left","right","top","bottom","insideH","insideV"].map(side=>`<a:${side}>${border}</a:${side}>`).join("")}</a:tcBdr></a:tcStyle></a:wholeTbl>`});
  f.parts['deck/theme.xml']=f.parts['deck/theme.xml'].replace('</a:themeElements>',scheme+'</a:themeElements>');return f;
}
export function matrixOverride(f,scheme,extra=""){
  f.parts['deck/_rels/z.xml.rels']=f.parts['deck/_rels/z.xml.rels'].replace('</Relationships>',`<Relationship Id="matrixOverride" Type="${f.R}/themeOverride" Target="matrix-override.xml"/></Relationships>`);
  f.parts['deck/matrix-override.xml']=`<a:themeOverride xmlns:a="${f.A}">${scheme}${extra}</a:themeOverride>`;f.types['deck/matrix-override.xml']='themeOverride';return f;
}
