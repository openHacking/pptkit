import { tableFixture, table, tableRow, cell, tableTransform } from "./table-fixture.mjs";
import { encodeStructure } from "./structure-fixture.mjs";
export const TABLE_STYLE_ID = "{12345678-1234-1234-1234-1234567890AB}";
export const paint = (color, scheme = false, extra = "") => `<a:solidFill><a:${scheme ? "schemeClr" : "srgbClr"} val="${color}">${extra}</a:${scheme ? "schemeClr" : "srgbClr"}></a:solidFill>`;
export const line = (color, width = 1) => `<a:ln w="${width*12700}">${paint(color)}<a:prstDash val="solid"/></a:ln>`;
export const region = (name, fill, text = "", borders = "") => `<a:${name}>${text ? `<a:tcTxStyle ${text}/>` : ""}<a:tcStyle>${fill ? `<a:fill>${fill}</a:fill>` : ""}${borders ? `<a:tcBdr>${borders}</a:tcBdr>` : ""}</a:tcStyle></a:${name}>`;
export function tableStyleFixture({ strict = false, flags = 'firstRow="1" lastRow="1" firstCol="1" lastCol="1" bandRow="1" bandCol="1"', rules, direct, rows = 4, columns = 4, inline = false, reference = true } = {}) {
  const body = Array.from({length:rows},(_,r)=>tableRow(Array.from({length:columns},(_,c)=>cell({text:`R${r}C${c}`,properties:r===1&&c===1&&direct ? direct : "<a:tcPr/>"})),30));
  const defined = rules ?? region("wholeTbl",paint("accent1",true), 'b="on"') + region("band1H",paint("112233")) + region("band2H",paint("223344")) + region("band1V",paint("334455")) + region("band2V",paint("445566")) + region("firstCol",paint("556677")) + region("lastCol",paint("667788")) + region("firstRow",paint("778899")) + region("lastRow",paint("8899AA")) + region("nwCell",paint("990000")) + region("neCell",paint("009900")) + region("swCell",paint("000099")) + region("seCell",paint("999900"));
  const f=tableFixture(table({columns:Array(columns).fill(70),rows:body,transform:tableTransform("",columns*70,rows*30),properties:`<a:tblPr ${flags}>${inline ? `<a:tableStyle styleId="${TABLE_STYLE_ID}" styleName="Inline">${defined}</a:tableStyle>` : reference ? `<a:tableStyleId>${TABLE_STYLE_ID}</a:tableStyleId>` : ""}</a:tblPr>`}),strict);
  const relation = `<Relationship Id="tableStyles" Type="${f.R}/tableStyles" Target="tableStyles.xml"/>`;
  f.parts['deck/_rels/main.data.rels']=f.parts['deck/_rels/main.data.rels'].replace('</Relationships>',relation+'</Relationships>');
  f.parts['deck/tableStyles.xml']=`<a:tblStyleLst xmlns:a="${f.A}" def="${TABLE_STYLE_ID}"><a:tblStyle styleId="${TABLE_STYLE_ID}" styleName="Fixture style">${defined}</a:tblStyle></a:tblStyleLst>`;
  f.types['deck/tableStyles.xml']='presentationml.tableStyles';
  f.parts['deck/theme.xml']=`<a:theme xmlns:a="${f.A}"><a:themeElements><a:clrScheme name="Fixture"><a:accent1><a:srgbClr val="3579BD"/></a:accent1><a:accent2><a:srgbClr val="CA6420"/></a:accent2><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1></a:clrScheme><a:fontScheme name="Fixture"><a:majorFont><a:latin typeface="Arial"/></a:majorFont><a:minorFont><a:latin typeface="Arial"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`;
  return f;
}
export const tableStyleBytes = options => encodeStructure(tableStyleFixture(options));
