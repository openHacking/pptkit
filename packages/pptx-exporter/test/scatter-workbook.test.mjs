import assert from 'node:assert/strict';
import test from 'node:test';
import { scatterWorkbook, scatterWorkbookRanges } from '../dist/ooxml/scatter-workbook.js';
import { readGeneratedZip } from '../../../tests/support/read-generated-zip.mjs';

test('scatter workbook keeps each series own X/Y range and blank cell positions',()=>{
 const data=[{name:'First & _x000A_',xValues:[-8,null,-0],yValues:[4,3,null]},{name:'Other',xValues:[7,-1],yValues:[-6,4]},{name:'Empty',xValues:[],yValues:[]}];
 assert.deepEqual(scatterWorkbookRanges(data),[
  {xColumn:'A',yColumn:'B',nameFormula:'Sheet1!$B$1',xFormula:'Sheet1!$A$2:$A$4',yFormula:'Sheet1!$B$2:$B$4'},
  {xColumn:'C',yColumn:'D',nameFormula:'Sheet1!$D$1',xFormula:'Sheet1!$C$2:$C$3',yFormula:'Sheet1!$D$2:$D$3'},
  {xColumn:'E',yColumn:'F',nameFormula:'Sheet1!$F$1'},
 ]);
 const sheet=readGeneratedZip(scatterWorkbook(data)).get('xl/worksheets/sheet1.xml').toString();
 assert.match(sheet,/<dimension ref="A1:F4"\/>/);assert.match(sheet,/First &amp; _x005F_x000A_/);
 for(const expected of ['<c r="A2"><v>-8</v></c>','<c r="B2"><v>4</v></c>','<c r="A3"/>','<c r="B3"><v>3</v></c>','<c r="A4"><v>-0</v></c>','<c r="B4"/>','<c r="C2"><v>7</v></c>','<c r="D3"><v>4</v></c>'])assert.ok(sheet.includes(expected),expected);
 assert.ok(!sheet.includes('<c r="C4"'));assert.ok(sheet.includes('<c r="F1"'));
});

test('scatter workbook rejects malformed, sparse and oversized data before writing',()=>{
 for(const input of [[],new Array(1),[{name:'',xValues:[],yValues:[]}],[{name:'S',xValues:[1],yValues:[]}],[{name:'S',xValues:[,],yValues:[1]}],[{name:'S',xValues:[Infinity],yValues:[1]}],new Array(8193)])assert.throws(()=>scatterWorkbook(input),RangeError);
 const tooMany=new Array(1_048_576);assert.throws(()=>scatterWorkbook([{name:'S',xValues:tooMany,yValues:tooMany}]),RangeError);
});

test('scatter workbook column names cross Excel letter boundaries',()=>{
 const data=Array.from({length:14},(_,i)=>({name:`S${i}`,xValues:[i],yValues:[i]})),last=scatterWorkbookRanges(data).at(-1);
 assert.deepEqual([last.xColumn,last.yColumn,last.xFormula,last.yFormula],['AA','AB','Sheet1!$AA$2:$AA$2','Sheet1!$AB$2:$AB$2']);
});

test('uneven series emit only their own rows without fabricated cells',()=>{
 const data=[{name:'long',xValues:Array.from({length:20},(_,i)=>i),yValues:Array.from({length:20},(_,i)=>-i)},...Array.from({length:30},(_,i)=>({name:`short${i}`,xValues:[i],yValues:[i]}))];
 const ranges=scatterWorkbookRanges(data),sheet=readGeneratedZip(scatterWorkbook(data)).get('xl/worksheets/sheet1.xml').toString();
 assert.ok(sheet.includes('<c r="A21"><v>19</v></c>'));for(const range of ranges.slice(1)){assert.ok(sheet.includes(`<c r="${range.xColumn}2">`));assert.ok(!sheet.includes(`<c r="${range.xColumn}3"`));}
});
