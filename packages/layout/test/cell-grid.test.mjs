import assert from 'node:assert/strict';
import test from 'node:test';
import {cellGrid} from '../dist/table/cell-grid.js';

// Small-grid reference deliberately uses independently summed slices.
function reference(table){
 const cells=[],invalid=[],occupied=table.columns.map(()=>0);let y=0;
 table.rows.forEach((row,rowIndex)=>{let column=0;row.cells.forEach((cell,cellIndex)=>{
  while((occupied[column]??0)>rowIndex)column++;
  const end=column+cell.colSpan;
  if(!Number.isSafeInteger(cell.colSpan)||cell.colSpan<=0||!Number.isSafeInteger(cell.rowSpan)||cell.rowSpan<=0||end>table.columns.length||rowIndex+cell.rowSpan>table.rows.length||occupied.slice(column,end).some(v=>v>rowIndex)){invalid.push({rowIndex,cellIndex});return;}
  const box={x:table.columns.slice(0,column).reduce((a,b)=>a+b,0),y,width:table.columns.slice(column,end).reduce((a,b)=>a+b,0),height:table.rows.slice(rowIndex,rowIndex+cell.rowSpan).reduce((a,b)=>a+b.height,0)};
  if(![box.x,box.y,box.width,box.height].every(Number.isFinite)){invalid.push({rowIndex,cellIndex});return;}
  cells.push({rowIndex,cellIndex,columnIndex:column,box});for(let i=column;i<end;i++)occupied[i]=rowIndex+cell.rowSpan;column=end;
 });y+=row.height;});return {cells,invalid};
}

test('cell geometry agrees with independent slicing across seeded valid and malformed grids',()=>{
 let seed=918273;const rand=n=>{seed=(Math.imul(seed,1664525)+1013904223)>>>0;return seed%n;};
 for(let iteration=0;iteration<600;iteration++){
  const cols=1+rand(12),rows=1+rand(8),table={columns:Array.from({length:cols},()=>[.1,1,3,1e16][rand(4)]),rows:Array.from({length:rows},()=>({height:[.2,2,7][rand(3)],cells:Array.from({length:rand(cols+3)},()=>({colSpan:rand(5),rowSpan:rand(5)}))}))};
  assert.deepEqual(cellGrid(table),reference(table),`iteration ${iteration}`);
 }
});

test('wide tables resolve cell coordinates without repeatedly copying column prefixes',()=>{
 const n=12000,table={columns:Array(n).fill(.5),rows:[{height:12,cells:Array.from({length:n},()=>({colSpan:1,rowSpan:1}))}]};
 const result=cellGrid(table);assert.deepEqual(result.invalid,[]);assert.equal(result.cells.length,n);assert.deepEqual(result.cells[n-1],{rowIndex:0,cellIndex:n-1,columnIndex:n-1,box:{x:(n-1)*.5,y:0,width:.5,height:12}});
});
