import assert from 'node:assert/strict';
import test from 'node:test';
import { scatterSeriesXml } from '../dist/ooxml/scatter-series.js';

const marker={shape:'circle',size:10,fill:{type:'solid',color:'#FF0000',opacity:1},stroke:{paint:{type:'none'},width:1,dash:'solid',beginArrow:'none',endArrow:'none'}};
test('scatter series caches preserve per-axis ranges, blanks, order and marker-only styling',()=>{
 const xml=scatterSeriesXml([{name:'A & B',xValues:[-8,null,2],yValues:[4,3,null],line:false,marker},{name:'Empty',xValues:[],yValues:[],line:false,marker:false}]);
 assert.match(xml,/<c:idx val="0"\/><c:order val="0"\/>/);assert.match(xml,/<c:idx val="1"\/><c:order val="1"\/>/);
 assert.match(xml,/<c:f>Sheet1!\$A\$2:\$A\$4<\/c:f>/);assert.match(xml,/<c:f>Sheet1!\$B\$2:\$B\$4<\/c:f>/);
 assert.match(xml,/<c:xVal>.*?<c:ptCount val="3"\/><c:pt idx="0"><c:v>-8<\/c:v><\/c:pt><c:pt idx="2"><c:v>2<\/c:v><\/c:pt>.*?<\/c:xVal>/);
 assert.match(xml,/<c:yVal>.*?<c:ptCount val="3"\/><c:pt idx="0"><c:v>4<\/c:v><\/c:pt><c:pt idx="1"><c:v>3<\/c:v><\/c:pt>.*?<\/c:yVal>/);
 assert.match(xml,/<a:ln><a:noFill\/><\/a:ln>/);assert.match(xml,/<c:symbol val="circle"\/>/);assert.match(xml,/<c:symbol val="none"\/>/);assert.match(xml,/A &amp; B/);assert.equal((xml.match(/<c:numLit>/g)??[]).length,2);
});
test('scatter series rejects invalid arrays and values before generating XML',()=>{
 assert.throws(()=>scatterSeriesXml([{name:'bad',xValues:[1],yValues:[],line:false,marker:false}]),RangeError);
 assert.throws(()=>scatterSeriesXml([{name:'bad',xValues:[NaN],yValues:[1],line:false,marker:false}]),RangeError);
});
