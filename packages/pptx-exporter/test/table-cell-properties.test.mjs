import assert from 'node:assert/strict';
import test from 'node:test';
import {createPresentation} from '@pptkit/core';
import {generatePptx} from '../dist/index.js';
import {readGeneratedZip} from '../../../tests/support/read-generated-zip.mjs';

test('table cell borders precede fill, and margins/anchor belong to cell properties across owners and nested groups',async()=>{
  for(const owner of ['slide','layout','master'])for(const alignment of ['top','middle','bottom'])for(const fill of [{type:'none'},{type:'solid',color:'CC2200'},{type:'solid',color:{theme:'accent1'},opacity:0.3}]){
    const d=createPresentation(),table={id:'table',type:'table',box:{x:0,y:0,width:200,height:80},columns:[100,100],rows:[{height:40,cells:[{content:'Merged',colSpan:2,style:{fill,margin:{left:0,right:3,top:5,bottom:7},verticalAlign:alignment}}]},{height:40,cells:[{content:'A',style:{fill,margin:0}},{content:'B',style:{fill,stroke:{paint:{type:'solid',color:'FF00FF'},width:2}}}]}]},group={id:'group',type:'group',box:{x:20,y:30,width:200,height:80},coordinateSize:{width:200,height:80},children:[table]};
    if(owner==='master'){d.defineSlideMaster({id:'m',name:'M',elements:[group]});d.defineSlideLayout({id:'l',name:'L',masterId:'m'});d.addSlide({layoutId:'l'});}
    else if(owner==='layout'){d.defineSlideLayout({id:'l',name:'L',elements:[group]});d.addSlide({layoutId:'l'});}
    else d.addSlide({elements:[group]});
    const parts=readGeneratedZip((await generatePptx(d)).bytes),part=owner==='master'?'ppt/slideMasters/slideMaster2.xml':owner==='layout'?'ppt/slideLayouts/slideLayout2.xml':'ppt/slides/slide1.xml',xml=parts.get(part).toString();
    const cells=[...xml.matchAll(/<a:tc(?:\s[^>]*)?>([\s\S]*?)<\/a:tc>/g)].map(m=>m[1]);assert.equal(cells.length,4);
    assert.match(cells[0],/<a:bodyPr\/>/);assert.doesNotMatch(cells[0],/\b[lrbt]Ins=/);
    assert.ok(cells[0].includes(`<a:tcPr marL="0" marR="38100" marT="63500" marB="88900" anchor="${{top:'t',middle:'ctr',bottom:'b'}[alignment]}">`));
    for(const cell of cells){if(cell.includes('<a:tcPr/>'))continue;const properties=/<a:tcPr(?:\s[^>]*)?>([\s\S]*?)<\/a:tcPr>/.exec(cell)[1];let tail=properties;for(const side of ['lnL','lnR','lnT','lnB']){assert.ok(tail.startsWith('<a:'+side+' '));tail=tail.slice(tail.indexOf('</a:'+side+'>')+side.length+5);}assert.ok(tail===''||/^<a:(noFill|solidFill|gradFill)/.test(tail));}
  }
});
