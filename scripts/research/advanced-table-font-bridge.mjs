import assert from 'node:assert/strict';
import {readGeneratedZip} from '../../tests/support/read-generated-zip.mjs';

/** Exact compatibility delta for the two named Arial master-table controls. */
export function verifyAdvancedTableParagraphFontBridge(historicalBytes,currentBytes){
  const before=readGeneratedZip(historicalBytes),after=readGeneratedZip(currentBytes);
  assert.deepEqual([...before.keys()],[...after.keys()]);
  let tables=0,paragraphs=0;const changed=[];
  for(const[part,bytes]of before){
    if(!['ppt/slideMasters/slideMaster2.xml','ppt/slideMasters/slideMaster3.xml'].includes(part)){assert.deepEqual(after.get(part),bytes,part);continue;}
    const expected=bytes.toString().replace(/<a:tbl>[\s\S]*?<\/a:tbl>/g,table=>{
      tables++;
      return table.replace(/<a:p>[\s\S]*?<\/a:p>/g,paragraph=>{
        if(!paragraph.includes('<a:t>'))return paragraph;
        assert.equal((paragraph.match(/<a:endParaRPr lang="en-US"\/>/g)??[]).length,1);
        assert.match(paragraph,/<a:latin typeface="Arial"\/><a:ea typeface="Arial"\/><a:cs typeface="Arial"\/>/);
        paragraphs++;
        return paragraph.replace('<a:endParaRPr lang="en-US"/>','<a:endParaRPr lang="en-US"><a:latin typeface="Arial"/><a:ea typeface="Arial"/><a:cs typeface="Arial"/></a:endParaRPr>');
      });
    });
    assert.equal(after.get(part).toString(),expected,part);changed.push(part);
  }
  assert.equal(tables,2);assert.equal(paragraphs,6);
  return {changedParts:changed,tables,paragraphs,delta:'Only the six nonempty Arial table paragraph marks acquire explicit Latin/EA/CS Arial fonts. Every other byte of every package payload is unchanged.'};
}
