export function presentationFormatDocuments(createPresentation) {
  const scenes=[['wide',960,540,'On-screen Show (16:9)'],['standard',720,540,'On-screen Show (4:3)'],['wide-ten',960,600,'On-screen Show (16:10)'],['portrait',540,960,'Custom'],['square',500,500,'Custom'],['paper',842,595,'Custom'],['rounded-wide',960+0.4/12700,540,'On-screen Show (16:9)'],['near-wide',960+0.6/12700,540,'Custom']];
  return scenes.map(([id,width,height,label])=>{const document=createPresentation({id,size:{width,height},metadata:{title:id,company:'A & <B>'}});document.addSlide({id:'page',notes:'Notes',elements:[{id:'text',type:'text',box:{x:20,y:20,width:300,height:50},content:id}]});return{document,label,cx:Math.round(width*12700),cy:Math.round(height*12700)};});
}
