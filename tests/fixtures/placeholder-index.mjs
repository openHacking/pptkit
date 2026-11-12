export function placeholderIndexDocument(createPresentation){
 const d=createPresentation({id:'placeholder-index',size:{width:720,height:540},metadata:{title:'Explicit placeholder index controls'},theme:{fonts:{heading:'Arial',body:'Arial'}}});
 d.registerAsset({id:'logo',kind:'image',source:{type:'bytes',encoding:'base64',value:'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC'},mimeType:'image/png',width:8,height:4});
 for(const[id,color,font]of [['red','CC2200','Arial'],['blue','0044CC','Courier New']])d.defineSlideMaster({id,name:id,theme:{colors:{accent1:color},fonts:{heading:font,body:font}},elements:[{id:id+'-anchor',type:'shape',shape:'rect',box:{x:620,y:20,width:60,height:40},style:{fill:{type:'solid',color:{theme:'accent1'}}}}]});
 const ph=(key,kind,x,y,width,height,index)=>({key,kind,box:{x,y,width,height},...(index===undefined?{}:{index}),textStyle:{frame:{margin:0},run:{fontSize:18,color:{theme:'accent1'}}}});
 d.defineSlideLayout({id:'source',name:'Source',masterId:'red',placeholders:[ph('old-body','body',30,110,280,80,0),ph('old-photo','image',400,150,160,80,4294967295),ph('title','title',30,30,500,50,73),ph('legacy','subtitle',30,440,400,50)]});
 d.defineSlideLayout({id:'target',name:'Target',masterId:'blue',placeholders:[ph('decoy','body',390,330,280,80,17),ph('legacy','subtitle',40,440,400,50),ph('reserved','body',390,420,280,80,2),ph('photo','image',430,200,200,100,4294967295),ph('new-body','body',40,130,300,100,0),ph('title','title',40,40,500,60,73)]});
 for(let i=0;i<3;i++)d.addSlide({id:'page-'+i,layoutId:'source',notes:'Index notes '+i,hidden:i===1,elements:[
  {id:'title-'+i,type:'text',placeholderKey:'title',content:'TITLE '+i},
  {id:'body-'+i,type:'text',placeholderKey:'old-body',content:'BODY '+i},
  {id:'image-'+i,type:'image',placeholderKey:'old-photo',assetId:'logo',fit:'contain'},
  {id:'legacy-'+i,type:'text',placeholderKey:'legacy',content:'LEGACY '+i},
  {id:'ordinary-'+i,type:'shape',shape:'rect',box:{x:30,y:300,width:180,height:50},text:{content:'Ordinary '+i},style:{fill:{type:'solid',color:{theme:'accent1'}},stroke:{paint:{type:'none'}}}},
 ]});
 return d;
}
