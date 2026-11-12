export function layoutRebindDocument(createPresentation) {
  const d=createPresentation({id:'layout-rebind',size:{width:720,height:540},metadata:{title:'Layout rebind controls'},theme:{fonts:{heading:'Arial',body:'Arial'}}});
  d.registerAsset({id:'logo',kind:'image',source:{type:'bytes',encoding:'base64',value:'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC'},mimeType:'image/png',width:8,height:4});
  for(const [id,color,font,bg]of [['red','CC2200','Arial','FFF8F0'],['blue','0044CC','Courier New','F0F8FF']])d.defineSlideMaster({id,name:id,theme:{colors:{accent1:color,background1:bg},fonts:{heading:font,body:font}},background:{type:'solid',color:{theme:'background1'}},elements:[{id:id+'-anchor',type:'shape',shape:'rect',box:{x:620,y:20,width:60,height:40},style:{fill:{type:'solid',color:{theme:'accent1'}}}}]});
  const ph=(key,kind,x,y,width,height,extra={})=>({key,kind,box:{x,y,width,height},...extra});
  d.defineSlideLayout({id:'source',name:'Source',masterId:'red',placeholders:[ph('title','title',20,20,500,50,{textStyle:{run:{fontFamily:{theme:'heading'},fontSize:24,color:{theme:'accent1'}}}}),ph('body-old','body',20,90,350,100,{textStyle:{frame:{margin:0},run:{fontSize:18,color:{theme:'accent1'}}}}),ph('picture-old','image',430,100,128,64)]});
  d.defineSlideLayout({id:'target',name:'Target',masterId:'blue',placeholders:[ph('photo','image',450,200,192,96),ph('copy','body',40,120,320,90,{textStyle:{frame:{margin:2},run:{fontFamily:{theme:'body'},fontSize:20,bold:true,color:{theme:'accent1'}}}}),ph('title','title',40,40,450,50,{textStyle:{run:{fontFamily:{theme:'heading'},fontSize:28,color:{theme:'accent1'}}}})]});
  d.defineSlideLayout({id:'choices',name:'Choices',masterId:'blue',placeholders:[ph('two','body',40,90,100,20),ph('one','body',180,90,100,20),ph('title','title',40,40,450,50),ph('photo','image',450,200,192,96)]});
  d.defineSlideLayout({id:'empty',name:'Empty',masterId:'blue'});
  for(let i=0;i<3;i++)d.addSlide({id:'page-'+i,layoutId:'source',notes:'Notes '+i,hidden:i===1,showInheritedShapes:i===1?false:undefined,elements:[
    {id:'title-'+i,type:'text',placeholderKey:'title',content:'TITLE '+i},
    {id:'body-'+i,type:'text',placeholderKey:'body-old',content:[{style:{spaceAfter:0},runs:[{text:'BODY '+i,style:{bold:false,italic:false,scriptFontFamilies:{Hans:{theme:'body'}}},action:{type:'url',url:'https://example.com/body'}}]}]},
    {id:'image-'+i,type:'image',placeholderKey:'picture-old',assetId:'logo',fit:'contain',crop:{left:0.1,right:0.1}},
    {id:'ordinary-'+i,type:'shape',shape:'rect',box:{x:20,y:250,width:140,height:40},style:{fill:{type:'solid',color:{theme:'accent1'},opacity:0.5},stroke:{paint:{type:'none'}}},text:{content:'Ordinary '+i}},
    {id:'group-'+i,type:'group',box:{x:180,y:300,width:480,height:160},coordinateSize:{width:480,height:160},children:[
      {id:'table-'+i,type:'table',box:{x:0,y:0,width:200,height:70},columns:[100,100],rows:[{height:35,cells:[{content:'Shared',colSpan:2}]},{height:35,cells:[{content:'A'},{content:'B'}]}],style:{wholeTable:{fill:{type:'solid',color:{theme:'accent1'}},text:{run:{color:'FFFFFF',fontFamily:{theme:'body'},fontSize:14}}}}},
      {id:'chart-'+i,type:'chart',chartType:'line',box:{x:230,y:0,width:230,height:150},categories:['A','B'],title:'Chart '+i,series:[{name:'Data',values:[2,4],color:{theme:'accent1'}}],legend:{visible:false}},
    ]},
    {id:'connector-'+i,type:'connector',start:{elementId:'body-'+i,anchor:'right'},end:{elementId:'image-'+i,anchor:'left'},style:{paint:{type:'solid',color:{theme:'accent1'}},width:2}},
  ]});
  return d;
}
