/** Named native controls for master-owned tables, charts and straight connectors. */
export function masterAdvancedDocument(createPresentation) {
  const d=createPresentation({id:'master-advanced',size:{width:720,height:540},metadata:{title:'Advanced master controls'},theme:{fonts:{heading:'Arial',body:'Arial'}}});
  const text=(id,value,x,y)=>({id,name:id,type:'text',box:{x,y,width:650,height:32},frame:{margin:0},content:[{runs:[{text:value,style:{fontFamily:'Arial',fontSize:18,color:{theme:'accent1'}}}]}]});
  for(const [id,color,values]of [['red','CC2200',[2,5]],['blue','0044CC',[3,4]]]){
    const shape=(suffix,x)=>({id:id+suffix,name:id+suffix,type:'shape',shape:'rect',box:{x,y:90,width:60,height:40},style:{fill:{type:'solid',color:{theme:'accent1'}},stroke:{paint:{type:'none'}}}});
    const charts=['bar','line','pie'].map((chartType,i)=>({id:id+'-'+chartType,name:id+'-'+chartType,type:'chart',chartType,box:{x:i*230,y:0,width:220,height:220},categories:['A','B'],title:id+' '+chartType,titleStyle:{fontFamily:'Arial',fontSize:14,bold:true},legend:{visible:false},style:{textStyle:{fontFamily:'Arial',fontSize:10}},series:[{name:id+' values',values,color:{theme:'accent1'},...(chartType==='pie'?{pointColors:[{theme:'accent1'},'44AA66']}:{})}]}));
    d.defineSlideMaster({id,name:id,theme:{colors:{accent1:color}},background:{type:'solid',color:'FFFFFF'},elements:[
      text(id+'-heading','MASTER '+id,20,15),
      {id:id+'-table-group',name:id+'-table-group',type:'group',box:{x:20,y:70,width:300,height:80},coordinateSize:{width:300,height:80},children:[{id:id+'-table',name:id+'-table',type:'table',box:{x:0,y:0,width:300,height:80},columns:[150,150],rows:[{height:40,cells:[{content:id+' shared table',colSpan:2}]},{height:40,cells:[{content:'Left '+id},{content:'Right '+id}]}],style:{wholeTable:{fill:{type:'solid',color:{theme:'accent1'}},text:{run:{fontFamily:'Arial',fontSize:14,color:'FFFFFF'}},margin:4}}}]},
      {id:id+'-edge',name:id+'-edge',type:'connector',start:{elementId:id+'-a',anchor:'right'},end:{elementId:id+'-b',anchor:'left'},style:{paint:{type:'solid',color:{theme:'accent1'}},width:3,endArrow:'triangle'}},shape('-a',400),shape('-b',600),
      {id:id+'-charts',name:id+'-charts',type:'group',box:{x:15,y:210,width:690,height:220},coordinateSize:{width:690,height:220},children:charts},
    ]});
    d.defineSlideLayout({id:id+'-layout',name:id+' layout',masterId:id});
  }
  for(const [i,id,show]of [[0,'red',undefined],[1,'red',false],[2,'red',undefined],[3,'blue',undefined]])d.addSlide({id:'page-'+i,layoutId:id+'-layout',...(show===undefined?{}:{showInheritedShapes:show}),notes:'Advanced notes '+i,elements:[text('page-label-'+i,'PAGE '+i,20,480)]});
  return d;
}

export function editMasterHeading(d){const red=d.masters.find(m=>m.id==='red');d.replaceSlideMaster('red',{...red,elements:red.elements.map(e=>e.id==='red-heading'?{...e,content:e.content.map(p=>({...p,runs:p.runs.map(r=>({...r,text:'MASTER red 2'}))}))}:e)});}
