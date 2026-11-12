export function masterArtworkDocument(createPresentation) {
  const d = createPresentation({ id: 'master-artwork', metadata: { title: 'Master artwork controls' }, size: { width: 720, height: 540 }, theme: { fonts: { heading: 'Arial', body: 'Arial' } } });
  d.registerAsset({ id: 'logo', kind: 'image', source: { type: 'bytes', encoding: 'base64', value: 'iVBORw0KGgoAAAANSUhEUgAAAAgAAAAECAIAAAA8r+mnAAAAIUlEQVR4nGNQCb8GRze90uCIAaeE0tE4OPp1RhSOcEoAAP18L9GcKmUbAAAAAElFTkSuQmCC' }, mimeType: 'image/png', width: 8, height: 4 });
  const text = (id, value, x, y) => ({ id, type: 'text', box: { x, y, width: 240, height: 44 }, frame: { margin: 0 }, content: [{ runs: [{ text: value, style: { fontFamily: 'Arial', fontSize: 20, color: { theme: 'accent1' } } }] }] });
  const shape = (id, color, x, y, width, height) => ({ id, type: 'shape', shape: 'rect', box: { x, y, width, height }, style: { fill: { type: 'solid', color }, stroke: { paint: { type: 'none' } } } });
  for (const [id, color, background] of [['red','CC2200','FFF8F0'],['blue','0044CC','F0F8FF']]) {
    d.defineSlideMaster({ id, name: id, theme: { colors: { accent1: color, background1: background } }, background: { type: 'solid', color: { theme: 'background1' } }, elements: [
      shape(id+'-master-shape', { theme: 'accent1' }, 20, 20, 260, 80), text(id+'-master-text', 'MASTER '+id, 20, 180),
      { id: id+'-master-logo', type: 'image', assetId: 'logo', box: { x: 20, y: 240, width: 64, height: 32 } },
      { id: id+'-master-group', type: 'group', box: { x: 500, y: 20, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [shape(id+'-group-shape','8844AA',0,0,50,50)] },
    ] });
  }
  for (const [id, masterId, show] of [['red-on','red',undefined],['red-off','red',false],['blue-on','blue',undefined]]) {
    d.defineSlideLayout({ id, name: id, masterId, ...(show===undefined?{}:{showMasterShapes:show}), elements: [shape(id+'-layout-shape','00AA55',60,40,140,80),text(id+'-layout-text','LAYOUT '+id,300,180),{ id:id+'-layout-logo',type:'image',assetId:'logo',box:{x:300,y:240,width:64,height:32} }], placeholders:[{key:'body',kind:'body',box:{x:20,y:340,width:400,height:50}}] });
  }
  for (const [index,layoutId,show,background] of [[0,'red-on',undefined,undefined],[1,'red-on',false,undefined],[2,'red-off',undefined,undefined],[3,'blue-on',undefined,undefined],[4,'red-off',true,undefined],[5,'red-on',undefined,{type:'solid',color:'FFFFDD'}]]) {
    d.addSlide({id:'page-'+index,layoutId,...(show===undefined?{}:{showInheritedShapes:show}),...(background?{background}:{}),notes:'Notes '+index,elements:[shape('local-shape-'+index,'FFCC00',100,60,60,80),text('local-text-'+index,'PAGE '+index,20,420),{id:'body-'+index,type:'text',placeholderKey:'body',content:'BODY '+index}]});
  }
  return d;
}
