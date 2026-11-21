import { matrixFixture } from "../packages/pptx-parser/test/theme-matrix-fixture.mjs";
import { supplementalFontFixture } from "../packages/pptx-parser/test/supplemental-font-fixture.mjs";
import { notesStyleFixture } from "../packages/pptx-parser/test/notes-style-fixture.mjs";
import { notesFixture, notesShape } from "../packages/pptx-parser/test/notes-fixture.mjs";
import { body, shape } from "../packages/pptx-parser/test/content-fixture.mjs";
import { textAppearanceFixture, fontCollection } from "../packages/pptx-parser/test/table-text-appearance-fixture.mjs";
import { tableStyleBytes } from "../packages/pptx-parser/test/table-style-fixture.mjs";
import { paragraphBytes } from "../packages/pptx-parser/test/paragraph-target-fixture.mjs";
import { timingBytes } from "../packages/pptx-parser/test/timing-fixture.mjs";
import assert from "node:assert/strict";
import { readGeneratedZip } from '../tests/support/read-generated-zip.mjs';
import { createHash } from "node:crypto";
import { mkdtemp, mkdir, cp, readFile, realpath, rm, stat, writeFile } from "node:fs/promises";
import { spawnSync } from "node:child_process";
import { tmpdir } from "node:os";
import path from "node:path";
import test from "node:test";
import { chromium } from "@playwright/test";
import { getPublishablePackages, repoRoot } from "./release.mjs";
import { structureFixture, encodeStructure } from "../packages/pptx-parser/test/structure-fixture.mjs";
import { diagramResourcesFixture } from "../packages/pptx-parser/test/diagram-resources-fixture.mjs";
import { diagramPreviewFixture, cacheTextTransform } from "../packages/pptx-parser/test/diagram-preview-fixture.mjs";
import { chartFixture, chartXml, series, workbook, encodeChart, frame } from "../packages/pptx-parser/test/chart-fixture.mjs";
import { diagramTextFixture } from "../packages/pptx-parser/test/diagram-text-fixture.mjs";
import { diagramDataFixture } from "../packages/pptx-parser/test/diagram-data-fixture.mjs";
import { withSections } from "../packages/pptx-parser/test/sections-fixture.mjs";
import { placeholderBytes } from "../packages/pptx-parser/test/placeholder-fixture.mjs";
import { textStyleFixture, setTextTheme, setLayoutThemeFont, layoutList } from "../packages/pptx-parser/test/text-style-fixture.mjs";
import { imageFixture, encodeImages, picture, png, validGif as gif } from "../packages/pptx-parser/test/image-fixture.mjs";

function command(command, args, cwd) {
  const result = spawnSync(command, args, { cwd, encoding: "utf8", timeout: 120000, maxBuffer: 4 * 1024 * 1024,
    env: { ...process.env, NODE_PATH: "", npm_config_audit: "false", npm_config_fund: "false", npm_config_update_notifier: "false" },
  });
  assert.equal(result.status, 0, `${command} ${args.join(" ")} failed: ${result.error?.message ?? ""}\n${result.stderr}\n${result.stdout}`);
  return result.stdout;
}

const dataModule = (source) => `data:text/javascript;base64,${Buffer.from(source).toString("base64")}`;

test("packed packages work in an isolated npm consumer through every declared entry", { timeout: 240000 }, async (t) => {
  const temporary = await mkdtemp(path.join(tmpdir(), "pptkit-installed-"));
  t.after(() => rm(temporary, { recursive: true, force: true }));
  const tarballs = path.join(temporary, "tarballs");
  const consumer = path.join(temporary, "consumer");
  await mkdir(tarballs);
  await mkdir(consumer);
  const packages = getPublishablePackages();
  const archives = [];
  for (const pkg of packages) {
    const destination = path.join(tarballs, `${pkg.name.replace("@", "").replace("/", "-")}.tgz`);
    command("pnpm", ["pack", "--out", destination], pkg.dir);
    assert.ok((await stat(destination)).size > 0);
    archives.push(destination);
  }
  await writeFile(path.join(consumer, "package.json"), JSON.stringify({ name: "pptkit-isolated-consumer", private: true, type: "module" }));
  command("npm", ["install", "--ignore-scripts", "--no-audit", "--no-fund", "--package-lock=false", "--fetch-timeout=30000", "--fetch-retries=0", ...archives], consumer);
  const installed = path.join(consumer, "node_modules", "@pptkit");
  for (const pkg of packages) {
    const packageRoot = path.join(installed, pkg.name.split("/")[1]);
    assert.ok((await realpath(packageRoot)).startsWith(await realpath(consumer)), "package must not link to the workspace");
    const manifest = JSON.parse(await readFile(path.join(packageRoot, "package.json"), "utf8"));
    assert.equal(manifest.version, pkg.version);
    assert.equal(manifest.engines?.node, ">=20.0.0");
    for (const [name, version] of Object.entries(manifest.dependencies ?? {})) {
      assert.doesNotMatch(version, /^(workspace|link|file):/, `${pkg.name} dependency ${name} is publishable`);
      if (name.startsWith("@pptkit/")) assert.equal(version, pkg.version, "workspace dependencies use the actual packed version");
    }
    for (const target of Object.values(manifest.exports)) {
      for (const file of Object.values(target)) assert.ok((await stat(path.resolve(packageRoot, file))).isFile(), `${pkg.name} export ${file} exists`);
    }
  }
  await cp(path.join(repoRoot, "tests/consumers/node.mjs"), path.join(consumer, "node.mjs"));
  await cp(path.join(repoRoot, "tests/consumers/picture-save.mjs"), path.join(consumer, "picture-save.mjs"));
  await writeFile(path.join(consumer, "picture-source.pptx"), encodeImages(imageFixture(picture({ id: 2 }) + picture({ id: 3 }))));
  await writeFile(path.join(consumer, "picture-replacement.gif"), gif);
  await cp(path.join(repoRoot, "tests/consumers/types.ts"), path.join(consumer, "types.ts"));
  await writeFile(path.join(consumer, "theme-matrix-source.pptx"), encodeStructure(matrixFixture()));
  await writeFile(path.join(consumer, "supplemental-font-source.pptx"), encodeStructure(supplementalFontFixture({ emptySlots: true })));
  await writeFile(path.join(consumer, "declared-slot-font-source.pptx"), encodeStructure(supplementalFontFixture({ run: '<a:latin typeface="+mn-lt"/>' })));
  await writeFile(path.join(consumer, "notes-style-source.pptx"), encodeStructure(notesStyleFixture()));
  await writeFile(path.join(consumer, "field-text-source.pptx"), encodeStructure(notesFixture(notesShape({ text: body('<a:p><a:r><a:t>before</a:t></a:r><a:fld id="{38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A}" type="slidenum"><a:t>42</a:t></a:fld><a:fld id="{38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A}" type="datetime3"><a:t>date</a:t></a:fld><a:fld id="{38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A}" type="datetime13"><a:t>time</a:t></a:fld></a:p>') }))));
  await writeFile(path.join(consumer, "table-text-appearance.pptx"), encodeStructure(textAppearanceFixture({ text: fontCollection("+mn-lt", "+mn-ea", "+mn-cs") + '<a:srgbClr val="102030"/>' })));
  await writeFile(path.join(consumer, "table-style-source.pptx"), tableStyleBytes({ flags: "" }));
  await writeFile(path.join(consumer, "native-sections.pptx"), encodeStructure(withSections(structureFixture())));
  await writeFile(path.join(consumer, "smartart-inventory.pptx"), diagramResourcesFixture().bytes());
  await writeFile(path.join(consumer, "smartart-text.pptx"), diagramTextFixture().bytes());
  await writeFile(path.join(consumer, "native-timing-source.pptx"), timingBytes());
  await writeFile(path.join(consumer, "paragraph-target-source.pptx"), paragraphBytes());
  await writeFile(path.join(consumer, "chart-value-source.pptx"), encodeChart(chartFixture({ xml: chartXml({ data: series({ refs: true }), external: true }), book: workbook() })));
  const fieldChart = chartFixture({ content: frame() + shape({ id: "20", txBox: true, text: body('<a:p><a:r><a:t>before</a:t></a:r><a:fld id="{38D9DC4A-2E99-E34A-A7A3-8E01BBC23A4A}" type="slidenum"><a:t>42</a:t></a:fld></a:p>') }), xml: chartXml({ data: series({ refs: true }), external: true }), book: workbook() });
  await writeFile(path.join(consumer, "field-chart-source.pptx"), encodeChart(fieldChart));
  const graphFixture = diagramDataFixture().bytes();
  await writeFile(path.join(consumer, "smartart-data.pptx"), graphFixture);
  const cachedTextFixture = diagramPreviewFixture();
  cachedTextFixture.setCache(cachedTextFixture.cacheShape("2", { extra: cacheTextTransform() }));
  await writeFile(path.join(consumer, "smartart-preview.pptx"), cachedTextFixture.bytes());
  await writeFile(path.join(consumer, "placeholder-inheritance.pptx"), placeholderBytes());
  const themedText = setLayoutThemeFont(setTextTheme(textStyleFixture(), { accent1: "13579B", major: "Installed Base Font" }), "Installed Theme Font");
  themedText.parts["deck/layout.xml"] = themedText.parts["deck/layout.xml"].replace(layoutList,
    '<a:lvl1pPr><a:defRPr sz="2000"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr>');
  const textFixture = encodeStructure(themedText);
  await writeFile(path.join(consumer, "text-inheritance.pptx"), textFixture);
  await writeFile(path.join(consumer, "tsconfig.json"), JSON.stringify({ compilerOptions: {
    target: "ES2022", module: "NodeNext", moduleResolution: "NodeNext", strict: true, noEmit: true, skipLibCheck: false, types: [], lib: ["ES2022", "DOM"],
  }, files: ["types.ts"] }));

  await t.test("Node consumers resolve installed files and export real PPTX", () => {
    const output = JSON.parse(command(process.execPath, ["node.mjs"], consumer));
    assert.equal(output.slideCount, 1);
    assert.equal(output.previewCount, 1);
    assert.ok(output.bytes > 0);
  });

  await t.test("installed PDF adapter uses its declared dependencies", () => {
    const source = `import {createPresentation} from '@pptkit/core'; import {resolveLayout} from '@pptkit/layout'; import {renderLayoutToPdf} from '@pptkit/pdf-renderer'; const d=createPresentation();d.addSlide({id:'s'});let closed=false;const browser={newContext:async()=>({route:async()=>{},newPage:async()=>({setContent:async()=>{},evaluate:async e=>e.includes('querySelectorAll')?0:true,pdf:async()=>new TextEncoder().encode('%PDF-test')}),close:async()=>{closed=true}})};const r=await renderLayoutToPdf(resolveLayout(d),{browser});console.log(JSON.stringify([r.pages[0].slideId,closed,r.bytes.length]));`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), '["s",true,9]');
  });

  await t.test("installed media inspector rejects invalid JPEG structure", () => {
    const source = `import {validateJpegStructure} from '@pptkit/media-inspection';try{validateJpegStructure(new Uint8Array([255,216,255,217]));throw Error('accepted')}catch(e){if(!String(e.message).includes('incomplete image'))throw e;console.log('rejected')}`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), 'rejected');
  });

  await t.test("installed DOM entry imports in Node without browser globals", () => {
    const source = `import {normalizeJpegOrientationInBrowser,extractPngIccProfileInBrowser,orientJpegCrop} from '@pptkit/pptx-exporter/dom';console.log([normalizeJpegOrientationInBrowser,extractPngIccProfileInBrowser,orientJpegCrop].map(x=>typeof x).join(','))`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), 'function,function,function');
  });

  await t.test("installed Core exposes description inspection", () => {
    const source = `import {createPresentation,inspectMissingDescriptions,inspectEffectiveMissingDescriptions} from '@pptkit/core'; const d=createPresentation(); d.defineSlideLayout({id:'l',name:'Layout',elements:[{id:'t',type:'table',box:{x:0,y:0,width:100,height:50},columns:[100],rows:[{cells:[{content:'A'}]}]}]}); d.addSlide({id:'s',layoutId:'l'}); d.addSlide({id:'hidden',layoutId:'l',showInheritedShapes:false}); console.log(JSON.stringify([inspectMissingDescriptions(d).findings.map(f=>[f.elementId,f.layoutId]),inspectEffectiveMissingDescriptions(d).findings.map(f=>[f.elementId,f.slideId,f.source.kind])]));`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), '[[["t","l"]],[["t","s","layout"]]]');
  });
  await t.test("installed Core atomically reorders slide objects", () => {
    const source = `import {createPresentation} from '@pptkit/core'; const d=createPresentation(), box={x:0,y:0,width:10,height:10}; const s=d.addSlide({elements:[{id:'a',type:'shape',shape:'rect',box},{id:'g',type:'group',box,coordinateSize:{width:10,height:10},children:[{id:'x',type:'shape',shape:'rect',box},{id:'y',type:'shape',shape:'rect',box}]}]}); s.reorderElements(['g','a']); s.reorderGroupElements('g',['y','x']); console.log(s.elements.map(e=>e.id).join(',')+':'+s.elements[0].children.map(e=>e.id).join(','));`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), 'g,a:y,x');
  });
  await t.test("installed exporter writes decorative metadata", () => {
    const source = `import {createPresentation} from '@pptkit/core'; import {generatePptx} from '@pptkit/pptx-exporter'; const d=createPresentation(); d.addSlide({elements:[{type:'shape',shape:'rect',box:{x:0,y:0,width:100,height:50},accessibility:{decorative:true}}]}); console.log(Buffer.from((await generatePptx(d)).bytes).toString('base64'));`;
    const bytes = Buffer.from(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), 'base64');
    assert.match(readGeneratedZip(bytes).get('ppt/slides/slide1.xml').toString(), /<adec:decorative[^>]+val="1"/);
  });
  await t.test("installed parser saves one existing object description", () => {
    const source = `import {createPresentation} from '@pptkit/core'; import {generatePptx} from '@pptkit/pptx-exporter'; import {importPptx} from '@pptkit/pptx-parser'; const d=createPresentation();d.addSlide({elements:[{type:'shape',shape:'rect',box:{x:0,y:0,width:100,height:50},accessibility:{description:'old'}}]});const s=await importPptx((await generatePptx(d)).bytes,{projection:'content',descriptionEditSaveProfile:'existing-v1'});const slide=s.document.slides[0],e=slide.elements[0];slide.replaceElement(e.id,{...e,accessibility:{...e.accessibility,description:'new'}});const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw Error(JSON.stringify(p.report));const saved=await s.save(p.plan),next=await importPptx(saved.bytes,{projection:'content'});console.log(next.document.slides[0].elements[0].accessibility.description);next.dispose();s.dispose();`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), 'new');
  });
  await t.test("installed parser saves an atomic description batch", () => {
    const source = `import {createPresentation} from '@pptkit/core'; import {generatePptx} from '@pptkit/pptx-exporter'; import {importPptx} from '@pptkit/pptx-parser'; const d=createPresentation(),box={x:0,y:0,width:100,height:50}; d.addSlide({elements:[{type:'shape',shape:'rect',box,accessibility:{description:'a'}},{type:'shape',shape:'rect',box,accessibility:{description:'b'}}]}); const s=await importPptx((await generatePptx(d)).bytes,{projection:'content',descriptionEditSaveProfile:'batch-v3'}), slide=s.document.slides[0];for(const [i,e] of [...slide.elements].entries())slide.replaceElement(e.id,{...e,accessibility:{...e.accessibility,description:'new'+i}});const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw Error(JSON.stringify(p.report));const next=await importPptx((await s.save(p.plan)).bytes,{projection:'content'});console.log(next.document.slides[0].elements.map(e=>e.accessibility.description).join(','));next.dispose();s.dispose();`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), 'new0,new1');
  });
  await t.test("installed parser saves combined text and description", () => {
    const source = `import {createPresentation} from '@pptkit/core'; import {generatePptx} from '@pptkit/pptx-exporter'; import {importPptx} from '@pptkit/pptx-parser'; const d=createPresentation();d.addSlide({elements:[{type:'text',box:{x:0,y:0,width:100,height:50},content:'old',accessibility:{description:'old alt'}}]});const s=await importPptx((await generatePptx(d)).bytes,{projection:'content',editSaveProfile:'text-runs-v3',descriptionEditSaveProfile:'mixed-v4'}),slide=s.document.slides[0],e=structuredClone(slide.elements[0]);e.content[0].runs[0].text='new';e.accessibility.description='new alt';slide.replaceElement(e.id,e);const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw Error(JSON.stringify(p.report));const next=await importPptx((await s.save(p.plan)).bytes,{projection:'content'});console.log(next.document.slides[0].elements[0].content[0].runs[0].text+':'+next.document.slides[0].elements[0].accessibility.description);next.dispose();s.dispose();`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), 'new:new alt');
  });
  await t.test("installed parser preserves an existing drawing name", () => {
    const source = `import {createPresentation} from '@pptkit/core';import {generatePptx} from '@pptkit/pptx-exporter';import {importPptx} from '@pptkit/pptx-parser';const d=createPresentation();d.addSlide({elements:[{type:'shape',shape:'rect',name:'old',box:{x:0,y:0,width:100,height:50}}]});const s=await importPptx((await generatePptx(d)).bytes,{projection:'content',drawingNameSaveProfile:'existing-v1'}),slide=s.document.slides[0],e=structuredClone(slide.elements[0]);e.name='new';slide.replaceElement(e.id,e);const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw Error(JSON.stringify(p.report));const next=await importPptx((await s.save(p.plan)).bytes,{projection:'content'});console.log(next.document.slides[0].elements[0].name);next.dispose();s.dispose();`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), 'new');
  });
  await t.test("installed parser preserves an existing decorative value", () => {
    const source = `import {createPresentation} from '@pptkit/core';import {generatePptx} from '@pptkit/pptx-exporter';import {importPptx} from '@pptkit/pptx-parser';const d=createPresentation();d.addSlide({elements:[{type:'shape',shape:'rect',name:'old',accessibility:{decorative:true},box:{x:0,y:0,width:100,height:50}}]});const s=await importPptx((await generatePptx(d)).bytes,{projection:'content',decorativeSaveProfile:'existing-v1'}),slide=s.document.slides[0],e=structuredClone(slide.elements[0]);e.accessibility.decorative=false;slide.replaceElement(e.id,e);const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw Error(JSON.stringify(p.report));const next=await importPptx((await s.save(p.plan)).bytes,{projection:'content'});console.log(next.document.slides[0].elements[0].accessibility.decorative);next.dispose();s.dispose();`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), 'false');
  });
  await t.test("installed parser preserves slide visibility", () => {
    const source = `import {createPresentation} from '@pptkit/core';import {generatePptx} from '@pptkit/pptx-exporter';import {importPptx} from '@pptkit/pptx-parser';const d=createPresentation();d.addSlide({hidden:true});d.addSlide({hidden:false});const s=await importPptx((await generatePptx(d)).bytes,{projection:'content',slideVisibilitySaveProfile:'attribute-v1'});s.document.slides[0].setHidden(false);s.document.slides[1].setHidden(true);const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw Error(JSON.stringify(p.report));const next=await importPptx((await s.save(p.plan)).bytes,{projection:'content'});console.log(next.document.slides.map(x=>x.hidden).join(','));next.dispose();s.dispose();`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), 'false,true');
  });
  await t.test("installed parser preserves bound core metadata strings", () => {
    const source = `import {createPresentation} from '@pptkit/core';import {generatePptx} from '@pptkit/pptx-exporter';import {importPptx} from '@pptkit/pptx-parser';const d=createPresentation({metadata:{title:'old',author:'Alice',revision:7}});d.addSlide();const s=await importPptx((await generatePptx(d)).bytes,{projection:'content',metadataSaveProfile:'core-strings-v1'});s.document.setMetadata({...s.document.metadata,title:'new'});const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready')throw Error(JSON.stringify(p.report));const next=await importPptx((await s.save(p.plan)).bytes,{projection:'content'});console.log(next.document.metadata.title+':'+next.document.metadata.author+':'+next.document.metadata.revision);next.dispose();s.dispose();`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), 'new:Alice:7');
  });
  await t.test("installed Core exposes conservative contrast inspection", () => {
    const source = `import {createPresentation,inspectSlideTextContrast} from '@pptkit/core'; const d=createPresentation(); d.addSlide({elements:[{type:'text',box:{x:0,y:0,width:100,height:50},content:[{runs:[{text:'A',style:{color:'000000',fontSize:12}}]}]}]}); console.log(inspectSlideTextContrast(d).findings[0].ratio);`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', source], consumer).trim(), '21');
  });
  await t.test("installed Core exposes explicit field evaluation", () => {
    const source = `import { evaluatePresentationField } from '@pptkit/core'; console.log(evaluatePresentationField({id:'p',kind:'slide-number'},{pageOrder:['a','b'],pageId:'b',firstPageNumber:4}).value);`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "5");
  });

  await t.test("installed Core atomically registers asset batches", () => {
    const source = `import {createPresentation} from '@pptkit/core'; const d=createPresentation(); const a={kind:'image',source:{type:'bytes',value:new Uint8Array([1])}}; try { d.registerAssets([{asset:a},{asset:{...a,source:{type:'bytes',encoding:'base64',value:'invalid'}}}]); throw Error('missing failure'); } catch(e) { if(e.message==='missing failure') throw e; } const result=d.registerAssets([{asset:a},{asset:a,options:{deduplicate:'content'}}]); const origin=createPresentation(); origin.registerAsset({...a,id:'original'}); const copy=d.copyAssetsFrom(origin,['original'],{deduplicate:'content',external:'reject'}); console.log(d.assets.length+':'+result[0].id+':'+result[1].id+':'+copy[0].destinationId+':'+copy[0].reused);`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "1:asset-1:asset-1:asset-1:true");
  });

  await t.test("installed Core copies a master, layout and page between documents", () => {
    const source = `import {createPresentation} from '@pptkit/core'; const a=createPresentation(),b=createPresentation(),assets={deduplicate:'content',external:'reject'}; a.defineSlideMaster({id:'m',name:'Master',theme:{colors:{accent1:'123456'}}}); a.defineSlideLayout({id:'l',name:'Layout',masterId:'m'}); a.addSlide({id:'s',elements:[{id:'t',type:'text',content:'copy',box:{x:0,y:0,width:100,height:30}}]}); b.copySlideMasterFrom(a,'m',{id:'m2',theme:'source-effective',assets}); b.copySlideLayoutFrom(a,'l',{id:'l2',masterId:'m2',theme:'destination',assets}); const result=b.copySlideFrom(a,'s',{layoutId:'l2',theme:'destination',geometry:'preserve',customData:'preserve',assets}); console.log([b.masters[0].theme.colors.accent1,b.layouts[0].masterId,result.slide.layoutId,result.slide.elements[0].content].join(':'));`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "123456:m2:l2:copy");
  });

  await t.test("installed Core batch copy remaps links between selected pages", () => {
    const source = `import {createPresentation} from '@pptkit/core'; const a=createPresentation(),b=createPresentation(),box={x:0,y:0,width:10,height:10}; a.addSlide({id:'a',elements:[{id:'one',type:'text',box,content:'A',action:{type:'slide',slideId:'b'}}]}); a.addSlide({id:'b',elements:[{id:'two',type:'text',box,content:'B',action:{type:'slide',slideId:'a'}}]}); const r=b.copySlidesFrom(a,['a','b'],{layoutId:null,theme:'destination',geometry:'preserve',customData:'preserve',assets:{deduplicate:'content',external:'reject'}}); console.log([r.slideIds.a,r.slideIds.b,r.slides[0].elements[0].action.slideId,r.slides[1].elements[0].action.slideId].join(':'));`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "slide-1:slide-2:slide-2:slide-1");
  });

  await t.test("installed Core resolves renamed placeholders by explicit index", () => {
    const source = `import {createPresentation} from '@pptkit/core'; const a=createPresentation(),b=createPresentation(),box={x:0,y:0,width:100,height:30},assets={deduplicate:'content',external:'reject'}; a.defineSlideLayout({id:'l',name:'Source',placeholders:[{key:'old',kind:'title',index:0,box}]}); a.addSlide({id:'s',layoutId:'l',elements:[{id:'t',type:'text',placeholderKey:'old',content:'copy'}]}); b.defineSlideLayout({id:'next',name:'Target',placeholders:[{key:'new',kind:'title',index:0,box}]}); const r=b.copySlideFrom(a,'s',{layoutId:'next',theme:'destination',geometry:'preserve',customData:'reject',placeholderRoles:'require-match-and-index',placeholderMatching:'explicit-index',assets}); console.log(r.slide.elements[0].placeholderKey+':'+r.elementIds.t);`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "new:element-1");
  });

  await t.test("installed Core reuses an equivalent copied layout", () => {
    const source = `import {createPresentation} from '@pptkit/core'; const a=createPresentation(),b=createPresentation(),assets={deduplicate:'content',external:'reject'}; a.defineSlideLayout({id:'l',name:'Shared',elements:[{id:'s',type:'shape',shape:'rect',box:{x:0,y:0,width:10,height:10}}]}); const one=b.copySlideLayoutFrom(a,'l',{id:'one',masterId:null,theme:'destination',assets}); const two=b.copySlideLayoutFrom(a,'l',{id:'two',masterId:null,theme:'destination',assets,reuseEquivalent:true}); console.log([two.layout.id,two.elementIds.s,one.elementIds.s,b.layouts.length].join(':'));`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "one:element-1:element-1:1");
  });

  await t.test("installed Core reuses an equivalent copied master", () => {
    const source = `import {createPresentation} from '@pptkit/core'; const a=createPresentation(),b=createPresentation(),assets={deduplicate:'content',external:'reject'}; a.defineSlideMaster({id:'m',name:'Shared',elements:[{id:'s',type:'shape',shape:'rect',box:{x:0,y:0,width:10,height:10}}]}); const one=b.copySlideMasterFrom(a,'m',{id:'one',theme:'destination',assets}); const two=b.copySlideMasterFrom(a,'m',{id:'two',theme:'destination',assets,reuseEquivalent:true}); console.log([two.master.id,two.elementIds.s,one.elementIds.s,b.masters.length].join(':'));`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "one:element-1:element-1:1");
  });

  await t.test("installed Layout refreshes a native page-number cache", () => {
    const source = `import {createPresentation} from '@pptkit/core'; import {resolveLayout} from '@pptkit/layout'; const d=createPresentation(); d.addSlide({elements:[{type:'text',box:{x:0,y:0,width:100,height:30},content:[{runs:[{text:'stale',field:{id:'38d9dc4a-2e99-e34a-a7a3-8e01bbc23a4a',kind:'slide-number'}}]}]}]}); console.log(resolveLayout(d,{fieldRefresh:{firstPageNumber:7,includeHidden:true}}).slides[0].elements[0].plainText);`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "7");
  });

  await t.test("installed SVG renderer exposes the byte budget error", () => {
    const source = `import {createPresentation} from '@pptkit/core'; import {renderPresentationToSvg,SvgRenderLimitError,SvgRenderAbortError} from '@pptkit/svg-renderer'; const d=createPresentation(); d.addSlide(); try { await renderPresentationToSvg(d,{maxSvgBytes:0}); throw Error('missing error'); } catch(e) { if(!(e instanceof SvgRenderLimitError)) throw e; console.log(e.code+':'+e.slideIndex); } const c=new AbortController(); c.abort(); try { await renderPresentationToSvg(d,{signal:c.signal}); throw Error('missing abort'); } catch(e) { if(!(e instanceof SvgRenderAbortError)) throw e; console.log(e.code); }`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "svg-limit-exceeded:0\noperation-aborted");
  });

  await t.test("installed parser preserves an existing point in sparse scatter data", () => {
    const source = `import {createPresentation} from '@pptkit/core';import {generatePptx} from '@pptkit/pptx-exporter';import {importPptx} from '@pptkit/pptx-parser';const p=createPresentation();p.addSlide().addElement({type:'chart',chartType:'scatter',box:{x:0,y:0,width:400,height:300},blankDisplay:'gap',series:[{name:'S',xValues:[1,null,3],yValues:[4,5,6]}]});const s=await importPptx((await generatePptx(p)).bytes,{projection:'content',chartEditSaveProfile:'existing-values-v2'});const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);chart.series[0].xValues[2]=9;slide.replaceElement(chart.id,chart);const plan=await s.prepareSave({mode:'edit'});if(plan.status!=='ready')throw Error(JSON.stringify(plan.report));const r=await importPptx((await s.save(plan.plan)).bytes,{projection:'content'});console.log(JSON.stringify(r.document.slides[0].elements[0].series[0].xValues));r.dispose();s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "[1,null,9]");
  });

  await t.test("installed parser combines field projection, notes text and starting-number preservation", () => {
    const source = `import { readFile } from 'node:fs/promises'; import { importPptx, inspectPptx } from '@pptkit/pptx-parser'; const s = await importPptx(await readFile('field-text-source.pptx'), {projection:'content',fieldTextProfile:'date-time-and-number-v1',editSaveProfile:'text-runs-v2',firstSlideNumberSaveProfile:'attribute-v2'}); const slide=s.document.slides.find(x=>Array.isArray(x.notes)&&x.notes[0].runs.some(r=>r.field)); const notes=structuredClone(slide.notes); if(notes[0].runs[2].field.updateFormat!=='day-month-year'||notes[0].runs[3].field.updateFormat!=='time-12-seconds') throw Error('Wrong date/time format'); notes[0].runs[0].text='edited'; slide.setNotes(notes); s.document.setFirstSlideNumber(17); const p=await s.prepareSave({mode:'edit'}); if(p.status!=='ready'||p.plan.report.changes.length!==2) throw Error('combined plan failed'); const out=await s.save(p.plan); const n=await inspectPptx(out.bytes); console.log(s.report.fieldTextProfile+':'+n.firstSlideNumber); s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "date-time-and-number-v1:17");
  });
  await t.test("installed parser saves starting number and chart/workbook values together", () => {
    const source = `import { readFile } from 'node:fs/promises'; import { importPptx, inspectPptx } from '@pptkit/pptx-parser'; const options={projection:'content',chartEditSaveProfile:'existing-values-v1',firstSlideNumberSaveProfile:'attribute-v2'}; const s=await importPptx(await readFile('chart-value-source.pptx'),options); const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]); chart.series[0].values[1]=37; slide.replaceElement(chart.id,chart); s.document.setFirstSlideNumber(17); const p=await s.prepareSave({mode:'edit'}); if(p.status!=='ready'||p.plan.report.changes.length!==2) throw Error('combined plan failed'); const output=await s.save(p.plan); const check=await inspectPptx(output.bytes),r=await importPptx(output.bytes,options); console.log(check.firstSlideNumber+':'+r.document.slides[0].elements[0].series[0].values[1]); r.dispose(); s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "17:37");
  });
  await t.test("installed parser combines description, chart values and numbering", () => {
    const source = `import { readFile } from 'node:fs/promises'; import { importPptx } from '@pptkit/pptx-parser'; const options={projection:'content',editSaveProfile:'text-runs-v3',descriptionEditSaveProfile:'mixed-v5',drawingNameSaveProfile:'existing-v1',chartEditSaveProfile:'existing-values-v1',firstSlideNumberSaveProfile:'attribute-v2'};const s=await importPptx(await readFile('chart-value-source.pptx'),options),slide=s.document.slides[0],chart=structuredClone(slide.elements[0]);chart.series[0].values[1]=37;chart.name='Chart name';chart.accessibility={...chart.accessibility,description:'Chart alt'};slide.replaceElement(chart.id,chart);s.document.setFirstSlideNumber(17);const p=await s.prepareSave({mode:'edit'});if(p.status!=='ready'||p.plan.report.changes.length!==4)throw Error(JSON.stringify(p.report));const r=await importPptx((await s.save(p.plan)).bytes,options);console.log(r.document.firstSlideNumber+':'+r.document.slides[0].elements[0].series[0].values[1]+':'+r.document.slides[0].elements[0].accessibility.description+':'+r.document.slides[0].elements[0].name);r.dispose();s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "17:37:Chart alt:Chart name");
  });
  await t.test("installed parser combines field, adjacent text, chart and starting number", () => {
    const source = `import { readFile } from 'node:fs/promises'; import { importPptx } from '@pptkit/pptx-parser'; const options={projection:'content',chartEditSaveProfile:'existing-values-v1',firstSlideNumberSaveProfile:'attribute-v2',editSaveProfile:'text-runs-v3',fieldTextProfile:'slide-number-v1'}; const s=await importPptx(await readFile('field-chart-source.pptx'),options); const slide=s.document.slides[0],chart=structuredClone(slide.elements[0]),text=structuredClone(slide.elements[1]); chart.series[0].values[1]=7; text.content[0].runs[0].text='after'; slide.replaceElement(chart.id,chart); slide.replaceElement(text.id,text); s.document.setFirstSlideNumber(17); const p=await s.prepareSave({mode:'edit'}); if(p.status!=='ready'||p.plan.report.changes.length!==3) throw Error(JSON.stringify(p.report)); const bytes=(await s.save(p.plan)).bytes, r=await importPptx(bytes,options); const result=[r.document.firstSlideNumber,r.document.slides[0].elements[0].series[0].values[1],r.document.slides[0].elements[1].content[0].runs[1].text]; console.log(result.join(':')); r.dispose(); s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "17:7:42");
  });
  await t.test("installed Node file and byte entries preserve a shared picture replacement", () => {
    const output = JSON.parse(command(process.execPath, ["picture-save.mjs"], consumer));
    assert.equal(output.assurance, "scoped-preservation"); assert.ok(output.sourceBytes > 0);
    assert.notEqual(output.media[0], output.media[1]); assert.equal(output.media[1], "/media/raster.bin");
  });
  await t.test("installed Node file entry combines picture replacement and starting number", () => {
    const source = `import {readFile} from 'node:fs/promises'; import {importPptx} from '@pptkit/pptx-parser'; import {importPptxFile} from '@pptkit/pptx-parser/node'; const image=await readFile('picture-replacement.gif'); const options={projection:'content',pictureEditSaveProfile:'replace-v1',firstSlideNumberSaveProfile:'attribute-v2'}; const {session:s}=await importPptxFile('picture-source.pptx',options); const slide=s.document.slides[0],imageElement=slide.elements[0]; s.document.registerAsset({id:'new',kind:'image',source:{type:'bytes',value:image},mimeType:'image/gif',width:1,height:1}); slide.replaceElement(imageElement.id,{...imageElement,assetId:'new'}); s.document.setFirstSlideNumber(17); const p=await s.prepareSave({mode:'edit'}); if(p.status!=='ready'||p.plan.report.changes.length!==3) throw Error(JSON.stringify(p.report)); const r=await importPptx((await s.save(p.plan)).bytes,options); console.log(r.document.firstSlideNumber+':'+(r.report.pictureSources[0].mediaPart!==r.report.pictureSources[1].mediaPart)); r.dispose(); s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", source], consumer).trim(), "17:true");
  });
  await t.test("public declaration files compile with NodeNext and no workspace paths", () => {
    command(process.execPath, [path.join(repoRoot, "node_modules/typescript/bin/tsc"), "--project", path.join(consumer, "tsconfig.json")], consumer);
  });
  await t.test("npm installs the CLI executable", async () => {
    const binary = path.join(consumer, "node_modules/.bin", process.platform === "win32" ? "pptkit.cmd" : "pptkit");
    const result = process.platform === "win32" ? command("cmd.exe", ["/c", binary, "--help"], consumer) : command(binary, ["--help"], consumer);
    assert.match(result, /PPTKit/);
    const invoke = args => spawnSync(process.platform === "win32" ? "cmd.exe" : binary,
      process.platform === "win32" ? ["/c", binary, ...args] : args, { cwd: consumer, encoding: "utf8", timeout: 30000 });
    await writeFile(path.join(consumer, 'description-occurrences.json'), JSON.stringify({ version: 'pptkit-authoring-v1', layouts: [{ id: 'l', name: 'Layout', elements: [{ id: 't', type: 'table', box: { x: 0, y: 0, width: 100, height: 50 }, columns: [100], rows: [{ cells: [{ content: 'A' }] }] }] }], slides: [{ id: 's', layoutId: 'l' }, { id: 'hidden', layoutId: 'l', showInheritedShapes: false }] }));
    const occurrences = invoke(['validate', 'description-occurrences.json', '--check-descriptions', '--description-occurrences', '--fail-on-missing-description', '--json']);
    assert.equal(occurrences.status, 1, occurrences.stderr); const occurrenceReport = JSON.parse(occurrences.stdout).report.descriptions;
    assert.equal(occurrenceReport.scope, 'modeled-slide-occurrences'); assert.deepEqual(occurrenceReport.findings.map(item => [item.slideId, item.elementId]), [['s', 't']]);
    const inspected = invoke(["inspect", "smartart-text.pptx", "--diagram-text-comparison", "direct-paragraphs", "--json"]);
    assert.equal(inspected.status, 0, inspected.stderr); assert.equal(inspected.stderr, "");
    const inspection = JSON.parse(inspected.stdout); assert.equal(inspection.schemaVersion, "pptkit-cli-report-v1");
    assert.equal(inspection.report.inspection.diagramInventory.textComparisons[0].shapes[0].status, "equal");
    assert.equal(inspection.inputs[0].path, "smartart-text.pptx");
    const capabilities = invoke(["inspect", "smartart-text.pptx", "--capabilities", "--max-capability-subjects", "1", "--json"]);
    assert.equal(capabilities.status, 0, capabilities.stderr); const capabilityReport = JSON.parse(capabilities.stdout).report;
    assert.equal(capabilityReport.schemaVersion, "pptkit-import-capabilities-v1"); assert.equal(capabilityReport.truncated, true); assert.equal(capabilityReport.subjects[0].operations.length, 7);

    const equal = invoke(["compare", "smartart-text.pptx", "smartart-text.pptx", "--json"]);
    assert.equal(equal.status, 0, equal.stderr); assert.equal(JSON.parse(equal.stdout).report.equal, true);
    const different = invoke(["compare", "smartart-text.pptx", "smartart-data.pptx", "--max-differences", "1", "--json"]);
    assert.equal(different.status, 1); assert.equal(different.stderr, ""); assert.equal(JSON.parse(different.stdout).report.truncated, true);
    const missing = invoke(["inspect", "missing.pptx", "--json"]);
    assert.equal(missing.status, 3); assert.equal(missing.stdout, ""); assert.equal(JSON.parse(missing.stderr).error.code, "file-io");
    const saved = invoke(["save", "smartart-text.pptx", "cli-copy.pptx", "--mode", "no-op", "--json"]);
    assert.equal(saved.status, 0, saved.stderr); assert.equal(JSON.parse(saved.stdout).report.status, "committed");
    assert.deepEqual(await readFile(path.join(consumer, "cli-copy.pptx")), await readFile(path.join(consumer, "smartart-text.pptx")));
    const repeated = invoke(["save", "smartart-text.pptx", "cli-copy.pptx", "--mode", "no-op", "--json"]);
    assert.equal(repeated.status, 3); assert.equal(JSON.parse(repeated.stderr).error.code, "file-target-exists");
    const number = invoke(["save", "smartart-text.pptx", "cli-number.pptx", "--mode", "edit", "--first-slide-number", "17", "--json"]);
    assert.equal(number.status, 0, number.stderr); const report = JSON.parse(number.stdout);
    assert.equal(report.edit.firstSlideNumber, 17); assert.equal(report.report.assurance, "scoped-preservation");
    const inspectedNumber = invoke(["inspect", "cli-number.pptx", "--json"]);
    assert.equal(inspectedNumber.status, 0, inspectedNumber.stderr); assert.equal(JSON.parse(inspectedNumber.stdout).report.inspection.firstSlideNumber, 17);
    await writeFile(path.join(consumer, "cli-authoring.json"), JSON.stringify({ version: "pptkit-authoring-v1", slides: [{ elements: [{ type: "text", box: { x: 20, y: 20, width: 200, height: 40 }, content: "Installed generation" }] }] }));
    const validation = invoke(["validate", "cli-authoring.json", "--json"]);
    assert.equal(validation.status, 0, validation.stderr); assert.equal(JSON.parse(validation.stdout).report.scope, "core-authoring");
    const exportValidation = invoke(["validate", "cli-authoring.json", "--export", "--json"]);
    assert.equal(exportValidation.status, 0, exportValidation.stderr); assert.equal(JSON.parse(exportValidation.stdout).report.export.status, "checked"); assert.equal(JSON.parse(exportValidation.stdout).report.export.slideCount, 1);
    await writeFile(path.join(consumer, "cli-bad-image.json"), JSON.stringify({ version: "pptkit-authoring-v1", assets: [{ id: "img", kind: "image", mimeType: "image/png", source: { type: "bytes", encoding: "base64", value: Buffer.from("bad image").toString("base64") }, width: 8, height: 4 }], slides: [{ elements: [{ type: "image", assetId: "img", box: { x: 0, y: 0, width: 80, height: 40 } }] }] }));
    const strictValidation = invoke(["validate", "cli-bad-image.json", "--export", "--fail-on-export-warning", "--json"]);
    assert.equal(strictValidation.status, 1); assert.ok(JSON.parse(strictValidation.stdout).report.export.warnings.length);
    const strictGeneration = invoke(["generate", "cli-bad-image.json", "cli-rejected.pptx", "--fail-on-export-warning", "--json"]);
    assert.equal(strictGeneration.status, 3); assert.equal(JSON.parse(strictGeneration.stderr).error.code, "export-warnings-rejected"); await assert.rejects(stat(path.join(consumer, "cli-rejected.pptx")), { code: "ENOENT" });
    await mkdir(path.join(consumer, "asset folder")); await writeFile(path.join(consumer, "asset folder", "image.png"), png);
    await writeFile(path.join(consumer, "cli-local.json"), JSON.stringify({ version: "pptkit-authoring-v1", assets: [{ id: "img", kind: "image", mimeType: "image/png", source: { type: "path", value: "image.png" }, width: 8, height: 4 }], slides: [{ elements: [{ type: "image", assetId: "img", box: { x: 0, y: 0, width: 80, height: 40 } }] }] }));
    const local = invoke(["generate", "cli-local.json", "cli-local.pptx", "--asset-base", "asset folder", "--fail-on-export-warning", "--json"]);
    assert.equal(local.status, 0, local.stderr); assert.equal(JSON.parse(local.stdout).assetInputs[0].path, "asset folder/image.png"); assert.equal(JSON.parse(local.stdout).assetInputs[0].byteLength, png.length);
    const preview = invoke(["preview", "cli-authoring.json", "--slide", "1", "--json"]);
    assert.equal(preview.status, 0, preview.stderr); assert.match(JSON.parse(preview.stdout).report.slide.svg, /Installed generation/);
    const jpeg12 = await readFile(path.join(repoRoot, 'tests/fixtures/jpeg-12bit-progressive.jpg'));
    await writeFile(path.join(consumer, 'cli-jpeg12.json'), JSON.stringify({ version: 'pptkit-authoring-v1', assets: [{ id: 'jpeg12', kind: 'image', mimeType: 'image/jpeg', source: { type: 'bytes', encoding: 'base64', value: jpeg12.toString('base64') } }], slides: [{ id: 'slide', elements: [{ type: 'image', assetId: 'jpeg12', box: { x: 0, y: 0, width: 80, height: 40 } }] }] }));
    const jpegPreview = invoke(['preview', 'cli-jpeg12.json', '--slide', '1', '--json']);
    assert.equal(jpegPreview.status, 0, jpegPreview.stderr);
    assert.deepEqual(JSON.parse(jpegPreview.stdout).report.warnings.map(w => [w.code, w.assetId]), [['jpeg-preview-precision-unverified', 'jpeg12']]);
    assert.ok(JSON.parse(jpegPreview.stdout).report.slide.svg.includes(jpeg12.toString('base64')));
    const jpegGenerate = invoke(['generate', 'cli-jpeg12.json', 'cli-jpeg12.pptx', '--json']); assert.equal(jpegGenerate.status, 0, jpegGenerate.stderr);
    const jpegImported = invoke(['preview', 'cli-jpeg12.pptx', '--input-format', 'pptx', '--slide', '1', '--json']);
    assert.equal(jpegImported.status, 0, jpegImported.stderr);
    assert.ok(JSON.parse(jpegImported.stdout).report.warnings.some(w => w.code === 'jpeg-preview-precision-unverified'));
    const svgFile = invoke(["preview", "cli-authoring.json", "--slide", "1", "--output", "installed preview.svg", "--json"]);
    assert.equal(svgFile.status, 0, svgFile.stderr); assert.equal(JSON.parse(svgFile.stdout).report.publication.status, "committed"); assert.equal(await readFile(path.join(consumer, "installed preview.svg"), "utf8"), JSON.parse(preview.stdout).report.slide.svg);
    const allPreview = invoke(["preview", "cli-authoring.json", "--slide", "all", "--json"]);
    assert.equal(allPreview.status, 0, allPreview.stderr); assert.equal(JSON.parse(allPreview.stdout).report.slides.length, 1);
    const allHtml = invoke(["preview", "cli-authoring.json", "--slide", "all", "--output", "installed deck.html", "--json"]);
    assert.equal(allHtml.status, 0, allHtml.stderr); assert.equal(JSON.parse(allHtml.stdout).report.publication.format, "html"); assert.match(await readFile(path.join(consumer, "installed deck.html"), "utf8"), /<iframe sandbox=""/);
    const generated = invoke(["generate", "cli-authoring.json", "cli-generated.pptx", "--json"]);
    assert.equal(generated.status, 0, generated.stderr); assert.equal(JSON.parse(generated.stdout).report.slideCount, 1);
    const notesInventory = invoke(["inspect", "cli-generated.pptx", "--notes-inventory", "--json"]);
    assert.equal(notesInventory.status, 0, notesInventory.stderr);
    assert.equal(JSON.parse(notesInventory.stdout).notesInventory.schemaVersion, 'pptkit-notes-source-inventory-v1');
    assert.equal(JSON.parse(notesInventory.stdout).notesInventory.fieldEvaluation, 'not-performed');
    const notesHuman = invoke(["inspect", "cli-generated.pptx", "--notes-inventory"]);
    assert.equal(notesHuman.status, 0, notesHuman.stderr); assert.match(notesHuman.stdout, /Notes source inventory:/);
    const notesBudget = invoke(["inspect", "cli-generated.pptx", "--notes-inventory", "--max-notes-records", "100000000", "--json"]);
    assert.equal(notesBudget.status, 0, notesBudget.stderr);
    assert.deepEqual(JSON.parse(notesBudget.stdout).notesInventory, JSON.parse(notesInventory.stdout).notesInventory);
    const invalidNotesBudget = invoke(["inspect", "missing.pptx", "--notes-inventory", "--max-notes-records", "0", "--json"]);
    assert.equal(invalidNotesBudget.status, 2);
    await writeFile(path.join(consumer, "cli-decorative-authoring.json"), JSON.stringify({ version: "pptkit-authoring-v1", slides: [{ elements: [{ type: "shape", shape: "rect", box: { x: 0, y: 0, width: 100, height: 50 }, accessibility: { decorative: true, description: "Alt" } }] }] }));
    const decorativeGenerate = invoke(["generate", "cli-decorative-authoring.json", "cli-decorative-source.pptx", "--json"]); assert.equal(decorativeGenerate.status, 0, decorativeGenerate.stderr);
    const decorativeInspect = invoke(["inspect", "cli-decorative-source.pptx", "--projection", "content", "--decorative-values", "--json"]); assert.equal(decorativeInspect.status, 0, decorativeInspect.stderr);
    const decorative = JSON.parse(decorativeInspect.stdout).decorativeEditCandidates; assert.equal(decorative.targets[0].expected, true);
    await writeFile(path.join(consumer, "cli-decorative.json"), JSON.stringify({ version: decorative.version, sourceSha256: decorative.sourceSha256, edits: [{ ...decorative.targets[0], replacement: false }] }));
    const decorativeSave = invoke(["save", "cli-decorative-source.pptx", "cli-decorative-output.pptx", "--mode", "edit", "--decorative-edits", "cli-decorative.json", "--json"]); assert.equal(decorativeSave.status, 0, decorativeSave.stderr);
    const decorativeReopen = invoke(["inspect", "cli-decorative-output.pptx", "--projection", "content", "--decorative-values", "--json"]); assert.equal(decorativeReopen.status, 0, decorativeReopen.stderr); assert.equal(JSON.parse(decorativeReopen.stdout).decorativeEditCandidates.targets[0].expected, false);
    const visibilityInspection = invoke(["inspect", "cli-generated.pptx", "--projection", "content", "--visibility-values", "--json"]); assert.equal(visibilityInspection.status, 0, visibilityInspection.stderr);
    const visibility = JSON.parse(visibilityInspection.stdout).visibilityEditCandidates; assert.equal(visibility.targets[0].expected, false);
    await writeFile(path.join(consumer, "cli-visibility.json"), JSON.stringify({ version: visibility.version, sourceSha256: visibility.sourceSha256, edits: [{ ...visibility.targets[0], replacement: true }] }));
    const visibilitySave = invoke(["save", "cli-generated.pptx", "cli-hidden.pptx", "--mode", "edit", "--visibility-edits", "cli-visibility.json", "--json"]); assert.equal(visibilitySave.status, 0, visibilitySave.stderr);
    const visibilityReopen = invoke(["inspect", "cli-hidden.pptx", "--projection", "content", "--visibility-values", "--json"]); assert.equal(visibilityReopen.status, 0, visibilityReopen.stderr); assert.equal(JSON.parse(visibilityReopen.stdout).visibilityEditCandidates.targets[0].expected, true);
    const metadataInspection = invoke(["inspect", "cli-generated.pptx", "--projection", "content", "--metadata-values", "--json"]); assert.equal(metadataInspection.status, 0, metadataInspection.stderr);
    const metadata = JSON.parse(metadataInspection.stdout).metadataEditCandidates, title = metadata.targets.find(item => item.field === "title"); assert.ok(title);
    await writeFile(path.join(consumer, "cli-metadata.json"), JSON.stringify({ version: metadata.version, sourceSha256: metadata.sourceSha256, edits: [{ ...title, replacement: "Installed metadata" }] }));
    const metadataSave = invoke(["save", "cli-generated.pptx", "cli-metadata.pptx", "--mode", "edit", "--metadata-edits", "cli-metadata.json", "--json"]); assert.equal(metadataSave.status, 0, metadataSave.stderr);
    const metadataReopen = invoke(["inspect", "cli-metadata.pptx", "--projection", "content", "--metadata-values", "--json"]); assert.equal(metadataReopen.status, 0, metadataReopen.stderr); assert.equal(JSON.parse(metadataReopen.stdout).metadataEditCandidates.targets.find(item => item.field === "title").expected, "Installed metadata");
    const descriptionInspection = invoke(["inspect", "cli-generated.pptx", "--projection", "content", "--description-values", "--name-values", "--json"]);
    assert.equal(descriptionInspection.status, 0, descriptionInspection.stderr); const description = JSON.parse(descriptionInspection.stdout).descriptionEditCandidates;
    assert.equal(description.targets.length, 1); await writeFile(path.join(consumer, "cli-description.json"), JSON.stringify({ version: description.version, sourceSha256: description.sourceSha256, edits: [{ ...description.targets[0], replacement: "Installed alt" }] }));
    const names = JSON.parse(descriptionInspection.stdout).nameEditCandidates; assert.equal(names.targets.length, 1); await writeFile(path.join(consumer, "cli-names.json"), JSON.stringify({ version: names.version, sourceSha256: names.sourceSha256, edits: [{ ...names.targets[0], replacement: "Installed name" }] }));
    const descriptionSave = invoke(["save", "cli-generated.pptx", "cli-description.pptx", "--mode", "edit", "--description-edits", "cli-description.json", "--name-edits", "cli-names.json", "--first-slide-number", "17", "--json"]);
    assert.equal(descriptionSave.status, 0, descriptionSave.stderr); assert.equal(JSON.parse(descriptionSave.stdout).report.changes.length, 3);
    const descriptionReopen = invoke(["inspect", "cli-description.pptx", "--projection", "content", "--description-values", "--name-values", "--json"]);
    assert.equal(descriptionReopen.status, 0, descriptionReopen.stderr); assert.equal(JSON.parse(descriptionReopen.stdout).descriptionEditCandidates.targets[0].expected, "Installed alt"); assert.equal(JSON.parse(descriptionReopen.stdout).nameEditCandidates.targets[0].expected, "Installed name");

    const origins = invoke(["inspect", "cli-generated.pptx", "--projection", "content", "--edit-save-profile", "text-runs-v1", "--json"]);
    assert.equal(origins.status, 0, origins.stderr); const originReport = JSON.parse(origins.stdout), origin = originReport.report.textRuns.find(item => item.status === "bound"); assert.ok(origin);
    await writeFile(path.join(consumer, "cli-text-edits.json"), JSON.stringify({ version: "pptkit-text-edits-v1", sourceSha256: originReport.inputs[0].sha256, edits: [{ slideId: origin.slideId, elementId: origin.elementId, paragraphIndex: origin.paragraphIndex, runIndex: origin.runIndex, expected: "Installed generation", replacement: "Installed edited" }] }));
    const textDryRun = invoke(["save", "cli-generated.pptx", "cli-text-saved.pptx", "--mode", "edit", "--text-edits", "cli-text-edits.json", "--dry-run", "--json"]);
    assert.equal(textDryRun.status, 0, textDryRun.stderr); assert.equal(JSON.parse(textDryRun.stdout).status, "ready"); assert.equal(JSON.parse(textDryRun.stdout).report.changes[0].text.after, "Installed edited"); await assert.rejects(stat(path.join(consumer, "cli-text-saved.pptx")), { code: "ENOENT" });
    const textSaved = invoke(["save", "cli-generated.pptx", "cli-text-saved.pptx", "--mode", "edit", "--text-edits", "cli-text-edits.json", "--json"]);
    assert.equal(textSaved.status, 0, textSaved.stderr); assert.equal(JSON.parse(textSaved.stdout).report.assurance, "scoped-preservation");
    const checkText = `import {readFile} from 'node:fs/promises'; import {importPptx} from '@pptkit/pptx-parser'; const s=await importPptx(await readFile('cli-text-saved.pptx'),{projection:'content'}); console.log(s.document.slides[0].elements[0].content[0].runs[0].text); s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", checkText], consumer).trim(), "Installed edited");
    await writeFile(path.join(consumer, "cli-group.json"), JSON.stringify({ version: "pptkit-authoring-v1", slides: [{ elements: [{ id: "group", type: "group", box: { x: 0, y: 0, width: 200, height: 100 }, coordinateSize: { width: 200, height: 100 }, children: [{ id: "child", type: "text", content: "Grouped original", box: { x: 10, y: 10, width: 100, height: 40 } }] }] }] }));
    const groupGenerated = invoke(["generate", "cli-group.json", "cli-group.pptx", "--json"]); assert.equal(groupGenerated.status, 0, groupGenerated.stderr);
    const groupInspection = invoke(["inspect", "cli-group.pptx", "--projection", "content", "--edit-save-profile", "text-runs-v1", "--json"]); assert.equal(groupInspection.status, 0, groupInspection.stderr);
    const groupReport = JSON.parse(groupInspection.stdout), groupOrigin = groupReport.report.textRuns.find(item => item.status === "bound"); assert.ok(groupOrigin);
    await writeFile(path.join(consumer, "cli-group-edits.json"), JSON.stringify({ version: "pptkit-text-edits-v1", sourceSha256: groupReport.inputs[0].sha256, edits: [{ slideId: groupOrigin.slideId, elementId: groupOrigin.elementId, paragraphIndex: groupOrigin.paragraphIndex, runIndex: groupOrigin.runIndex, expected: "Grouped original", replacement: "Grouped edited" }] }));
    const groupSaved = invoke(["save", "cli-group.pptx", "cli-group-saved.pptx", "--mode", "edit", "--text-edits", "cli-group-edits.json", "--json"]); assert.equal(groupSaved.status, 0, groupSaved.stderr);
    const checkGroup = `import {readFile} from 'node:fs/promises'; import {importPptx} from '@pptkit/pptx-parser'; const s=await importPptx(await readFile('cli-group-saved.pptx'),{projection:'content'}); console.log(s.document.slides[0].elements[0].children[0].content[0].runs[0].text); s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", checkGroup], consumer).trim(), "Grouped edited");
    await writeFile(path.join(consumer, "cli-bad-image.json"), JSON.stringify({version: "pptkit-authoring-v1", assets: [{id: "bad", kind: "image", source: {type: "bytes", encoding: "base64", value: "iVBORw0KGgo="}}], slides: [{elements: [{type: "image", assetId: "bad", box: {x: 10, y: 10, width: 50, height: 50}}]}]}));
    const checkedImage = invoke(["validate", "cli-bad-image.json", "--export", "--image-diagnostics", "--json"]); assert.equal(checkedImage.status, 0, checkedImage.stderr); assert.ok(JSON.parse(checkedImage.stdout).report.export.warnings.some(item => item.code === "png-inspection-incomplete"));
    const rejectImage = invoke(["generate", "cli-bad-image.json", "rejected-image.pptx", "--image-diagnostics", "--fail-on-export-warning", "--json"]); assert.equal(rejectImage.status, 3); assert.equal(JSON.parse(rejectImage.stderr).error.code, "export-warnings-rejected"); await assert.rejects(readFile(path.join(consumer, "rejected-image.pptx")), {code: "ENOENT"});
    await writeFile(path.join(consumer, "cli-round-path.json"), JSON.stringify({version: "pptkit-authoring-v1", slides: [{elements: [{type: "connector", start: {x: 20, y: 20}, route: [{x: 80, y: 20}, {x: 80, y: 60}], end: {x: 20, y: 20}, pathMode: "closed-round", style: {paint: {type: "solid", color: "FF0000", opacity: .5}, width: 3}}]}]}));
    const roundPathGenerated = invoke(["generate", "cli-round-path.json", "cli-round-path.pptx", "--json"]); assert.equal(roundPathGenerated.status, 0, roundPathGenerated.stderr);
    const roundPathPreview = invoke(["preview", "cli-round-path.pptx", "--input-format", "pptx", "--slide", "1", "--json"]); assert.equal(roundPathPreview.status, 0, roundPathPreview.stderr); assert.match(JSON.parse(roundPathPreview.stdout).report.slide.svg, /<polygon /);
    const checkRoundPath = `import {readFile} from 'node:fs/promises';import {importPptx} from '@pptkit/pptx-parser';const s=await importPptx(await readFile('cli-round-path.pptx'),{projection:'content'});const e=s.document.slides[0].elements[0];console.log(JSON.stringify([e.pathMode,e.start,e.end,e.style.paint.opacity]));s.dispose();`;
    assert.deepEqual(JSON.parse(command(process.execPath, ["--input-type=module", "-e", checkRoundPath], consumer)), ["closed-round", {x: 20, y: 20}, {x: 20, y: 20}, .5]);
    await writeFile(path.join(consumer, "cli-v2.json"), JSON.stringify({ version: "pptkit-authoring-v1", slides: [{ notes: "Notes original", elements: [{ id: "shape", type: "shape", shape: "rect", box: { x: 0, y: 0, width: 100, height: 50 }, text: { content: "Shape original" } }] }] }));
    const v2Generated = invoke(["generate", "cli-v2.json", "cli-v2.pptx", "--json"]); assert.equal(v2Generated.status, 0, v2Generated.stderr);
    const v2Inspect = invoke(["inspect", "cli-v2.pptx", "--projection", "content", "--edit-save-profile", "text-runs-v2", "--json"]); assert.equal(v2Inspect.status, 0, v2Inspect.stderr);
    const v2Report = JSON.parse(v2Inspect.stdout), v2Shape = v2Report.report.textRuns.find(item => item.owner === "shape-text" && item.status === "bound"), v2Notes = v2Report.report.notesTextRuns.find(item => item.status === "bound"); assert.ok(v2Shape && v2Notes);
    await writeFile(path.join(consumer, "cli-v2-edits.json"), JSON.stringify({ version: "pptkit-text-edits-v2", sourceSha256: v2Report.inputs[0].sha256, edits: [{ owner: "shape-text", slideId: v2Shape.slideId, elementId: v2Shape.elementId, paragraphIndex: v2Shape.paragraphIndex, runIndex: v2Shape.runIndex, expected: "Shape original", replacement: "Shape edited" }, { owner: "notes", slideId: v2Notes.slideId, paragraphIndex: v2Notes.paragraphIndex, runIndex: v2Notes.runIndex, expected: "Notes original", replacement: "Notes edited" }] }));
    const v2Saved = invoke(["save", "cli-v2.pptx", "cli-v2-saved.pptx", "--mode", "edit", "--text-edits", "cli-v2-edits.json", "--json"]); assert.equal(v2Saved.status, 0, v2Saved.stderr);
    const checkV2 = `import {readFile} from 'node:fs/promises'; import {importPptx} from '@pptkit/pptx-parser'; const s=await importPptx(await readFile('cli-v2-saved.pptx'),{projection:'content'}); console.log(s.document.slides[0].elements[0].text.content[0].runs[0].text+':'+s.document.slides[0].notes[0].runs[0].text); s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", checkV2], consumer).trim(), "Shape edited:Notes edited");
    await writeFile(path.join(consumer, "cli-v3.json"), JSON.stringify({ version: "pptkit-authoring-v1", slides: [{ elements: [{ type: "table", box: { x: 0, y: 0, width: 120, height: 40 }, columns: [60, 60], rows: [{ height: 40, cells: [{ content: "Cell original" }, { content: "Other" }] }] }] }] }));
    const v3Generated = invoke(["generate", "cli-v3.json", "cli-v3.pptx", "--json"]); assert.equal(v3Generated.status, 0, v3Generated.stderr);
    const v3Inspect = invoke(["inspect", "cli-v3.pptx", "--projection", "content", "--edit-save-profile", "text-runs-v3", "--text-values", "--json"]); assert.equal(v3Inspect.status, 0, v3Inspect.stderr);
    const v3Report = JSON.parse(v3Inspect.stdout), v3Candidates = v3Report.textEditCandidates, v3Cell = v3Candidates.targets.find(item => item.owner === "table-cell" && item.rowIndex === 0 && item.cellIndex === 0); assert.equal(v3Cell.expected, "Cell original");
    await writeFile(path.join(consumer, "cli-v3-edits.json"), JSON.stringify({ version: v3Candidates.version, sourceSha256: v3Candidates.sourceSha256, edits: [{ ...v3Cell, replacement: "Cell edited" }] }));
    const v3Saved = invoke(["save", "cli-v3.pptx", "cli-v3-saved.pptx", "--mode", "edit", "--text-edits", "cli-v3-edits.json", "--json"]); assert.equal(v3Saved.status, 0, v3Saved.stderr);
    const checkV3 = `import {readFile} from 'node:fs/promises'; import {importPptx} from '@pptkit/pptx-parser'; const s=await importPptx(await readFile('cli-v3-saved.pptx'),{projection:'content'}); console.log(s.document.slides[0].elements[0].rows[0].cells[0].content[0].runs[0].text); s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", checkV3], consumer).trim(), "Cell edited");
    await writeFile(path.join(consumer, "cli-chart.json"), JSON.stringify({ version: "pptkit-authoring-v1", slides: [{ elements: [{ type: "chart", chartType: "bar", box: { x: 10, y: 20, width: 100, height: 80 }, categories: ["A", "B"], series: [{ name: "Series", values: [1, 2] }] }] }] }));
    const chartGenerated = invoke(["generate", "cli-chart.json", "cli-chart.pptx", "--json"]); assert.equal(chartGenerated.status, 0, chartGenerated.stderr);
    const descriptions = invoke(["validate", "cli-chart.json", "--check-descriptions", "--fail-on-missing-description", "--json"]); assert.equal(descriptions.status, 1, descriptions.stderr);
    assert.equal(JSON.parse(descriptions.stdout).report.descriptions.findings[0].elementType, 'chart');
    await writeFile(path.join(consumer, 'cli-contrast.json'), JSON.stringify({ version: 'pptkit-authoring-v1', slides: [{ elements: [{ type: 'text', box: { x: 0, y: 0, width: 100, height: 50 }, content: [{ runs: [{ text: 'A', style: { color: '777777', fontSize: 12 } }] }] }] }] }));
    const contrast = invoke(['validate', 'cli-contrast.json', '--check-contrast', '--fail-on-contrast', '--json']); assert.equal(contrast.status, 1, contrast.stderr); assert.equal(JSON.parse(contrast.stdout).report.contrast.findings[0].status, 'below-threshold');
    const chartInspection = invoke(["inspect", "cli-chart.pptx", "--projection", "content", "--chart-edit-save-profile", "existing-values-v1", "--chart-values", "--json"]); assert.equal(chartInspection.status, 0, chartInspection.stderr);
    const chartCandidates = JSON.parse(chartInspection.stdout).chartEditCandidates, chartTarget = chartCandidates.targets.find(item => item.seriesIndex === 0 && item.pointIndex === 1); assert.equal(chartTarget.expected, "2");
    await writeFile(path.join(consumer, "cli-chart-edits.json"), JSON.stringify({ version: chartCandidates.version, sourceSha256: chartCandidates.sourceSha256, edits: [{ ...chartTarget, replacement: "7.5" }] }));
    const chartSaved = invoke(["save", "cli-chart.pptx", "cli-chart-saved.pptx", "--mode", "edit", "--chart-edits", "cli-chart-edits.json", "--json"]); assert.equal(chartSaved.status, 0, chartSaved.stderr);
    const checkChart = `import {readFile} from 'node:fs/promises'; import {importPptx} from '@pptkit/pptx-parser'; const s=await importPptx(await readFile('cli-chart-saved.pptx'),{projection:'content',chartEditSaveProfile:'existing-values-v1'}); console.log(s.document.slides[0].elements[0].series[0].values[1]+':'+s.report.chartValues[1].status); s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", checkChart], consumer).trim(), "7.5:bound");
    await writeFile(path.join(consumer, 'cli-sparse-scatter.json'), JSON.stringify({ version: 'pptkit-authoring-v1', slides: [{ elements: [{ type: 'chart', chartType: 'scatter', box: { x: 10, y: 20, width: 200, height: 100 }, blankDisplay: 'gap', series: [{ name: 'S', xValues: [1, null, 3], yValues: [4, 5, 6] }] }] }] }));
    const sparseGenerated = invoke(['generate', 'cli-sparse-scatter.json', 'cli-sparse-scatter.pptx', '--json']); assert.equal(sparseGenerated.status, 0, sparseGenerated.stderr);
    const sparseInspection = invoke(['inspect', 'cli-sparse-scatter.pptx', '--projection', 'content', '--chart-edit-save-profile', 'existing-values-v2', '--chart-values', '--json']); assert.equal(sparseInspection.status, 0, sparseInspection.stderr);
    const sparse = JSON.parse(sparseInspection.stdout).chartEditCandidatesV3, sparseTarget = sparse.targets.find(target => target.coordinate === 'x' && target.pointIndex === 2); assert.ok(sparseTarget);
    await writeFile(path.join(consumer, 'cli-sparse-edits.json'), JSON.stringify({ version: sparse.version, sourceSha256: sparse.sourceSha256, edits: [{ ...sparseTarget, replacement: '9' }] }));
    const sparseSaved = invoke(['save', 'cli-sparse-scatter.pptx', 'cli-sparse-saved.pptx', '--mode', 'edit', '--chart-edits', 'cli-sparse-edits.json', '--json']); assert.equal(sparseSaved.status, 0, sparseSaved.stderr);
    const checkSparse = `import {readFile} from 'node:fs/promises';import {importPptx} from '@pptkit/pptx-parser';const s=await importPptx(await readFile('cli-sparse-saved.pptx'),{projection:'content'});console.log(JSON.stringify(s.document.slides[0].elements[0].series[0].xValues));s.dispose();`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', checkSparse], consumer).trim(), '[1,null,9]');
    const pictureInspection = invoke(["inspect", "picture-source.pptx", "--projection", "content", "--picture-source-profile", "embedded-v1", "--json"]); assert.equal(pictureInspection.status, 0, pictureInspection.stderr);
    const packageSummary = invoke(["inspect", "picture-source.pptx", "--package-summary", "--json"]); assert.equal(packageSummary.status, 0, packageSummary.stderr);
    const summary = JSON.parse(packageSummary.stdout).packageSummary; assert.equal(summary.schemaVersion, 'pptkit-cli-package-summary-v1'); assert.ok(summary.media.some(part => part.contentType === 'image/png' && part.incomingRelationships > 0));
    const pictureReport = JSON.parse(pictureInspection.stdout), pictureOrigin = pictureReport.report.pictureSources[0];
    await writeFile(path.join(consumer, "cli-picture.json"), JSON.stringify({ version: "pptkit-picture-edit-v1", sourceSha256: pictureReport.inputs[0].sha256, replacementSha256: createHash('sha256').update(gif).digest('hex'), slideId: pictureOrigin.slideId, elementId: pictureOrigin.elementId, expectedAssetId: pictureOrigin.assetId, mimeType: 'image/gif', width: 1, height: 1 }));
    const pictureSaved = invoke(["save", "picture-source.pptx", "cli-picture-saved.pptx", "--mode", "edit", "--picture-edit", "cli-picture.json", "--picture-bytes", "picture-replacement.gif", "--json"]); assert.equal(pictureSaved.status, 0, pictureSaved.stderr);
    const checkPicture = `import {readFile} from 'node:fs/promises'; import {importPptx} from '@pptkit/pptx-parser'; const s=await importPptx(await readFile('cli-picture-saved.pptx'),{projection:'content',pictureSourceProfile:'embedded-v1'}); console.log(s.report.pictureSources[0].mediaPart!==s.report.pictureSources[1].mediaPart); s.dispose();`;
    assert.equal(command(process.execPath, ["--input-type=module", "-e", checkPicture], consumer).trim(), "true");
    const inspectGenerated = invoke(["inspect", "cli-generated.pptx", "--projection", "content", "--json"]);
    assert.equal(inspectGenerated.status, 0, inspectGenerated.stderr); assert.equal(JSON.parse(inspectGenerated.stdout).report.inspection.slides.length, 1);
    const pptxPreview = invoke(["preview", "cli-generated.pptx", "--input-format", "pptx", "--slide", "1", "--json"]);
    assert.equal(pptxPreview.status, 0, pptxPreview.stderr); const projected = JSON.parse(pptxPreview.stdout).report;
    assert.match(projected.slide.svg, /Installed generation/); assert.equal(projected.importReport.projection, "content");
    const allPptx = invoke(["preview", "cli-generated.pptx", "--input-format", "pptx", "--slide", "all", "--json"]);
    assert.equal(allPptx.status, 0, allPptx.stderr); assert.equal(JSON.parse(allPptx.stdout).report.slides.length, 1);

  });
  await t.test("installed media entry resolves without a DOM in Node", () => {
    const code = `import { mountBrowserMedia } from '@pptkit/svg-renderer/media'; console.log(typeof mountBrowserMedia)`;
    assert.equal(command(process.execPath, ['--input-type=module', '-e', code], consumer).trim(), 'function');
  });
  await t.test("installed DOM export resolves its actual browser module graph and normalizes JPEG", async (t) => {
    const root = path.join(installed, "pptx-exporter");
    const manifest = JSON.parse(await readFile(path.join(root, "package.json"), "utf8"));
    const target = manifest.exports["./dom"].default;
    assert.match(target, /^\.\/dist\/.*\.js$/);
    const requests = [];
    const browser = await chromium.launch({ headless: true });
    t.after(() => browser.close());
    const page = await browser.newPage();
    await page.route("**/*", async route => {
      const url = new URL(route.request().url());
      requests.push(url.href);
      if (url.origin !== "https://installed.invalid" || !/^\/dist\/[a-zA-Z0-9_./-]+\.js$/.test(url.pathname) || url.pathname.includes("..")) return route.abort();
      try {
        const body = await readFile(path.join(root, url.pathname.slice(1)));
        await route.fulfill({ status: 200, contentType: "text/javascript", headers: { "access-control-allow-origin": "*" }, body });
      } catch { await route.abort(); }
    });
    const jpeg = await readFile(path.join(repoRoot, "tests/fixtures/exif-orientation6.jpg"));
    const png = await readFile(path.join(repoRoot, "tests/fixtures/exif-orientation6-normalized.png"));
    const pngIcc = await readFile(path.join(repoRoot, "tests/fixtures/png-icc-srgb.png"));
    const expectedIcc = (await import('../packages/pptx-exporter/dist/node.js')).extractPngIccProfile(pngIcc);
    const actual = await page.evaluate(async ({ entry, jpeg, png, pngIcc }) => {
      const { normalizeJpegOrientationInBrowser, orientJpegCrop, extractPngIccProfileInBrowser } = await import(entry);
      const crop = orientJpegCrop({ left: 0.1, top: 0.2, right: 0.3, bottom: 0.4 }, 6);
      const result = await normalizeJpegOrientationInBrowser(Uint8Array.from(jpeg));
      const icc = await extractPngIccProfileInBrowser(Uint8Array.from(pngIcc));
      async function pixels(bytes) {
        const image = await createImageBitmap(new Blob([Uint8Array.from(bytes)], { type: "image/png" }));
        const canvas = document.createElement("canvas"); canvas.width = image.width; canvas.height = image.height;
        const context = canvas.getContext("2d"); context.drawImage(image, 0, 0); image.close();
        return Array.from(context.getImageData(0, 0, canvas.width, canvas.height).data);
      }
      return { crop, width: result.width, height: result.height, orientation: result.sourceOrientation,
        mime: result.mimeType, declaration: result.orientationNormalization,
        output: await pixels(result.bytes), control: await pixels(png), icc: [...icc.profileBytes] };
    }, { entry: new URL(target, "https://installed.invalid/").href, jpeg: [...jpeg], png: [...png], pngIcc: [...pngIcc] });
    assert.deepEqual([actual.width, actual.height, actual.orientation, actual.mime, actual.declaration], [2, 3, 6, "image/png", "pixels"]);
    assert.deepEqual(actual.crop, { left: 0.4, top: 0.1, right: 0.2, bottom: 0.3 });
    assert.deepEqual(actual.output, actual.control);
    assert.deepEqual(actual.icc, [...expectedIcc.profileBytes]);
    assert.ok(requests.some(url => url.endsWith("/dist/browser/dom.js")));
    assert.ok(requests.every(url => url.startsWith("https://installed.invalid/dist/")));
  });
  await t.test("installed browser bundles run without Node or registry access", async (t) => {
    const browser = await chromium.launch({ headless: true });
    t.after(() => browser.close());
    const page = await browser.newPage();
    await page.route("**/*", (route) => route.abort());
    const imports = {};
    const globals = [];
    for (const name of ["core", "layout", "pptx-exporter", "svg-renderer", "pptx-parser"]) {
      const pkgRoot = path.join(installed, name);
      const manifest = JSON.parse(await readFile(path.join(pkgRoot, "package.json"), "utf8"));
      const source = await readFile(path.resolve(pkgRoot, manifest.exports["./browser"].default), "utf8");
      imports[`@pptkit/${name}`] = imports[`@pptkit/${name}/browser`] = dataModule(source);
      globals.push(await readFile(path.join(pkgRoot, "dist/browser/global.js"), "utf8"));
    }
    imports["@pptkit/layout/dom"] = dataModule(await readFile(path.join(installed, "layout/dist/browser/dom.js"), "utf8"));
    imports["@pptkit/pptx-exporter/dom"] = dataModule(await readFile(path.join(installed, "pptx-exporter/dist/browser/dom.js"), "utf8"));
    imports["@pptkit/svg-renderer/media"] = dataModule(await readFile(path.join(installed, "svg-renderer/dist/browser/media.js"), "utf8"));
    await page.setContent(`<script type="importmap">${JSON.stringify({ imports })}</script>`);
    const result = await page.evaluate(async ({ textFixtureBase64, graphFixtureBase64 }) => {
      const { mountBrowserMedia, resolveAndMountBrowserMedia, mountLayoutMedia } = await import("@pptkit/svg-renderer/media");
      const mediaHost = globalThis.document.createElement("div"); globalThis.document.body.append(mediaHost);
      const media = mountBrowserMedia(mediaHost, { kind: "audio", source: new Blob(["fixture"], { type: "audio/wav" }), contentType: "audio/wav" });
      media.dispose(); if (media.state.status !== "disposed" || mediaHost.children.length) throw new Error("Installed media adapter cleanup failed");
      const mediaEvents = [];
      const resolved = await resolveAndMountBrowserMedia(mediaHost, { kind: "audio", contentType: "audio/wav", maxBytes: 100, onResolveState: state => mediaEvents.push(state.status), resolve: () => new Blob(["fixture"], { type: "audio/wav" }) });
      resolved.dispose(); if (resolved.state.status !== "disposed" || mediaHost.children.length || mediaEvents.join(",") !== "loading,mounted") throw new Error("Installed media resolver cleanup failed");
      const { createPresentation, planImageReplacement } = await import("@pptkit/core/browser");
      const imagePlan = planImageReplacement({ type: "image", assetId: "old", fit: "crop", crop: { left: 0.2 }, box: { x: 1, y: 2, width: 100, height: 100 } }, { id: "new", kind: "image", source: { type: "url", value: "local" }, width: 300, height: 200 }, { fit: "cover" });
      if (imagePlan.replacement.assetId !== "new" || imagePlan.replacement.fit !== "cover" || imagePlan.replacement.crop !== undefined || imagePlan.previous.crop.left !== 0.2) throw new Error("Installed image replacement planning failed");
      const { resolveLayout, planVisibleImageReplacement, planNestedVisibleImageReplacement, planSharedVisibleImageReplacement } = await import("@pptkit/layout/browser");
      const mediaDeck = createPresentation();
      mediaDeck.registerAsset({ id: "installed-sound", kind: "audio", mimeType: "audio/wav", source: { type: "bytes", value: new Uint8Array([1, 2, 3]) } });
      mediaDeck.addSlide({ id: "media-slide", elements: [{ id: "media-object", type: "audio", audioAssetId: "installed-sound", box: { x: 0, y: 0, width: 30, height: 30 } }] });
      const mountedLayout = await mountLayoutMedia(mediaHost, resolveLayout(mediaDeck), { slideId: "media-slide", elementId: "media-object", maxBytes: 10 });
      if (mountedLayout.assetId !== "installed-sound") throw new Error("Installed layout media target failed");
      mountedLayout.controller.dispose(); if (mediaHost.children.length) throw new Error("Installed layout media cleanup failed");
      const visibleDeck = createPresentation();
      visibleDeck.registerAsset({ id: "old", kind: "image", source: { type: "url", value: "old" }, width: 400, height: 200 });
      visibleDeck.registerAsset({ id: "new", kind: "image", source: { type: "url", value: "new" }, width: 100, height: 300 });
      visibleDeck.addSlide({ id: "visible-slide", elements: [{ id: "visible-pic", type: "image", assetId: "old", fit: "contain", box: { x: 0, y: 0, width: 100, height: 100 } }] });
      const visiblePlan = planVisibleImageReplacement(visibleDeck, "visible-slide", "visible-pic", "new", "stretch", { composition: [{ type: "align", container: { kind: "slide", id: "visible-slide" }, elementIds: ["visible-pic"], area: { x: 200, y: 100, width: 200, height: 200 }, horizontal: "end", vertical: "center" }] });
      if (visiblePlan.replacement.box.x !== 300 || visiblePlan.replacement.box.y !== 175 || visiblePlan.replacement.box.height !== 50 || visiblePlan.previous.assetId !== "old") throw new Error("Installed visible image plan failed");
      visibleDeck.slides[0].addElement({ type: "group", id: "group", box: { x: 0, y: 0, width: 100, height: 100 }, coordinateSize: { width: 100, height: 100 }, children: [{ type: "image", id: "nested", assetId: "old", fit: "contain", box: { x: 0, y: 0, width: 100, height: 100 } }] });
      const nestedPlan = planNestedVisibleImageReplacement(visibleDeck, "visible-slide", "nested", "new", "stretch");
      if (nestedPlan.rootElementId !== "group" || nestedPlan.replacement.children[0].box.y !== 25 || nestedPlan.previous.children[0].assetId !== "old") throw new Error("Installed nested image plan failed");
      visibleDeck.defineSlideMaster({ id: "shared-master", name: "Shared", elements: [{ id: "master-image", type: "image", assetId: "old", fit: "contain", box: { x: 0, y: 0, width: 100, height: 100 } }] });
      const sharedPlan = planSharedVisibleImageReplacement(visibleDeck, { kind: "master", id: "shared-master" }, "master-image", "new", "stretch");
      if (sharedPlan.kind !== "master" || sharedPlan.replacement.elements[0].box.y !== 25 || sharedPlan.previous.elements[0].assetId !== "old") throw new Error("Installed shared image plan failed");
      const { generatePptx } = await import("@pptkit/pptx-exporter/browser");
      const { normalizeJpegOrientationInBrowser } = await import("@pptkit/pptx-exporter/dom");
      if (typeof normalizeJpegOrientationInBrowser !== "function") throw new Error("Installed JPEG normalization adapter missing");
      const { inspectZipArchive, inspectXmlPart, inspectOpcPackage, inspectPresentationPackage, inspectMarkupCompatibility, inspectPptx, inspectPptxNotes, inspectPptxMasterPictures, importPptx } = await import("@pptkit/pptx-parser/browser");
      if ((await inspectXmlPart(new TextEncoder().encode('<p:r xmlns:p="urn:installed"/>'))).root.namespaceUri !== "urn:installed") throw new Error("Installed XML namespace inspection failed");
      const { renderLayoutToSvg } = await import("@pptkit/svg-renderer/browser");
      const { createBrowserTextMeasurement } = await import("@pptkit/layout/dom");
      const adapter = createBrowserTextMeasurement({ document: globalThis.document, fonts: [] });
      if (adapter.resolveFont({ family: "Missing", fontSize: 20, bold: false, italic: false, language: "en", text: "A" }) !== undefined) throw new Error("Unregistered font accepted");
      adapter.dispose();
      const document = createPresentation({ masters: [{ id: "installed", name: "Installed", theme: { colors: { accent1: "ABCDEF" } } }] });
      document.defineSlideLayout({ id: "installed-layout", name: "Installed", masterId: "installed" });
      document.addSlide().addElement({ type: "text", content: "Installed browser", box: { x: 40, y: 40, width: 300, height: 80 } });
      const svg = await renderLayoutToSvg(resolveLayout(document));
      const pptx = await generatePptx(document);
      if ((await inspectPptx(pptx.bytes)).slideMasters.length !== 2) throw new Error("Installed browser master graph failed");
      const masterPictures = await inspectPptxMasterPictures(pptx.bytes);
      if (masterPictures.masters.length !== 2 || masterPictures.mapping !== "not-performed") throw new Error("Installed master picture inventory failed");
      if ((await inspectMarkupCompatibility(new TextEncoder().encode("<r/>"), { understoodNamespaces: [] })).output.elements !== 1) throw new Error("Installed browser MCE inspection failed");
      if ((await inspectPptx(pptx.bytes)).metadata.extended.properties.find((p) => p.name === "Application").value.value !== "PPTKit") throw new Error("Installed browser metadata inspection failed");
      if ((await inspectPptxNotes(pptx.bytes)).authoredMapping !== "not-performed") throw new Error("Installed notes inventory failed");
      const session = await importPptx(pptx.bytes, { projection: "structure" });
      const prepared = await session.prepareSave({ mode: "no-op" });
      if (prepared.status !== "ready") throw new Error("Installed session preparation blocked");
      const saved = await session.save(prepared.plan);
      if (saved.bytes.length !== pptx.bytes.length || !saved.bytes.every((b, i) => b === pptx.bytes[i])) throw new Error("Installed no-op bytes changed");
      session.dispose();
      if ((await inspectPptx(pptx.bytes)).slides.length !== 1) throw new Error("Installed browser structure inspection failed");
      if ((await inspectPresentationPackage(pptx.bytes)).dialect !== "transitional") throw new Error("Installed browser presentation classification failed");
      if (!(await inspectOpcPackage(pptx.bytes)).relationships.some((r) => r.targetPart === "/ppt/presentation.xml")) throw new Error("Installed browser OPC graph missing main part");
      if (!(await inspectZipArchive(pptx.bytes)).entries.some((entry) => entry.name === "ppt/presentation.xml")) throw new Error("Browser archive inventory missing presentation part");
      const textBytes = Uint8Array.from(atob(textFixtureBase64), char => char.charCodeAt(0));
      const textSession = await importPptx(textBytes, { projection: "content" });
      const inherited = textSession.document.slides[0].elements[0].text;
      const provenance = textSession.report.textStyleInheritance.properties;
      if (inherited.frame.margin.left !== 5 || inherited.content[0].runs[0].style.fontSize !== 20 || inherited.content[0].runs[0].style.color !== "13579B" || inherited.content[0].runs[0].style.fontFamily !== "Installed Theme Font") throw new Error("Installed browser text/theme inheritance projection failed");
      if (!provenance.some(property => property.property === "run.fontSize" && property.level === "layout-placeholder")) throw new Error("Installed browser text provenance missing");
      if (!provenance.some(property => property.property === "run.color" && property.resolvedFrom?.some(source => source.part === "/deck/theme.xml"))) throw new Error("Installed browser theme provenance missing");
      textSession.dispose();
      const graphBytes = Uint8Array.from(atob(graphFixtureBase64), char => char.charCodeAt(0));
      const graph = (await inspectPptx(graphBytes, { diagramLimits: { maxFindings: 100 } })).diagramInventory.dataModels[0];
      if (graph.points[1].text[0].runs[0].text !== "A & B 😀" || graph.connections[0].destination.targetElementIndex !== graph.points[1].elementIndex) throw new Error("Installed browser data graph failed");
      const mapped = (await inspectPptx(graphBytes)).diagramInventory.cacheMappings[0];
      if (mapped.shapes[0].model.targetElementIndex !== graph.points[5].elementIndex || graph.presentationMappings[0].targetElementIndex !== graph.points[1].elementIndex) throw new Error("Installed browser diagram mapping failed");
      return { svg: svg.slides[0].svg, status: pptx.status, signature: [...pptx.bytes.subarray(0, 2)], textSize: inherited.content[0].runs[0].style.fontSize, textProperties: provenance.length };
    }, { textFixtureBase64: Buffer.from(textFixture).toString("base64"), graphFixtureBase64: Buffer.from(graphFixture).toString("base64") });
    assert.match(result.svg, /Installed browser/);
    assert.equal(result.status, "generated");
    assert.deepEqual(result.signature, [80, 75]);
    assert.equal(result.textSize, 20);
    assert.ok(result.textProperties > 0);
    const globalPage = await browser.newPage();
    await globalPage.route("**/*", (route) => route.abort());
    for (const source of globals) await globalPage.addScriptTag({ content: source });
    assert.equal(await globalPage.evaluate(async graphFixtureBase64 => {
      if ((await PPTKit.inspectXmlPart(new TextEncoder().encode("<r/>"))).elements !== 1) throw new Error("Global XML inspection failed");
      const document = PPTKit.createPresentation({ masters: [{ id: "installed", name: "Installed", theme: { colors: { accent1: "ABCDEF" } } }] });
      document.defineSlideLayout({ id: "installed-layout", name: "Installed", masterId: "installed" });
      document.addSlide();
      const generated = await PPTKit.generatePptx(document);
      if ((await PPTKit.inspectPptx(generated.bytes)).slideMasters.length !== 2) throw new Error("Installed global master graph failed");
      if ((await PPTKit.inspectMarkupCompatibility(new TextEncoder().encode("<r/>"), { understoodNamespaces: [] })).output.elements !== 1) throw new Error("Global MCE inspection failed");
      if ((await PPTKit.inspectPptx(generated.bytes)).metadata.extended.properties.find((p) => p.name === "Application").value.value !== "PPTKit") throw new Error("Global metadata inspection failed");
      if ((await PPTKit.inspectPptxNotes(generated.bytes)).geometryPolicy !== "unresolved-source-only") throw new Error("Installed global notes inventory failed");
      const session = await PPTKit.importPptx(generated.bytes, { projection: "structure" });
      const prepared = await session.prepareSave({ mode: "no-op" });
      if (prepared.status !== "ready") throw new Error("Global session preparation blocked");
      const saved = await session.save(prepared.plan);
      if (saved.bytes.length !== generated.bytes.length || !saved.bytes.every((b, i) => b === generated.bytes[i])) throw new Error("Global no-op bytes changed");
      session.dispose();
      if ((await PPTKit.inspectPptx(generated.bytes)).slides.length !== 1) throw new Error("Global structure inspection failed");
      if ((await PPTKit.inspectPresentationPackage(generated.bytes)).suggestedExtension !== "pptx") throw new Error("Global presentation classification failed");
      if (!(await PPTKit.inspectOpcPackage(generated.bytes)).parts.length) throw new Error("Global OPC graph missing parts");
      if (!(await PPTKit.inspectZipArchive(generated.bytes)).entries.length) throw new Error("Global parser bundle returned no entries");
      const graphBytes = Uint8Array.from(atob(graphFixtureBase64), char => char.charCodeAt(0));
      const graph = (await PPTKit.inspectPptx(graphBytes)).diagramInventory.dataModels[0];
      if (graph.points.length !== 7 || graph.connections[0].source.status !== "resolved" || PPTKit.DEFAULT_DIAGRAM_LIMITS.maxFindings !== 10000) throw new Error("Installed global data graph failed");
      const mapped = (await PPTKit.inspectPptx(graphBytes)).diagramInventory.cacheMappings[0];
      if (mapped.status !== "inspected" || graph.presentationMappings[0].agreement !== "same-target") throw new Error("Installed global diagram mapping failed");
      return generated.slideCount;
    }, Buffer.from(graphFixture).toString("base64")), 1);
  });
});
