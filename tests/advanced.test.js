import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { performance } from 'node:perf_hooks';
import JSZip from 'jszip';
import { Client } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';
import { buildDeck } from '../src/build.js';
import { readPackage, parseXml, descendants, NS, checkNativeSafe } from '../src/ooxml.js';
import { indexPackage } from '../src/file-engine.js';
import { TaskHost, encodeTargetRef, decodeTargetRef } from '../src/task.js';
import { NativeHost } from '../src/native-host.js';
import { hash, writeJson } from '../src/storage.js';
import { listProcesses, createOwnedJob, sameProcess } from '../src/windows.js';
import { advancedDeck } from '../scripts/fixtures/advanced.js';
import { smartArtSeed } from '../scripts/fixtures/smartart-seed.js';

const rootFor = () => path.resolve('work/tests/advanced', randomUUID());
test('editable charts retain exact categories, data, workbooks and table geometry; invalid data is rejected', async () => {
  const deck = advancedDeck();
  deck.slides[0].items.push({ type:'table',name:'SmallTable',left:690,top:465,width:220,height:50,rows:[['A','B'],['1','2']] });
  const parts = await readPackage(await buildDeck(deck)), index = indexPackage(parts), charts = index.objects.filter(o => o.kind === 'chart');
  assert.equal(charts.length, 7); assert.deepEqual(charts[0].chart.types, ['barChart','lineChart']);
  assert.deepEqual(charts[0].chart.series[0].values, [18,24,31,39]);
  assert.deepEqual(charts[0].chart.series[1].values, [42,55,68,74]);
  assert.deepEqual(charts[0].chart.series[0].labels, ['Q1','Q2','Q3','Q4']);
  for (const object of charts) {
    assert.equal(object.chart.workbooks.length, 1);
    const workbook = await JSZip.loadAsync(parts.get(object.chart.workbooks[0].part));
    assert(workbook.file('xl/worksheets/sheet1.xml'));
  }
  await checkNativeSafe(parts);
  const workbookPart = charts[0].chart.workbooks[0].part;
  for (const active of ['formula', 'external', 'macro']) {
    const changed = new Map(parts), workbook = await JSZip.loadAsync(parts.get(workbookPart));
    if (active === 'macro') workbook.file('xl/vbaProject.bin',Buffer.from('synthetic'));
    if (active === 'formula') workbook.file('xl/worksheets/sheet1.xml',(await workbook.file('xl/worksheets/sheet1.xml').async('string')).replace('</worksheet>','<f>1+1</f></worksheet>'));
    if (active === 'external') workbook.file('xl/_rels/workbook.xml.rels',(await workbook.file('xl/_rels/workbook.xml.rels').async('string')).replace('</Relationships>','<Relationship Id="rIdUnsafe" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="https://example.invalid/data.xlsx" TargetMode="External"/></Relationships>'));
    changed.set(workbookPart,await workbook.generateAsync({type:'nodebuffer'}));
    await assert.rejects(checkNativeSafe(changed),{code:active==='external'?'NATIVE_EXTERNAL_CONTENT':'NATIVE_CONTENT_UNSUPPORTED'});
  }
  const table = index.objects.find(o => o.kind === 'table'); assert.equal(table.geometry.width,220); assert.equal(table.geometry.height,50);
  assert.equal(index.objects.find(o => o.name === 'Title').runs[0].style.color,'152D4A');
  const bad = structuredClone(deck); bad.slides[0].items[1].series[0].values.pop(); await assert.rejects(buildDeck(bad));
  const pie = { ...deck.slides[2].items[1], chartType:'pie',series:[{name:'Invalid',labels:['A'],values:[-1]}] };
  await assert.rejects(buildDeck({slides:[{items:[pie]}]}));
});

test('layout dry-run preserves content and revision; explicit apply is idempotent and rejects stale or duplicate references', async () => {
  const root = rootFor(), host = new TaskHost({baseDir:path.join(root,'tasks')});
  const deck = advancedDeck(), built = await host.build({deck:{theme:'executive',slides:[deck.slides.at(-1)]},operationId:'build'});
  const snapshot = await host.inspect({documentId:built.documentId});
  const foreign = encodeTargetRef({ ...decodeTargetRef(snapshot.objects.find(o=>o.name==='Title').targetRef), taskId:randomUUID() });
  await assert.rejects(host.apply({documentId:built.documentId,expectedRevision:0,operationId:'foreign_task',operations:[{type:'set_geometry',targetRef:foreign,geometry:{left:20}}]}),{code:'CROSS_TASK_REFERENCE'});
  const other = new TaskHost({baseDir:path.join(root,'tasks')});
  const otherBuilt = await other.build({deck:{slides:[deck.slides.at(-1)]},operationId:'build'});
  await assert.rejects(other.apply({documentId:otherBuilt.documentId,expectedRevision:0,operationId:'foreign_document',operations:[{type:'set_geometry',targetRef:snapshot.objects.find(o=>o.name==='Title').targetRef,geometry:{left:20}}]}),{code:'CROSS_TASK_REFERENCE'});
  assert.equal((await other.inspect({documentId:otherBuilt.documentId})).revision,0);
  const bindings = (await other.status()).bindings;
  assert.equal(bindings[0].taskId,other.taskId);assert.equal(bindings[0].documentId,otherBuilt.documentId);
  await other.finish({reviewIds:[]});
  const args = {documentId:built.documentId,expectedRevision:0,operationId:'plan',slide:1,theme:'executive',layout:'grid',titleRef:snapshot.objects.find(o=>o.name==='Title').targetRef,
    targetRefs:snapshot.objects.filter(o=>o.capabilities.includes('set_geometry')&&o.name!=='Title').map(o=>o.targetRef)};
  const plan = await host.relayout(args); assert.equal(plan.outcome,'layout_planned'); assert.equal((await host.inspect({documentId:built.documentId})).revision,0);
  const cli = spawnSync(process.execPath,[path.resolve('src/cli.js'),'relayout','--task',host.taskId,'--base-dir',host.baseDir,'--stdin'],{input:JSON.stringify({...args,operationId:'cli_plan'}),encoding:'utf8',windowsHide:true});
  assert.equal(cli.status,0,cli.stderr);assert.equal(JSON.parse(cli.stdout).outcome,'layout_planned');
  assert.equal(plan.placements.find(p=>p.name==='LayoutChart').after.rotation,undefined);
  for (const p of plan.placements) { assert(p.after.left>=40 && p.after.top>=40); assert(p.after.left+p.after.width<=920.01); assert(p.after.top+p.after.height<=500.01); }
  await assert.rejects(host.relayout({...args,operationId:'duplicate',targetRefs:[args.targetRefs[0],args.targetRefs[0]]}),{code:'DUPLICATE_TARGET'});
  const applied = await host.relayout({...args,operationId:'apply',dryRun:false}); assert.equal(applied.revision,1);
  assert.deepEqual(await host.relayout({...args,operationId:'apply',dryRun:false}),applied);
  const after = await host.inspect({documentId:built.documentId});
  assert.deepEqual(after.objects.map(o=>o.text),snapshot.objects.map(o=>o.text));
  assert.deepEqual(after.objects.find(o=>o.kind==='chart').chart,snapshot.objects.find(o=>o.kind==='chart').chart);
  await assert.rejects(host.relayout({...args,operationId:'stale',dryRun:false}),{code:'REVISION_MISMATCH'});
  await host.finish({reviewIds:[]});
});

test('real MCP creates and edits native SmartArt, reflows charts, and preserves advanced objects through template composition', {timeout:180000}, async () => {
  assert.equal(listProcesses('POWERPNT.EXE').length,0,'Use a desktop without existing PowerPoint for native ownership verification');
  const root=rootFor(); await fs.mkdir(root,{recursive:true});
  const client=new Client({name:'advanced-workflow',version:'1.0.0'}), transport=new StdioClientTransport({command:process.execPath,args:[path.resolve('src/mcp.js'),'--base-dir',path.join(root,'tasks')],stderr:'pipe'});
  const started=performance.now(), timings={};
  const reviewIds=[];
  const call=async(name,args={})=>{const t=performance.now(),result=await client.callTool({name,arguments:args});timings[name]=(timings[name]||0)+performance.now()-t;const data=JSON.parse(result.content[0].text);assert(!result.isError,JSON.stringify(data));return data;};
  try {
    await client.connect(transport);
    const tools=await client.listTools(); assert(tools.tools.some(t=>t.name==='ppt_relayout'));
    const seedPath=path.join(root,'smartart-seed.pptx'); await fs.writeFile(seedPath,await smartArtSeed());
    await call('ppt_open',{path:seedPath,mode:'native-copy',allowOffice:true,operationId:'seed'});
    const status = await call('ppt_status'), binding = status.bindings[0];
    assert.equal(binding.taskId,status.taskId);assert.equal(binding.documentId,binding.native.documentId);
    assert.notEqual(binding.native.openedPath,binding.originalPath);
    assert(sameProcess(binding.native.worker)); assert(sameProcess(binding.native.office));
    const job = createOwnedJob(), competitor = new NativeHost(job);
    try {
      await competitor.start();
      const identity=competitor.identity, ownerKey=competitor.ownerKey;
      competitor.identity={...identity,created:'0'};
      await assert.rejects(competitor.request('inspect',{documentId:binding.documentId}),{code:'NATIVE_PROCESS_MISMATCH'});
      competitor.identity=identity;competitor.ownerKey=randomUUID();
      await assert.rejects(competitor.request('initialize'),{code:'ALREADY_STARTED'});
      await assert.rejects(competitor.request('inspect',{documentId:binding.documentId}),{code:'TASK_BINDING_MISMATCH'});
      competitor.ownerKey=ownerKey;
      await assert.rejects(competitor.request('inspect',{documentId:binding.documentId}),{code:'DOCUMENT_UNAVAILABLE'});
      await assert.rejects(competitor.request('open',{documentId:randomUUID(),path:seedPath,directory:root,visible:false}),{code:'NATIVE_BUSY'});
    } finally { const report=await competitor.shutdown();assert.equal(report.workerExited,true);job.close(); }
    assert(sameProcess(binding.native.office));
    const built=await call('ppt_build',{deck:advancedDeck(),mode:'native-copy',allowOffice:true,operationId:'build'});
    let snapshot=await call('ppt_inspect',{documentId:built.documentId,limit:300});
    assert.equal(snapshot.objects.find(o=>o.name==='Title').runs[0].style.color,'152D4A');
    assert.equal(snapshot.objects.find(o=>o.kind==='background').color,'F5F7FB');
    for (const layout of ['process','cycle','hierarchy']) {
      const target=snapshot.objects.find(o=>o.name===`SmartArt_${layout}`);
      await call('ppt_apply',{documentId:built.documentId,expectedRevision:snapshot.revision,operationId:`convert_${layout}`,operations:[{type:'convert_to_smartart',targetRef:target.targetRef,layout:'hierarchy'}]});
      snapshot=await call('ppt_inspect',{documentId:built.documentId,limit:300});
      const art=snapshot.objects.find(o=>o.slide===target.slide&&o.kind==='smartart'); assert(art); assert.equal(art.smartart.nodes.length,3);
    }
    const art=snapshot.objects.find(o=>o.kind==='smartart');
    await call('ppt_apply',{documentId:built.documentId,expectedRevision:snapshot.revision,operationId:'edit_node',operations:[{type:'set_smartart_text',targetRef:art.targetRef,nodeIndex:2,text:'设计'}]});
    snapshot=await call('ppt_inspect',{documentId:built.documentId,limit:300});
    const args={documentId:built.documentId,expectedRevision:snapshot.revision,slide:8,theme:'executive',layout:'grid',titleRef:snapshot.objects.find(o=>o.slide===8&&o.name==='Title').targetRef,
      targetRefs:snapshot.objects.filter(o=>o.slide===8&&o.capabilities.includes('set_geometry')&&o.name!=='Title').map(o=>o.targetRef)};
    const before=await call('ppt_commit',{documentId:built.documentId,expectedRevision:snapshot.revision,operationId:'before',outputPath:path.join(root,'before-layout.pptx')});
    reviewIds.push(before.reviewId);
    const beforePreview=await call('ppt_render',{documentId:built.documentId,reviewId:before.reviewId,slides:[8],width:1200,allowOffice:true});
    await writeJson(path.join(root,'before-preview.json'),beforePreview);
    const planned=await call('ppt_relayout',{...args,operationId:'plan'}); assert.equal(planned.outcome,'layout_planned');
    const changed=await call('ppt_relayout',{...args,operationId:'layout',dryRun:false});
    const committed=await call('ppt_commit',{documentId:built.documentId,expectedRevision:changed.revision,operationId:'commit',outputPath:path.join(root,'advanced.pptx')});
    reviewIds.push(committed.reviewId);
    const preview=await call('ppt_render',{documentId:built.documentId,reviewId:committed.reviewId,slides:[1,2,3,4,5,6,7,8],width:1200,allowOffice:true});
    assert.equal(preview.nativeValidation.snapshot.objects.filter(o=>o.kind==='chart').length,7);
    assert.equal(preview.nativeValidation.snapshot.objects.filter(o=>o.kind==='smartart').length,3);
    assert(preview.nativeValidation.snapshot.objects.some(o=>o.smartart?.nodes.some(n=>n.text==='设计')));
    const ended=await call('ppt_finish',{reviewIds:[committed.reviewId]}); assert.equal(ended.shutdownReport.officeExited,true);
    await writeJson(path.join(root,'preview.json'),preview);
    await call('ppt_start',{taskId:randomUUID()});
    const template=path.join(root,'template.pptx');
    await fs.writeFile(template,await buildDeck({theme:'editorial',slides:[{items:[{type:'text',name:'TemplateTitle',text:'Template / advanced content',left:40,top:40,width:850,height:70,fontSize:32,bold:true}]}]}));
    const source=await fs.readFile(committed.outputPath), sourceParts=await readPackage(source), originalIndex=indexPackage(sourceParts);
    assert.equal(originalIndex.objects.filter(o=>o.kind==='smartart').length,3);
    const composed=await call('ppt_compose',{templatePath:template,sourcePath:committed.outputPath,beautify:false,allowOffice:true,operationId:'compose'});
    const result=await call('ppt_commit',{documentId:composed.documentId,expectedRevision:0,operationId:'publish',outputPath:path.join(root,'advanced-composed.pptx')});
    reviewIds.push(result.reviewId);
    const composedParts=await readPackage(await fs.readFile(result.outputPath)), composedIndex=indexPackage(composedParts);
    for (const kind of ['chart','smartart']) {
      const oldObjects=originalIndex.objects.filter(o=>o.kind===kind), newObjects=composedIndex.objects.filter(o=>o.kind===kind);
      assert.equal(newObjects.length,oldObjects.length);
      for (let i=0;i<oldObjects.length;i++) { assert.equal(newObjects[i][kind].sha256,oldObjects[i][kind].sha256); if(kind==='chart')assert.deepEqual(newObjects[i].chart.workbooks.map(w=>w.sha256),oldObjects[i].chart.workbooks.map(w=>w.sha256)); }
    }
    const composedPreview=await call('ppt_render',{documentId:composed.documentId,reviewId:result.reviewId,slides:[2,6,7,8,9],width:1200,allowOffice:true});
    assert.equal(composedPreview.nativeValidation.snapshot.objects.filter(o=>o.kind==='smartart').length,3);
    assert.equal(composedPreview.nativeValidation.snapshot.objects.filter(o=>o.kind==='chart').length,7);
    const finish=await call('ppt_finish',{reviewIds:[result.reviewId]}); assert.equal(finish.shutdownReport.officeExited,true);
    assert.equal(hash(await fs.readFile(committed.outputPath)),hash(source));
    await writeJson(path.join(root,'composed-preview.json'),composedPreview);
    await writeJson(path.join(root,'report.json'),{root,node:process.version,totalMs:performance.now()-started,timings,charts:7,smartArt:3,layout:planned,sourceUnchanged:true,officeExited:true});
    console.log(JSON.stringify({advancedRoot:root,totalMs:performance.now()-started,timings}));
  } finally {try {await client.callTool({name:'ppt_finish',arguments:{reviewIds,preserveCheckpoints:true}});} catch {} await client.close();}
});
