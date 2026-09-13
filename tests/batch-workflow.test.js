import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { Client, InMemoryTransport } from '@modelcontextprotocol/client';
import { TaskHost } from '../src/task.js';
import { createPptMcpServer } from '../src/mcp.js';
import { readJson, writeJson, hash, stableJson } from '../src/storage.js';
import { contracts } from '../src/contracts.js';
import { layoutPageHashes } from '../src/file-engine.js';
import { listProcesses } from '../src/windows.js';
import { buildDeck } from '../src/build.js';
import { startStdioPeer } from './helpers/stdio-peer.js';

const text = (name, left) => ({ type: 'text', name, text: name, left, top: 30, width: 160, height: 60 });
const deck = { width: 400, height: 300, slides: [{ items: [text('A',20),text('B',100)] },
  { items: [text('C',20)] }, { items: [text('D',300)] }] };
const makeHost = () => new TaskHost({ baseDir: path.resolve('work/tests/batch-workflow',randomUUID(),'tasks') });

test('page measurement coverage persists independently and aggregates only current known evidence', async () => {
  const host = makeHost();
  host.ensureNativeHost = async () => { throw Error('Coverage regression must not start Office'); };
  try {
    const { documentId } = await host.build({ operationId: 'build', deck: { width: 640, height: 360, slides: [
      { items: [text('Measured', 40)] }, { items: [text('Unmeasured', 100)] }
    ] } });
    let session = await host.getSession(documentId);
    const native = session.engine.inspect();
    for (const item of native.objects) if (item.geometry) {
      item.geometry.source = 'native-effective';
      if (item.name === 'Measured') item.textBounds = { left: 45, top: 35, width: 100, height: 30 };
    }
    const audit = await host.recordLayoutAudit(session, native, {});
    const before = await readJson(audit.worklistPath);
    assert.equal(before.pages[0].coverage.textFit, 'native-text-bounds-partial');
    assert.deepEqual(before.pages[1].coverage.textFitDetails, { measured: 0, unmeasured: 1, reasons: { native_measurement_unavailable: 1 } });
    assert.deepEqual(before.coverage.textFitDetails, { measured: 1, unmeasured: 1, reasons: { native_measurement_unavailable: 1 } });
    const inspected = await host.inspect({ documentId, slide: 1 });
    const changed = await host.apply({ documentId, expectedRevision: 0, operationId: 'edit', operations: [
      { type: 'set_geometry', targetRef: inspected.objects.find(o => o.name === 'Measured').targetRef, geometry: { top: 70 } }
    ] });
    const after = await readJson(changed.layoutAudit.worklistPath);
    assert.deepEqual(after.pages[1], before.pages[1]);
    assert.deepEqual(after.coverage.textFitDetails, { measured: 0, unmeasured: 2, reasons: { native_readback_not_run: 1, native_measurement_unavailable: 1 } });
    const restored = await new TaskHost({ taskId: host.taskId, baseDir: host.baseDir }).status();
    assert.deepEqual(restored.documents.find(d => d.documentId === documentId).layout.pages[1].coverage, before.pages[1].coverage);

    // Older worklists do not contain per-page counts; do not invent zeroes.
    delete after.pages[1].coverage.textFitDetails;
    await writeJson(changed.layoutAudit.worklistPath, after);
    session = await host.getSession(documentId);
    await host.recordLayoutAudit(session, session.engine.inspect({ slides: [1] }), { slides: [1] });
    const legacy = await readJson(changed.layoutAudit.worklistPath);
    assert.equal(legacy.coverage.textFitDetails.unknownPageCount, 1);
    assert.deepEqual(legacy.pages[1].coverage, after.pages[1].coverage);

    session.generation++;
    await host.recordLayoutAudit(session, session.engine.inspect({ slides: [1] }), { slides: [1] });
    const stale = await readJson(changed.layoutAudit.worklistPath);
    assert.equal(stale.pages[1].stale, true);
    assert.equal(stale.coverage.textFitDetails.unmeasured, 1);
    assert.equal(stale.coverage.textFitDetails.unknownPageCount, 1);
    await host.finish({ reviewIds: [] });
  } finally { await host.cleanupCommand(); }
});

test('batch edits return page reports, preserve unchanged design decisions, and expose recoverable task summaries', async () => {
  const host=makeHost(),instance=createPptMcpServer({taskHost:host}),client=new Client({name:'batch',version:'1'});
  host.ensureNativeHost=async()=>{throw Error('Batch drafting must not start Office');};
  const [ct,st]=InMemoryTransport.createLinkedPair();await instance.server.connect(st);await client.connect(ct);
  const call=async(name,args)=>{const r=await client.callTool({name,arguments:args});assert(!r.isError,r.content[0].text);return r.structuredContent;};
  try {
    const built=await call('ppt_build',{deck,operationId:'build'}),id=built.documentId;
    assert.deepEqual(built.layoutAudit.summary.attentionSlides,[1,3]);
    assert.deepEqual(built.layoutAudit.pendingSlides.map(p=>p.slide),[1,3]);
    assert.equal(built.layoutAudit.workflow.pendingSlides,undefined,'Summary worklists have one canonical pending-page list');
    const parts=(await host.getSession(id)).engine.parts, dependencies=new Map(parts);
    const theme=[...parts.keys()].find(p=>p.startsWith('ppt/theme/')&&p.endsWith('.xml'));
    dependencies.set(theme,Buffer.from(parts.get(theme).toString().replace('</a:theme>','<!-- changed theme -->\n</a:theme>')));
    const before=layoutPageHashes(parts),after=layoutPageHashes(dependencies);
    for(const slide of [1,2,3]) assert.notEqual(before.get(slide),after.get(slide),'Inherited theme changes invalidate dependent page evidence');
    const pair=built.layoutAudit.pages[0].issues.find(i=>i.code==='POSSIBLE_OVERLAP').objects.map(o=>o.key);
    await call('ppt_validate',{documentId:id,layoutCheck:true,checks:'layout',slides:[1],expectedRevision:0,allowedOverlapPairs:[pair]});
    const full=await call('ppt_inspect',{documentId:id}),brief=await call('ppt_inspect',{documentId:id,detail:'summary'});
    assert(Buffer.byteLength(JSON.stringify(brief))<Buffer.byteLength(JSON.stringify(full))*0.85);
    const changed=await call('ppt_apply',{documentId:id,expectedRevision:0,operationId:'batch-fix',operations:['C','D'].map(name=>({type:'set_geometry',targetRef:full.objects.find(o=>o.name===name).targetRef,geometry:{left:200}}))});
    assert.deepEqual(changed.layoutAudit.summary.clearSlides,[2,3]);
    assert.deepEqual(changed.layoutAudit.pendingSlides,[]);
    assert.equal(changed.layoutAudit.workflow.finalWholeDeckCheck,'required');
    assert.deepEqual(await call('ppt_apply',{documentId:id,expectedRevision:0,operationId:'batch-fix',operations:['C','D'].map(name=>({type:'set_geometry',targetRef:full.objects.find(o=>o.name===name).targetRef,geometry:{left:200}}))}),changed);
    const persisted=await new TaskHost({taskId:host.taskId,baseDir:host.baseDir}).status();
    assert.equal(persisted.documents.find(d=>d.documentId===id).layout.pages[0].stale,false);
    const final=await call('ppt_validate',{documentId:id,layoutCheck:true});
    assert.deepEqual(final.layoutAudit.pendingSlides,[]);
    assert.deepEqual(final.layoutAudit.workflow.pendingSlides,[],'Full responses keep the existing shape');
    const reset=await call('ppt_validate',{documentId:id,layoutCheck:true,slides:[1],allowedOverlapPairs:[]});
    assert.equal(reset.layoutAudit.screenshotRequiredNow,true);
    const receiptPath=path.join(host.taskDir,'receipts','build.json'),legacy=await readJson(receiptPath);
    legacy.result.layoutAudit.workflow.pendingSlides=structuredClone(legacy.result.layoutAudit.pendingSlides);
    await writeJson(receiptPath,legacy);
    assert.deepEqual(await call('ppt_build',{deck,operationId:'build'}),legacy.result,'Existing receipts replay their original response shape');
    const {layoutCheck,...oldParams}=contracts.ppt_build.parse({deck,operationId:'build'});
    delete legacy.receiptVersion;delete legacy.result.layoutAudit;legacy.paramsHash=hash(stableJson(oldParams));
    await writeJson(receiptPath,legacy);
    const replay=await call('ppt_build',{deck,operationId:'build'});
    assert.equal(replay.documentId,id);assert.equal(replay.layoutAudit,undefined,'Legacy receipts replay without regenerating a deck');
    await call('ppt_finish',{reviewIds:[]});
  } finally {await instance.cleanup();await client.close();await instance.server.close();}
});

test('file edits retain unchanged native page evidence and invalidate changed dependencies or generations', async () => {
  const host=makeHost();
  host.ensureNativeHost=async()=>{throw Error('This state-transition fixture must not start Office');};
  const nativeSnapshot=session=>{
    const snapshot=session.engine.inspect();
    for(const object of snapshot.objects) if(object.geometry){
      object.geometry.source='native-effective';
      if(object.kind==='text') object.textBounds={left:object.geometry.left,top:object.geometry.top,width:100,height:30};
    }
    return snapshot;
  };
  try {
    const {documentId}=await host.build({deck,operationId:'build'});
    let session=await host.getSession(documentId);
    const snapshot=nativeSnapshot(session),pair=snapshot.objects.filter(o=>['A','B'].includes(o.name)).map(o=>o.key);
    const native=await host.recordLayoutAudit(session,snapshot,{allowedOverlapPairs:[pair]});
    const before=await readJson(native.worklistPath);
    const inspected=await host.inspect({documentId,slide:2,detail:'summary'});
    const changed=await host.apply({documentId,expectedRevision:0,operationId:'file-edit',operations:[
      {type:'set_geometry',targetRef:inspected.objects.find(o=>o.name==='C').targetRef,geometry:{left:80}}
    ]});
    assert.deepEqual(changed.layoutAudit.pendingSlides.map(p=>p.slide),[3],'Only the existing issue remains pending');
    assert.equal(changed.layoutAudit.workflow.finalWholeDeckCheck,'required');
    const after=await readJson(changed.layoutAudit.worklistPath);
    for(const slide of [1,3]) assert.deepEqual(after.pages.find(p=>p.slide===slide),before.pages.find(p=>p.slide===slide));
    const edited=after.pages.find(p=>p.slide===2);
    assert.equal(edited.backend,'file');assert.equal(edited.coverage.textFit,'not_checked');assert.equal(edited.checkedRevision,1);
    assert.equal(after.pages[2].status,'issues','Preserving evidence must not clear existing native issues');
    const restored=await new TaskHost({taskId:host.taskId,baseDir:host.baseDir}).status();
    assert.deepEqual(restored.documents.find(d=>d.documentId===documentId).layout.pendingSlides.map(p=>p.slide),[3]);

    // Simulate an inherited dependency update through the audit boundary.
    session=await host.getSession(documentId);
    const theme=[...session.engine.parts.keys()].find(p=>p.startsWith('ppt/theme/')&&p.endsWith('.xml'));
    session.engine.parts.set(theme,Buffer.from(session.engine.parts.get(theme).toString().replace('</a:theme>','<!-- changed dependency --></a:theme>')));
    const dependency=await host.recordLayoutAudit(session,session.engine.inspect({slides:[2]}),{slides:[2],detail:'summary'});
    assert.deepEqual(dependency.pendingSlides.filter(p=>p.stale).map(p=>p.slide),[1,3]);
    await host.recordLayoutAudit(session,nativeSnapshot(session),{allowedOverlapPairs:[pair]});
    session.generation++;
    const generation=await host.recordLayoutAudit(session,session.engine.inspect({slides:[2]}),{slides:[2],detail:'summary'});
    assert.deepEqual(generation.pendingSlides.filter(p=>p.stale).map(p=>p.slide),[1,3]);
    await host.finish({reviewIds:[]});
  } finally {await host.cleanupCommand();}
});

test('finishing without render evidence records incomplete acceptance and checks even unrendered output bytes', async () => {
  const host=makeHost();
  try {
    const {documentId}=await host.build({deck,operationId:'build'});
    const outputPath=path.join(host.baseDir,'unreviewed.pptx');
    const committed=await host.commit({documentId,expectedRevision:0,operationId:'commit',outputPath});
    await assert.rejects(host.finish({reviewIds:[committed.reviewId],requireAccepted:true}),{code:'ACCEPTANCE_INCOMPLETE'});
    assert.equal((await host.status()).status,'active');
    const bytes=await fs.readFile(outputPath);await fs.writeFile(outputPath,'changed');
    await assert.rejects(host.finish({reviewIds:[committed.reviewId]}),{code:'OUTPUT_CHANGED'});
    await fs.writeFile(outputPath,bytes);
    const ended=await host.finish({reviewIds:[committed.reviewId]});
    assert.equal(ended.status,'finished');assert.equal(ended.acceptance.status,'incomplete');
    const review=(await host.status({reviewId:committed.reviewId})).reviewBundle;
    assert.equal(review.callerConfirmed,true);assert.equal(review.visualReviewed,false);
    assert.equal(review.acceptance.status,'incomplete');
  } finally {await host.cleanupCommand();}
});

test('combined render checks the whole readback without treating unrendered pages as visually reviewed', async () => {
  const host=makeHost();
  try {
    const {documentId}=await host.build({deck:{width:400,height:300,slides:[{items:[text('First',40)]},{items:[text('Second',40)]}]},allowOffice:true,operationId:'build'});
    const snapshot=(await host.getSession(documentId)).engine.inspect();
    for(const object of snapshot.objects) if(object.geometry) object.geometry.source='native-effective';
    let reads=0;
    host.ensureNativeHost=async()=>({request:async(method,args)=>{
      assert.equal(method,'validateFile'); reads++;
      const images=[];
      for(const slide of args.render.slides){
        const imagePath=path.join(args.render.directory,`slide-${slide}.png`);
        const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jWZkAAAAASUVORK5CYII=','base64');
        await fs.writeFile(imagePath,png); images.push({slide,path:imagePath,sha256:hash(png)});
      }
      return {nativeReadback:'passed',readOnly:true,sha256:args.sha256,slidesCount:2,snapshot,images};
    }});
    const committed=await host.commit({documentId,expectedRevision:0,outputPath:path.join(host.baseDir,'merged.pptx'),operationId:'commit'});
    const args={documentId,reviewId:committed.reviewId,allowOffice:true,layoutCheck:true,detail:'summary'};
    const first=await host.render({...args,slides:[1]});
    assert.equal(reads,1); assert.deepEqual(first.layoutAudit.summary.clearSlides,[1,2]);
    assert.equal(first.layoutAudit.workflow.finalWholeDeckCheck,'current');
    assert.equal(first.nativeValidation.snapshot,undefined);
    const recorded=await readJson(first.reviewBundlePath);
    assert.equal(recorded.preview.nativeValidation.snapshot.slides.length,2);
    await assert.rejects(host.finish({reviewIds:[committed.reviewId],requireAccepted:true}),{code:'ACCEPTANCE_INCOMPLETE'});
    await host.render({...args,slides:[2]}); assert.equal(reads,2);
    assert.equal((await host.finish({reviewIds:[committed.reviewId],requireAccepted:true})).acceptance.status,'accepted');
  } finally {await host.cleanupCommand();}
});

test('file drafts preview exact revisions and committed previews accumulate page evidence before acceptance', {timeout:90000}, async t => {
  if(listProcesses('POWERPNT.EXE').length){t.skip('Existing PowerPoint session');return;}
  const host=makeHost();
  try {
    const {documentId}=await host.build({deck:{width:400,height:300,slides:[{items:[text('First page',40)]},{items:[text('Second page',40)]}]},allowOffice:true,operationId:'build'});
    const args={documentId,slides:[1],allowOffice:true,width:800,detail:'summary'};
    await assert.rejects(host.render({...args,expectedRevision:1}),{code:'REVISION_MISMATCH'});
    const draft=await host.render({...args,expectedRevision:0});
    assert.equal(draft.basis,'current-checkpoint-bytes');assert.equal(draft.revision,0);assert.equal(draft.images.length,1);
    assert.equal((await fs.readdir(path.join(host.taskDir,'reviews'))).length,0);
    const committed=await host.commit({documentId,expectedRevision:0,outputPath:path.join(host.baseDir,'batch-preview.pptx'),operationId:'commit'});
    const first=await host.render({...args,reviewId:committed.reviewId});
    assert.equal(first.nativeValidation.snapshot,undefined);
    await assert.rejects(host.finish({reviewIds:[committed.reviewId],requireAccepted:true}),{code:'ACCEPTANCE_INCOMPLETE'});
    const second=await host.render({...args,reviewId:committed.reviewId,slides:[2],layoutCheck:true});
    assert.equal(second.images.length,1);
    assert.equal(second.layoutAudit.workflow.finalWholeDeckCheck,'current');
    assert.deepEqual(second.layoutAudit.summary.clearSlides,[1,2]);
    assert.equal(second.reviewBundlePath,path.join(host.taskDir,'reviews',`${committed.reviewId}.json`));
    const recorded=await readJson(path.join(host.taskDir,'reviews',`${committed.reviewId}.json`));
    assert.deepEqual(recorded.preview.images.map(i=>i.slide),[1,2]);
    assert(recorded.preview.nativeValidation.snapshot.objects.some(o=>o.textBounds?.height>0));
    const ended=await host.finish({reviewIds:[committed.reviewId],requireAccepted:true});
    assert.equal(ended.acceptance.status,'accepted');assert.equal(ended.shutdownReport.officeExited,true);
    console.log('BATCH_VISUAL_ARTIFACTS '+JSON.stringify(recorded.preview.images.map(i=>i.path)));
  } finally {await host.cleanupCommand();}
});

test('STDIO file refinement preserves real native evidence and accepts the newly rendered output', {timeout:90000}, async t => {
  if(listProcesses('POWERPNT.EXE').length){t.skip('Existing PowerPoint session');return;}
  const root=path.resolve('work/tests/batch-native-evidence',randomUUID());
  await fs.mkdir(root,{recursive:true});
  const source=path.join(root,'source.pptx'),bytes=await buildDeck({width:640,height:360,slides:[
    {items:[{...text('First page',40),width:500}]},{items:[{...text('Second page',40),width:500}]}
  ]});
  await fs.writeFile(source,bytes);
  const peer=await startStdioPeer(path.join(root,'tasks'));
  try {
    const runtime=await peer.call('ppt_diagnose');
    assert.equal(runtime.version,(await readJson(path.resolve('package.json'))).version);
    const {documentId}=await peer.call('ppt_open',{path:source,mode:'file',allowOffice:true,operationId:'open'});
    const first=await peer.call('ppt_commit',{documentId,expectedRevision:0,outputPath:path.join(root,'initial.pptx'),operationId:'initial'});
    const renderArgs={documentId,slides:[1,2],allowOffice:true,layoutCheck:true,detail:'summary',width:800};
    const native=await peer.call('ppt_render',{...renderArgs,reviewId:first.reviewId},'initial-render');
    const initialProgress=peer.progress.filter(p=>p.progressToken==='initial-render');
    assert(initialProgress.some(p=>p.message.includes('reading native object snapshot')));
    assert(initialProgress.some(p=>p.message.includes('rendering slides (2/2 slides)')));
    assert(initialProgress.at(-1).message.includes('completed'));
    assert(initialProgress.every((p,i)=>!i||p.progress>initialProgress[i-1].progress));
    const before=await readJson(native.layoutAudit.worklistPath);
    assert(before.pages.every(p=>p.backend==='native'&&p.coverage.textFit==='native-text-bounds-partial'));
    for(const page of before.pages) assert.deepEqual(page.coverage.textFitDetails,{measured:1,unmeasured:0,reasons:{}});
    assert.deepEqual(before.coverage.textFitDetails,{measured:2,unmeasured:0,reasons:{}});
    const inspected=await peer.call('ppt_inspect',{documentId,slide:1,detail:'summary'});
    const changed=await peer.call('ppt_apply',{documentId,expectedRevision:0,operationId:'refine',operations:[
      {type:'set_geometry',targetRef:inspected.objects.find(o=>o.name==='First page').targetRef,geometry:{top:70}}
    ]});
    assert.deepEqual(changed.layoutAudit.pendingSlides,[]);
    assert.equal(changed.layoutAudit.workflow.pendingSlides,undefined);
    assert.equal(changed.layoutAudit.workflow.finalWholeDeckCheck,'required');
    const after=await readJson(changed.layoutAudit.worklistPath);
    assert.deepEqual(after.pages[1],before.pages[1]);
    assert.equal(after.pages[0].backend,'file');assert.equal(after.pages[0].coverage.textFit,'not_checked');
    assert.deepEqual(after.coverage.textFitDetails,{measured:1,unmeasured:1,reasons:{native_readback_not_run:1}});
    const final=await peer.call('ppt_commit',{documentId,expectedRevision:1,outputPath:path.join(root,'final.pptx'),operationId:'final'});
    const preview=await peer.call('ppt_render',{...renderArgs,reviewId:final.reviewId},'final-render');
    const finalProgress=peer.progress.filter(p=>p.progressToken==='final-render');
    assert(finalProgress.some(p=>p.message.includes('rendering slides (2/2 slides)')));
    assert(finalProgress.at(-1).message.includes('completed'));
    assert(peer.progress.every(p=>['initial-render','final-render'].includes(p.progressToken)));
    assert.deepEqual(preview.layoutAudit.summary.clearSlides,[1,2]);
    assert.deepEqual(preview.layoutAudit.coverage.textFitDetails,{measured:2,unmeasured:0,reasons:{}});
    const reviewed=await readJson(preview.reviewBundlePath);
    const objects=reviewed.preview.nativeValidation.snapshot.objects;
    assert.equal(objects.find(o=>o.name==='First page').geometry.top,70);
    assert.equal(objects.find(o=>o.name==='Second page').geometry.top,30);
    assert.deepEqual(await fs.readFile(source),bytes);
    const ended=await peer.call('ppt_finish',{reviewIds:[final.reviewId],requireAccepted:true});
    assert.equal(ended.acceptance.status,'accepted');assert.equal(ended.shutdownReport.officeExited,true);
    console.log('OPTIMIZATION_VISUAL_ARTIFACTS '+JSON.stringify({version:runtime.version,images:preview.images.map(i=>i.path)}));
  } finally {const exit=await peer.stop();assert.equal(exit.code,0,exit.stderr);}
});
