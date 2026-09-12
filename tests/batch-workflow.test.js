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

const text = (name, left) => ({ type: 'text', name, text: name, left, top: 30, width: 160, height: 60 });
const deck = { width: 400, height: 300, slides: [{ items: [text('A',20),text('B',100)] },
  { items: [text('C',20)] }, { items: [text('D',300)] }] };
const makeHost = () => new TaskHost({ baseDir: path.resolve('work/tests/batch-workflow',randomUUID(),'tasks') });

test('batch edits return page reports, preserve unchanged design decisions, and expose recoverable task summaries', async () => {
  const host=makeHost(),instance=createPptMcpServer({taskHost:host}),client=new Client({name:'batch',version:'1'});
  host.ensureNativeHost=async()=>{throw Error('Batch drafting must not start Office');};
  const [ct,st]=InMemoryTransport.createLinkedPair();await instance.server.connect(st);await client.connect(ct);
  const call=async(name,args)=>{const r=await client.callTool({name,arguments:args});assert(!r.isError,r.content[0].text);return r.structuredContent;};
  try {
    const built=await call('ppt_build',{deck,operationId:'build'}),id=built.documentId;
    assert.deepEqual(built.layoutAudit.summary.attentionSlides,[1,3]);
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
    const reset=await call('ppt_validate',{documentId:id,layoutCheck:true,slides:[1],allowedOverlapPairs:[]});
    assert.equal(reset.layoutAudit.screenshotRequiredNow,true);
    const receiptPath=path.join(host.taskDir,'receipts','build.json'),legacy=await readJson(receiptPath);
    const {layoutCheck,...oldParams}=contracts.ppt_build.parse({deck,operationId:'build'});
    delete legacy.receiptVersion;delete legacy.result.layoutAudit;legacy.paramsHash=hash(stableJson(oldParams));
    await writeJson(receiptPath,legacy);
    const replay=await call('ppt_build',{deck,operationId:'build'});
    assert.equal(replay.documentId,id);assert.equal(replay.layoutAudit,undefined,'Legacy receipts replay without regenerating a deck');
    await call('ppt_finish',{reviewIds:[]});
  } finally {await instance.cleanup();await client.close();await instance.server.close();}
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
