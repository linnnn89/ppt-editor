import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { Client, InMemoryTransport } from '@modelcontextprotocol/client';
import { auditLayout } from '../src/layout-audit.js';
import { TaskHost } from '../src/task.js';
import { createPptMcpServer } from '../src/mcp.js';
import { buildDeck } from '../src/build.js';
import { listProcesses } from '../src/windows.js';

const object = (key,left,top,width,height,extra={}) => ({ key,name:key,kind:'text',slide:1,groupPath:[],geometry:{left,top,width,height,rotation:0},...extra });

test('text-fit coverage counts actual measurements per page and explains unsupported text', () => {
  const native = { left: 40, top: 40, width: 200, height: 60, rotation: 0, source: 'native-effective' };
  const objects = [
    { textBounds: { left: 45, top: 45, width: 160, height: 25 } },
    { geometry: { ...native, rotation: 20 } },
    { groupPath: [9] },
    {},
    { geometry: { ...native, source: 'slide-explicit' } },
    { textBounds: { left: 45, top: 45, width: -1, height: 25 } }
  ].map((extra, i) => object(`text-${i + 1}`, 40, 40, 200, 60, { slide: i + 1, text: 'Visible text', geometry: native, ...extra }));
  const snapshot = { width: 640, height: 360, slides: objects.map(o => ({ slide: o.slide })), objects };
  const result = auditLayout(snapshot);
  assert.equal(result.pages[0].coverage.textFit, 'native-text-bounds-partial');
  assert.deepEqual(result.pages[0].coverage.textFitDetails, { measured: 1, unmeasured: 0, reasons: {} });
  const reasons = ['rotated_text', 'group_child', 'native_measurement_unavailable', 'native_readback_not_run', 'invalid_text_bounds'];
  for (let i = 1; i < result.pages.length; i++) {
    assert.equal(result.pages[i].coverage.textFit, 'not_checked');
    assert.deepEqual(result.pages[i].coverage.textFitDetails, { measured: 0, unmeasured: 1, reasons: { [reasons[i - 1]]: 1 } });
  }
  assert.deepEqual(result.coverage.textFitDetails, { measured: 1, unmeasured: 5, reasons: Object.fromEntries(reasons.map(r => [r, 1])) });
  assert.equal(auditLayout(snapshot, { slides: [2] }).coverage.textFit, 'not_checked');
  assert.equal(result.pages[1].status, 'clear', 'Unmeasured text does not change the established geometry acceptance rule');
});
test('layout audit distinguishes declared design overlaps, suspicious overlaps, page bounds and unknown geometry', () => {
  const snapshot={width:100,height:100,slides:[{slide:1}],objects:[object('a',10,10,30,20),object('b',30,15,30,20),object('edge',90,90,10,10),object('bg',0,0,100,100,{kind:'background'})]};
  const result=auditLayout(snapshot);
  assert.deepEqual(result.pages[0].issues[0].intersectionPt,{width:10,height:15});
  assert.equal(result.screenshotRequiredNow,true);
  assert.equal(result.nextAction,'review_overlap_before_continuing');
  const designed=auditLayout(snapshot,{allowedOverlapPairs:[['a','b']]});
  assert.equal(designed.pages[0].status,'clear');assert.equal(designed.pages[0].allowedOverlapCount,1);
  assert.equal(designed.screenshotRequiredNow,false);
  assert.throws(()=>auditLayout(snapshot,{allowedOverlapPairs:[['a','missing']]}),{code:'INVALID_OVERLAP_EXCEPTION'});
  const bounds=auditLayout({...snapshot,objects:[object('out',95,20,10,20),object('unknown',null,0,10,10),object('child',0,0,10,10,{groupPath:[2]})]});
  assert.equal(bounds.pages[0].counts.outOfBounds,1);assert.equal(bounds.pages[0].counts.unresolved,2);
  assert.equal(bounds.pages[0].issues[0].overflowPt.right,5);
  const rotated=auditLayout({...snapshot,objects:[object('rotated',80,40,20,40,{geometry:{left:80,top:40,width:20,height:40,rotation:90}})]});
  assert.equal(rotated.pages[0].issues[0].overflowPt.right,10);
  assert.equal(result.coverage.textFit,'not_checked');
  const textBounds=auditLayout({...snapshot,objects:[object('text',10,10,40,20,{textBounds:{left:12,top:12,width:30,height:30}})]});
  assert.equal(textBounds.pages[0].counts.textOverflow,1);
  assert.equal(textBounds.pages[0].issues[0].overflowPt.bottom,12);
});

const text=(name,left,top=20)=>({type:'text',name,text:name,left,top,width:160,height:40});
const deck={width:400,height:300,slides:[{items:[text('A',20),text('B',100)]},{items:[text('C',20)]},{items:[text('D',300)]}]};
test('MCP per-page checks persist a revision-bound worklist, skip rendering, and final audit reflects repairs', async () => {
  const root=path.resolve('work/tests/layout-audit',randomUUID());await fs.mkdir(root,{recursive:true});
  const bytes=await buildDeck(deck), source=path.join(root,'source.pptx');await fs.writeFile(source,bytes);
  const host=new TaskHost({baseDir:path.join(root,'tasks')});
  host.ensureNativeHost=async()=>{throw Error('Page checks must not launch Office');};
  const instance=createPptMcpServer({taskHost:host}),client=new Client({name:'layout-audit',version:'1.0.0'});
  const [ct,st]=InMemoryTransport.createLinkedPair();await instance.server.connect(st);await client.connect(ct);
  const call=async(name,args)=>{const result=await client.callTool({name,arguments:args});assert(!result.isError,result.content[0].text);return JSON.parse(result.content[0].text);};
  try {
    const {documentId}=await call('ppt_open',{path:source,operationId:'open'});
    const first=(await call('ppt_validate',{documentId,layoutCheck:true,slides:[1]})).layoutAudit;
    assert.equal(first.screenshotRequiredNow,true);assert.deepEqual(first.summary.attentionSlides,[1]);
    const pair=first.pages[0].issues[0].objects.map(o=>o.key);
    const designed=(await call('ppt_validate',{documentId,layoutCheck:true,slides:[1],expectedRevision:0,allowedOverlapPairs:[pair]})).layoutAudit;
    assert.equal(designed.pages[0].status,'clear');
    const second=(await call('ppt_validate',{documentId,layoutCheck:true,slides:[2]})).layoutAudit;
    assert.equal(second.screenshotRequiredNow,false);assert.equal(second.pages[0].status,'clear');
    const saved=JSON.parse(await fs.readFile(second.worklistPath,'utf8'));
    assert.deepEqual(saved.pages.find(p=>p.slide===1).allowedOverlapPairs,[pair]);
    const cli=spawnSync(process.execPath,[path.resolve('src/cli.js'),'validate','--task',host.taskId,'--base-dir',host.baseDir,'--doc',documentId,'--layout-check','--slides','[2]'],{encoding:'utf8',windowsHide:true});
    assert.equal(cli.status,0,cli.stderr);assert.equal(JSON.parse(cli.stdout).layoutAudit.screenshotRequiredNow,false);
    const initial=(await call('ppt_validate',{documentId,layoutCheck:true,allowedOverlapPairs:[]})).layoutAudit;
    assert.deepEqual(initial.summary.attentionSlides,[1,3]); // explicit [] clears design declarations
    const snapshot=await call('ppt_inspect',{documentId});
    await call('ppt_apply',{documentId,expectedRevision:0,operationId:'repair',operations:[
      {type:'set_geometry',targetRef:snapshot.objects.find(o=>o.name==='B').targetRef,geometry:{left:220}},
      {type:'set_geometry',targetRef:snapshot.objects.find(o=>o.name==='D').targetRef,geometry:{left:220}}
    ]});
    await assert.rejects(host.validate({documentId,layoutCheck:true,expectedRevision:0,allowedOverlapPairs:[pair]}),{code:'REVISION_MISMATCH'});
    const repaired=(await call('ppt_validate',{documentId,layoutCheck:true,slides:[1]})).layoutAudit;
    assert.equal(repaired.revision,1);assert(!repaired.pendingSlides.some(p=>p.slide===2)); // unchanged page remains checked
    const final=(await call('ppt_validate',{documentId,layoutCheck:true})).layoutAudit;
    assert.deepEqual(final.summary.attentionSlides,[]);assert.deepEqual(final.pendingSlides,[]);
    const worklist=JSON.parse(await fs.readFile(final.worklistPath,'utf8'));
    assert.equal(worklist.pages.length,3);assert(worklist.pages.every(p=>!p.stale&&p.checkedRevision===1));
    assert.equal((await call('ppt_inspect',{documentId})).revision,1);
    assert.deepEqual(await fs.readFile(source),bytes);
    assert.deepEqual(await fs.readdir(path.join(host.taskDir,'reviews')),[]);
    await call('ppt_finish',{reviewIds:[]});assert((await fs.stat(final.worklistPath)).isFile());
  } finally {await instance.cleanup();await client.close();await instance.server.close();}
});

test('native page audit measures the bound presentation without exporting an image and exits cleanly', {timeout:60000}, async t => {
  if(listProcesses('POWERPNT.EXE').length){t.skip('Existing PowerPoint session');return;}
  const root=path.resolve('work/tests/layout-audit',randomUUID()),host=new TaskHost({baseDir:root});
  try {
    const {documentId}=await host.build({deck,mode:'native-copy',allowOffice:true,operationId:'build'});
    const report=await host.validate({documentId,layoutCheck:true,slides:[1,2]});
    assert.equal(report.nativeReadback,'passed');assert.equal(report.layoutAudit.screenshotRequiredNow,true);
    assert.deepEqual(report.layoutAudit.summary.clearSlides,[2]);assert.deepEqual(report.layoutAudit.summary.attentionSlides,[1]);
    assert.equal(report.layoutAudit.revision,0);assert.equal(report.layoutAudit.generation,1);
    assert.deepEqual(await fs.readdir(path.join(host.taskDir,'reviews')),[]);
    const file=await host.open({path:(await host.getSession(documentId)).originalPath,mode:'file',allowOffice:true,operationId:'file'});
    const readback=await host.validate({documentId:file.documentId,layoutCheck:true,slides:[1,2],nativeReadback:true,allowOffice:true});
    assert.equal(readback.nativeReadback,'passed');assert.deepEqual(readback.layoutAudit.summary.attentionSlides,[1]);
    assert.deepEqual(readback.layoutAudit.summary.clearSlides,[2]);
    const ended=await host.finish({reviewIds:[]});assert.equal(ended.shutdownReport.workerExited,true);assert.equal(ended.shutdownReport.officeExited,true);
  } finally {await host.cleanupCommand();}
});
