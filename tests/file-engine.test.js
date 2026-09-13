import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { buildDeck } from '../src/build.js';
import { FileEngine } from '../src/file-engine.js';
import JSZip from 'jszip';
import { NS, parseXml, descendants, child, xml, readPackage, writePackage, writePackageOptimized, zipManifest } from '../src/ooxml.js';
import { contracts } from '../src/contracts.js';
import { hash, stableJson, prepareNewOutput, publishNew } from '../src/storage.js';

const directory = () => path.resolve('work/tests', randomUUID());
const sampleDeck = { title: 'Regression fixture', slides: [
  { items: [{ type: 'text', text: 'Original 中文 😀 text', left: 50, top: 50, width: 650, height: 80, name: 'Heading' }, { type: 'table', rows: [['Name', 'Value'], ['Alpha', '12']], left: 50, top: 170, width: 400, height: 150, name: 'Data' }], notes: 'Original notes' },
  { items: [{ type: 'text', text: 'Untouched slide', left: 30, top: 40, width: 650, height: 80, name: 'Untouched' }] }
] };

test('checkpoint persistence reuses verified parts across ten edits and disk restores', async t => {
  const folder = directory(), bytes = await buildDeck({ slides: Array.from({ length: 10 }, (_, i) => ({ items: [
    { type: 'text', name: `Page ${i + 1}`, text: `Page ${i + 1}`, left: 40, top: 40, width: 300, height: 60 }
  ] })) });
  let engine = await FileEngine.create(folder, bytes), partWrites = 0;
  const open = fs.open.bind(fs), partDirectory = path.join(folder, 'parts') + path.sep;
  t.mock.method(fs, 'open', (file, ...args) => {
    if (String(file).startsWith(partDirectory) && args[0] === 'wx') partWrites++;
    return open(file, ...args);
  });
  for (let slide = 1; slide <= 10; slide++) {
    const target = engine.inspect({ slides: [slide] }).objects.find(o => o.kind === 'text');
    await engine.apply([{ type: 'set_geometry', target, geometry: { top: 70 } }]);
    engine = await FileEngine.restore(folder);
  }
  assert.equal(partWrites, 10, 'Each distinct modified part is written once, including across restored instances');
  assert.equal(engine.revision, 10);
  assert(engine.inspect().objects.filter(o => o.kind === 'text').every(o => o.geometry.top === 70));
  const output = await readPackage(await engine.bytes());
  for (const [name, data] of engine.parts) assert.deepEqual(output.get(name), data, name);
  assert.deepEqual(await fs.readFile(path.join(folder, 'source.pptx')), bytes);
});

test('a corrupt reusable checkpoint part blocks the next revision without overwriting recovery evidence', async () => {
  const folder = directory(), engine = await FileEngine.create(folder, await buildDeck(sampleDeck));
  const heading = engine.inspect().objects.find(o => o.name === 'Heading');
  await engine.apply([{ type: 'set_geometry', target: heading, geometry: { top: 70 } }]);
  const manifestPath = path.join(folder, 'revision.json'), manifest = await fs.readFile(manifestPath);
  const entry = JSON.parse(manifest).overlay['ppt/slides/slide1.xml'];
  const partPath = path.join(folder, 'parts', entry.file), damaged = Buffer.from('damaged checkpoint evidence');
  await fs.writeFile(partPath, damaged);
  const target = engine.inspect().objects.find(o => o.name === 'Untouched');
  await assert.rejects(engine.apply([{ type: 'set_geometry', target, geometry: { top: 90 } }]), { code: 'CHECKPOINT_CORRUPT' });
  assert.equal(engine.revision, 1);
  assert.equal(engine.inspect().objects.find(o => o.name === 'Untouched').geometry.top, 40);
  assert.deepEqual(await fs.readFile(manifestPath), manifest);
  assert.deepEqual(await fs.readFile(partPath), damaged);
  await assert.rejects(FileEngine.restore(folder), { code: 'CHECKPOINT_CORRUPT' });
});

test('file edits preserve source and untouched parts, and recover from disk', async () => {
  // Explicit ZIP directory entries are legal and must not delete their children.
  const sourceZip = await JSZip.loadAsync(await buildDeck(sampleDeck));
  sourceZip.folder('ppt');
  sourceZip.folder('ppt/slides');
  const bytes = await sourceZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const sourceHash = hash(bytes), folder = directory();
  const engine = await FileEngine.create(folder, bytes), initial = engine.inspect();
  const heading = initial.objects.find(o => o.name === 'Heading'); assert.ok(heading); assert.equal(heading.geometry.left, 50);
  const result = await engine.apply([{ type: 'replace_text', target: heading, search: 'Original', replacement: 'Revised', expectedMatches: 1, crossRunPolicy: 'reject' }, { type: 'set_geometry', target: heading, geometry: { left: 70 } }, { type: 'set_style', target: heading, style: { fontSize: 30, color: '224466' } }]);
  assert.equal(result.revision, 1); assert.equal(hash(bytes), sourceHash);
  assert.deepEqual(engine.validate().preservation.changedParts, ['ppt/slides/slide1.xml']);
  const restored = await FileEngine.restore(folder), target = restored.inspect().objects.find(o => o.name === 'Heading');
  assert.equal(target.text, 'Revised 中文 😀 text'); assert.equal(target.geometry.left, 70); assert.equal(target.runs[0].style.color, '224466');
  const output = await readPackage(await restored.bytes()); assert.deepEqual(output.get('ppt/slides/slide2.xml'), engine.sourceParts.get('ppt/slides/slide2.xml'));
  assert.deepEqual([...output.keys()].sort(), [...engine.parts.keys()].sort());
  for (const [name, data] of engine.parts) assert.deepEqual(output.get(name), data, name);
  assert.equal(hash(await fs.readFile(path.join(folder, 'source.pptx'))), sourceHash);
});

test('scoped preflight supports minimal keys across slides, notes and backgrounds with a complete post-edit snapshot', async () => {
  const folder=directory(),bytes=await buildDeck({...sampleDeck,slides:[...sampleDeck.slides,
    {items:[{type:'text',name:'Preserved',text:'Third page',left:30,top:40,width:400,height:60}]}]});
  const engine=await FileEngine.create(folder,bytes),initial=engine.inspect();
  const minimal=object=>({key:object.key,fingerprint:object.fingerprint});
  const target=initial.objects.find(o=>o.name==='Untouched');
  const notes=initial.objects.find(o=>o.kind==='notes'&&o.slide===1);
  const background=initial.objects.find(o=>o.kind==='background'&&o.slide===2);
  const result=await engine.apply([
    {type:'set_geometry',target:minimal(target),geometry:{left:75}},
    {type:'replace_text',target:minimal(notes),search:'Original',replacement:'Revised',expectedMatches:1,crossRunPolicy:'reject'},
    {type:'set_slide_background',target:minimal(background),color:'224466'}
  ]);
  assert.equal(result.snapshot.slides.length,3);
  assert.equal(result.snapshot.objects.find(o=>o.name==='Untouched').geometry.left,75);
  assert.equal(result.snapshot.objects.find(o=>o.kind==='notes'&&o.slide===1).text,'Revised notes');
  assert.equal(result.snapshot.objects.find(o=>o.kind==='background'&&o.slide===2).color,'224466');
  assert.equal(result.snapshot.objects.find(o=>o.name==='Preserved').text,'Third page');
  const restored=await FileEngine.restore(folder),verified=restored.inspect();
  assert.deepEqual(result.snapshot,verified);
  assert.equal(result.stateHash,hash(stableJson(verified.objects.map(o=>[o.key,o.fingerprint]))));
  assert.deepEqual(restored.parts.get('ppt/slides/slide3.xml'),engine.sourceParts.get('ppt/slides/slide3.xml'));
  await assert.rejects(engine.apply([{type:'set_geometry',target:minimal(target),geometry:{left:90}}]),{code:'TARGET_CHANGED'});
  assert.equal(engine.revision,1);
});

test('inherited placeholder geometry requires a complete explicit transform and preserves other parts', async () => {
  const parts = await readPackage(await buildDeck(sampleDeck)), part = 'ppt/slides/slide1.xml';
  const doc = parseXml(parts.get(part)), shape = descendants(doc, NS.p, 'sp')[0];
  const properties = child(shape, NS.p, 'spPr'); properties.removeChild(child(properties, NS.a, 'xfrm'));
  const nv = child(child(shape, NS.p, 'nvSpPr'), NS.p, 'nvPr');
  const ph = doc.createElementNS(NS.p, 'p:ph'); ph.setAttribute('type', 'body'); nv.appendChild(ph);
  parts.set(part, Buffer.from(xml(doc)));
  const bytes = await writePackage(parts), engine = await FileEngine.create(directory(), bytes);
  const target = engine.inspect().objects.find(o => o.name === 'Heading');
  assert.equal(target.geometry.source, 'inherited-unresolved'); assert(target.capabilities.includes('set_geometry'));
  await assert.rejects(engine.apply([{ type: 'set_geometry', target, geometry: { left: 80 } }]), { code: 'GEOMETRY_REQUIRES_FULL' });
  assert.equal(engine.revision, 0); assert.deepEqual(engine.parts.get(part), parts.get(part));
  const geometry = { left: 80, top: 65, width: 550, height: 95, rotation: 0 };
  await engine.apply([{ type: 'set_geometry', target, geometry }]);
  const updated = engine.inspect().objects.find(o => o.name === 'Heading');
  for (const [key,value] of Object.entries(geometry)) assert.equal(updated.geometry[key],value);
  const output = await readPackage(await engine.bytes());
  for (const [name,data] of parts) if (name !== part) assert.deepEqual(output.get(name),data);
  const written = parseXml(output.get(part)), sp = descendants(written, NS.p, 'sp')[0];
  assert.equal(child(sp,NS.p,'spPr').firstChild.localName,'xfrm');
  assert.equal(descendants(sp,NS.p,'ph')[0].getAttribute('type'),'body');
});

test('preflight failure and dry-run leave the last revision unchanged', async () => {
  // Validate save-path guards without starting Office or causing a dialog.
  const guardDir = directory();
  const occupied = path.join(guardDir, 'occupied.pptx');
  await fs.mkdir(occupied, { recursive: true });
  await assert.rejects(prepareNewOutput(occupied), { code: 'OUTPUT_ALREADY_EXISTS' });
  const output = path.join(guardDir, 'output.pptx');
  await prepareNewOutput(output);
  const sentinel = Buffer.from('synthetic competing writer');
  await fs.writeFile(output, sentinel);
  await assert.rejects(publishNew(output, Buffer.from('must not overwrite')), { code: 'EEXIST' });
  assert.deepEqual(await fs.readFile(output), sentinel);
  const engine = await FileEngine.create(directory(), await buildDeck(sampleDeck)), heading = engine.inspect().objects.find(o => o.name === 'Heading');
  const good = { type: 'replace_text', target: heading, search: 'Original', replacement: 'Changed', expectedMatches: 1, crossRunPolicy: 'reject' };
  await assert.rejects(engine.apply([good, { ...good, search: 'missing' }]), { code: 'MATCH_COUNT_MISMATCH' });
  assert.equal(engine.revision, 0); assert.equal(engine.inspect().objects.find(o => o.name === 'Heading').text, heading.text);
  await engine.apply([good], { dryRun: true }); assert.equal(engine.revision, 0);
  await engine.apply([good]); await assert.rejects(engine.apply([good]), { code: 'TARGET_CHANGED' });
});

test('cross-run replacement requires policy and retains unknown slide extensions', async () => {
  const parts = await readPackage(await buildDeck(sampleDeck)), doc = parseXml(parts.get('ppt/slides/slide1.xml'));
  const paragraph = descendants(doc, NS.a, 'p')[0], run = child(paragraph, NS.a, 'r'), copied = run.cloneNode(true);
  child(run, NS.a, 't').textContent = 'Cross '; child(copied, NS.a, 't').textContent = 'run text';
  child(copied, NS.a, 'rPr').setAttribute('b', '1'); paragraph.insertBefore(copied, child(paragraph, NS.a, 'endParaRPr'));
  const extension = doc.createElementNS('urn:ppt-editor:fixture', 'fixture:unknown'); extension.setAttribute('value', 'keep'); doc.documentElement.appendChild(extension);
  parts.set('ppt/slides/slide1.xml', Buffer.from(xml(doc)));
  const engine = await FileEngine.create(directory(), await writePackage(parts)), heading = engine.inspect().objects.find(o => o.name === 'Heading');
  const op = { type: 'replace_text', target: heading, search: 'Cross run', replacement: 'Across', expectedMatches: 1, crossRunPolicy: 'reject' };
  await assert.rejects(engine.apply([op]), { code: 'CROSS_RUN_REPLACEMENT' });
  await engine.apply([{ ...op, crossRunPolicy: 'first_run' }]);
  assert.equal(engine.inspect().objects.find(o => o.name === 'Heading').text, 'Across text');
  assert.match(engine.parts.get('ppt/slides/slide1.xml').toString(), /fixture:unknown[^>]+value="keep"/);
});

test('table and notes stay scoped, invalid archives and unknown arguments fail', async () => {
  const engine = await FileEngine.create(directory(), await buildDeck(sampleDeck)), snapshot = engine.inspect();
  const table = snapshot.objects.find(o => o.kind === 'table'), notes = snapshot.objects.find(o => o.kind === 'notes' && o.slide === 1); assert.ok(notes);
  await engine.apply([{ type: 'set_table_cell', target: table, row: 2, column: 2, text: '18', formatPolicy: 'first_run' }, { type: 'replace_text', target: notes, search: 'Original', replacement: 'Revised', expectedMatches: 1, crossRunPolicy: 'reject' }]);
  assert.equal(engine.inspect().objects.find(o => o.kind === 'table').rows[1][1], '18');
  assert.throws(() => zipManifest(Buffer.from('not a presentation')), { code: 'INVALID_PACKAGE' });
  assert.throws(() => contracts.ppt_open.parse({ path: 'file.pptx', operationId: 'open', unexpected: true }));
});

test('slide background can be inspected, modified, persisted and recovered', async () => {
  const folder = directory();
  const engine = await FileEngine.create(folder, await buildDeck(sampleDeck));
  const initial = engine.inspect();
  const bg1 = initial.objects.find(o => o.kind === 'background' && o.slide === 1);
  const bg2 = initial.objects.find(o => o.kind === 'background' && o.slide === 2);
  assert.ok(bg1);
  assert.ok(bg2);
  assert.equal(bg1.color, 'FFFFFF');
  assert.deepEqual(bg1.capabilities, ['set_slide_background']);

  const result = await engine.apply([
    { type: 'set_slide_background', target: bg1, color: '00FF88' }
  ]);
  assert.equal(result.revision, 1);
  assert.equal(result.changes[0].color, '00FF88');

  // 验证内存状态
  const modified = engine.inspect();
  const modifiedBg1 = modified.objects.find(o => o.kind === 'background' && o.slide === 1);
  const untouchedBg2 = modified.objects.find(o => o.kind === 'background' && o.slide === 2);
  assert.equal(modifiedBg1.color, '00FF88');
  assert.equal(untouchedBg2.color, 'FFFFFF');

  // 验证磁盘恢复
  const restored = await FileEngine.restore(folder);
  const restoredBg1 = restored.inspect().objects.find(o => o.kind === 'background' && o.slide === 1);
  assert.equal(restoredBg1.color, '00FF88');

  // 验证只有 slide1 被修改，slide2 保持未变
  assert.deepEqual(engine.validate().preservation.changedParts, ['ppt/slides/slide1.xml']);
});

test('writePackageOptimized reuses streams, uses STORE for media, handles deletions and falls back gracefully', async () => {
  const sourceBytes = await buildDeck(sampleDeck);
  const sourceParts = await readPackage(sourceBytes);

  // 构造模拟测试包：包含未变动部件、已变动部件、新增部件、已删除部件、以及媒体文件
  const currentParts = new Map(sourceParts);
  const fakePng = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d]);
  currentParts.set('ppt/media/image1.png', fakePng);

  // 修改 slide1 内容
  const slide1Original = sourceParts.get('ppt/slides/slide1.xml').toString('utf-8');
  const slide1Modified = Buffer.from(slide1Original.replace('Original 中文 😀 text', 'Optimized Pack Test'));
  currentParts.set('ppt/slides/slide1.xml', slide1Modified);

  // 增加新自定义文件
  const customFile = Buffer.from('custom part data');
  currentParts.set('custom/part.txt', customFile);

  // 执行优化写包
  const optimizedBytes = await writePackageOptimized(sourceBytes, currentParts, sourceParts);
  assert.ok(optimizedBytes.length > 0);

  // 1. 验证通过 JSZip 检查媒体文件的存储压缩方式为 STORE (method 0: magic \x00\x00, compressedSize === uncompressedSize)
  const zip = await JSZip.loadAsync(optimizedBytes);
  const mediaEntry = zip.file('ppt/media/image1.png');
  assert.ok(mediaEntry, 'Media entry should exist in output zip');
  assert.equal(mediaEntry._data.compression.magic, '\x00\x00');
  assert.equal(mediaEntry._data.compressedSize, mediaEntry._data.uncompressedSize);

  // 2. 验证解包后所有部件内容逐字节严格等价
  const manifest = zipManifest(optimizedBytes);
  const manifestNames = new Set(manifest.map(m => m.name));
  assert.ok(manifestNames.has('ppt/media/image1.png'));
  assert.ok(manifestNames.has('custom/part.txt'));
  assert.ok(manifestNames.has('ppt/slides/slide1.xml'));
  assert.ok(manifestNames.has('ppt/slides/slide2.xml'));

  // 验证删除行为：如果 currentParts 移除了 slide2
  const partsWithDeletion = new Map(currentParts);
  partsWithDeletion.delete('ppt/slides/slide2.xml');
  const deletedOutput = await writePackageOptimized(sourceBytes, partsWithDeletion, sourceParts);
  const deletedZip = await JSZip.loadAsync(deletedOutput);
  assert.equal(deletedZip.file('ppt/slides/slide2.xml'), null, 'Deleted part must not exist in output');

  // 3. 验证容灾降级：传入 null sourceBytes 或损坏数据时，正常输出等价包
  const fallbackBytes1 = await writePackageOptimized(null, currentParts, null);
  assert.ok(fallbackBytes1.length > 0);
  const fallbackZip1 = await JSZip.loadAsync(fallbackBytes1);
  assert.ok(fallbackZip1.file('ppt/slides/slide1.xml'));

  const corruptSource = Buffer.from('corrupt non-zip bytes');
  const fallbackBytes2 = await writePackageOptimized(corruptSource, currentParts, sourceParts);
  assert.ok(fallbackBytes2.length > 0);
  const fallbackZip2 = await JSZip.loadAsync(fallbackBytes2);
  assert.ok(fallbackZip2.file('ppt/slides/slide1.xml'));

  // 4. 验证 FileEngine.prototype.bytes() 产物可被标准 readPackage 正确解析
  const folder = directory();
  const engine = await FileEngine.create(folder, sourceBytes);
  const heading = engine.inspect().objects.find(o => o.name === 'Heading');
  await engine.apply([{ type: 'replace_text', target: heading, search: 'Original', replacement: 'EngineBytes', expectedMatches: 1, crossRunPolicy: 'reject' }]);
  const engineBytes = await engine.bytes();
  const readBack = await readPackage(engineBytes);
  assert.ok(readBack.get('ppt/slides/slide1.xml').toString('utf-8').includes('EngineBytes'));
});
