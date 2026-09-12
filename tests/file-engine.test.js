import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { buildDeck } from '../src/build.js';
import { FileEngine } from '../src/file-engine.js';
import { NS, parseXml, descendants, child, xml, readPackage, writePackage, zipManifest } from '../src/ooxml.js';
import { contracts } from '../src/contracts.js';
import { hash } from '../src/storage.js';

const directory = () => path.resolve('work/tests', randomUUID());
const sampleDeck = { title: 'Regression fixture', slides: [
  { items: [{ type: 'text', text: 'Original 中文 😀 text', left: 50, top: 50, width: 650, height: 80, name: 'Heading' }, { type: 'table', rows: [['Name', 'Value'], ['Alpha', '12']], left: 50, top: 170, width: 400, height: 150, name: 'Data' }], notes: 'Original notes' },
  { items: [{ type: 'text', text: 'Untouched slide', left: 30, top: 40, width: 650, height: 80, name: 'Untouched' }] }
] };

test('file edits preserve source and untouched parts, and recover from disk', async () => {
  const bytes = await buildDeck(sampleDeck), sourceHash = hash(bytes), folder = directory();
  const engine = await FileEngine.create(folder, bytes), initial = engine.inspect();
  const heading = initial.objects.find(o => o.name === 'Heading'); assert.ok(heading); assert.equal(heading.geometry.left, 50);
  const result = await engine.apply([{ type: 'replace_text', target: heading, search: 'Original', replacement: 'Revised', expectedMatches: 1, crossRunPolicy: 'reject' }, { type: 'set_geometry', target: heading, geometry: { left: 70 } }, { type: 'set_style', target: heading, style: { fontSize: 30, color: '224466' } }]);
  assert.equal(result.revision, 1); assert.equal(hash(bytes), sourceHash);
  assert.deepEqual(engine.validate().preservation.changedParts, ['ppt/slides/slide1.xml']);
  const restored = await FileEngine.restore(folder), target = restored.inspect().objects.find(o => o.name === 'Heading');
  assert.equal(target.text, 'Revised 中文 😀 text'); assert.equal(target.geometry.left, 70); assert.equal(target.runs[0].style.color, '224466');
  const output = await readPackage(await restored.bytes()); assert.deepEqual(output.get('ppt/slides/slide2.xml'), engine.sourceParts.get('ppt/slides/slide2.xml'));
  assert.equal(hash(await fs.readFile(path.join(folder, 'source.pptx'))), sourceHash);
});

test('preflight failure and dry-run leave the last revision unchanged', async () => {
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
