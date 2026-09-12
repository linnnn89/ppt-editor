import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { performance } from 'node:perf_hooks';
import { deflateSync, crc32 } from 'node:zlib';
import { Client } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';
import { buildDeck } from '../src/build.js';
import { TaskHost } from '../src/task.js';
import { readPackage, parseXml, NS, descendants } from '../src/ooxml.js';
import { readJson, writeJson, hash } from '../src/storage.js';
import { listProcesses } from '../src/windows.js';
import { bicycle, bicycleDeck } from '../scripts/fixtures/bicycle.js';
import { indexPackage } from '../src/file-engine.js';

const rootFor = () => path.resolve('work/tests/build-review', randomUUID());
const line = { type: 'shape', shape: 'line', left: 10, top: 20, width: 80, height: 60, name: 'Line' };
async function imageFixture(root) {
  await fs.mkdir(root, { recursive: true });
  const chunk = (type, data) => { const name = Buffer.from(type), length = Buffer.alloc(4), crc = Buffer.alloc(4); length.writeUInt32BE(data.length); crc.writeUInt32BE(crc32(Buffer.concat([name, data]))); return Buffer.concat([length, name, data, crc]); };
  const header = Buffer.alloc(13); header.writeUInt32BE(16); header.writeUInt32BE(16, 4); header[8] = 8; header[9] = 2;
  const rows = Buffer.alloc(16 * (1 + 16 * 3));
  for (let y = 0; y < 16; y++) for (let x = 0; x < 16; x++) { const at = y * 49 + 1 + x * 3; rows[at] = x < 8 ? 220 : 245; rows[at + 1] = y < 8 ? 230 : 250; rows[at + 2] = 255; }
  const bytes = Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]), chunk('IHDR', header), chunk('IDAT', deflateSync(rows)), chunk('IEND', Buffer.alloc(0))]);
  const image = path.join(root, 'pattern.png'); await fs.writeFile(image, bytes); return image;
}
const backgrounds = image => [
  { type: 'solid', color: 'DDEEFF', transparency: 10 },
  { type: 'gradient', startColor: 'DDEEFF', endColor: 'FFFFFF', angle: 35, transparency: 0 },
  { type: 'pattern', pattern: 'diagonalCross', foreground: 'CDDDEE', background: 'FFFFFF', transparency: 0 },
  { type: 'image', path: image, transparency: 0 }, { type: 'texture', path: image, transparency: 0 }, { type: 'inherit' }
];

test('formal drawing preserves old geometry and supports editable flipped and rotated shapes', async () => {
  const items = [line, { ...line, name: 'Rising', flipV: true }, { ...line, name: 'Mirrored', flipH: true }, { ...line, name: 'Rotated', rotation: 30 }, { ...line, name: 'Vertical', width: 0 }, { ...line, name: 'Horizontal', height: 0 }];
  const parts = await readPackage(await buildDeck({ slides: [{ items }] }));
  const doc = parseXml(parts.get('ppt/slides/slide1.xml'));
  const props = descendants(doc, NS.p, 'cNvPr');
  const transform = name => descendants(props.find(n => n.getAttribute('name') === name).parentNode.parentNode, NS.a, 'xfrm')[0];
  assert.equal(transform('Line').hasAttribute('flipV'), false);
  assert.equal(transform('Rising').getAttribute('flipV'), '1');
  assert.equal(transform('Mirrored').getAttribute('flipH'), '1');
  assert.equal(transform('Rotated').getAttribute('rot'), '1800000');
  assert.equal(descendants(transform('Vertical'), NS.a, 'ext')[0].getAttribute('cx'), '0');
  assert.equal(descendants(transform('Horizontal'), NS.a, 'ext')[0].getAttribute('cy'), '0');
  assert.equal(descendants(doc, NS.p, 'pic').length, 0);
  await assert.rejects(buildDeck({ slides: [{ items: [{ ...line, flipV: 'true' }] }] }));
  const bicycleParts = await readPackage(await buildDeck(bicycleDeck()));
  assert.equal(descendants(parseXml(bicycleParts.get('ppt/slides/slide1.xml')), NS.p, 'sp').length, 50);
  const root = rootFor(), image = await imageFixture(root), fills = backgrounds(image);
  const rich = await buildDeck({ slides: fills.map(background => ({ background, items: [{ type: 'text', text: 'Styled 中文', left: 20, top: 20, width: 200, height: 60, italic: true, bold: true, underline: true }] })) });
  const index = indexPackage(await readPackage(rich));
  assert.deepEqual(index.objects.filter(o => o.kind === 'background').map(o => o.fill.type), fills.map(f => f.type));
  assert(index.objects.filter(o => o.text).every(o => o.runs.every(r => r.style.underline && r.style.bold && r.style.italic)));
});

test('committed preview rejects missing consent, changed outputs and cross-instance stale reviews before Office', async () => {
  const root = rootFor(), host = new TaskHost({ baseDir: path.join(root, 'tasks') });
  const { documentId } = await host.build({ deck: { slides: [{ items: [line] }] }, operationId: 'build' });
  const committed = await host.commit({ documentId, expectedRevision: 0, outputPath: path.join(root, 'out.pptx'), operationId: 'commit' });
  const args = { documentId, reviewId: committed.reviewId, slides: [1], allowOffice: true };
  await assert.rejects(host.render(args), { code: 'OFFICE_NOT_ALLOWED' });
  assert.equal(host.nativeHost, null);
  const docDir = path.join(host.taskDir, 'documents', documentId), metaFile = path.join(docDir, 'meta.json');
  await writeJson(metaFile, { ...await readJson(metaFile), allowOffice: true });
  await assert.rejects(host.render({ ...args, reviewId: randomUUID() }), { code: 'REVIEW_NOT_FOUND' });
  const original = await fs.readFile(committed.outputPath);
  await fs.writeFile(committed.outputPath, 'replaced');
  await assert.rejects(host.render(args), { code: 'OUTPUT_CHANGED' });
  assert.equal(host.nativeHost, null);
  await fs.writeFile(committed.outputPath, original);
  // Controlled replacement after native readback starts, before evidence is accepted.
  host.ensureNativeHost = async () => ({ request: async (_method, request) => {
    await fs.writeFile(committed.outputPath, 'changed during render');
    return { nativeReadback: 'passed', readOnly: true, sha256: request.sha256, images: [] };
  } });
  await assert.rejects(host.render(args), { code: 'OUTPUT_CHANGED' });
  assert.equal((await readJson(path.join(host.taskDir, 'reviews', `${committed.reviewId}.json`))).preview, undefined);
  await fs.writeFile(committed.outputPath, original);
  const other = new TaskHost({ taskId: host.taskId, baseDir: host.baseDir });
  const inspected = await other.inspect({ documentId });
  await other.apply({ documentId, expectedRevision: 0, operationId: 'move', operations: [{ type: 'set_geometry', targetRef: inspected.objects.find(o => o.name === 'Line').targetRef, geometry: { left: 30 } }] });
  await assert.rejects(host.render(args), { code: 'REVIEW_STALE' });
  const status = await host.status({ reviewId: committed.reviewId });
  assert.equal(status.reviewBundle.stale, true);
  await host.finish({ reviewIds: [] });
});

test('real STDIO A workflow builds bicycle, renders committed bytes and retains reviewed evidence after finish', { timeout: 120000 }, async () => {
  assert.equal(listProcesses('POWERPNT.EXE').length, 0, 'Use a desktop without existing PowerPoint for ownership verification');
  const root = rootFor(); await fs.mkdir(root, { recursive: true });
  const client = new Client({ name: 'formal-bicycle-regression', version: '1.0.0' });
  const transport = new StdioClientTransport({ command: process.execPath, args: [path.resolve('src/mcp.js'), '--base-dir', path.join(root, 'tasks')], stderr: 'pipe' });
  const stages = {}, start = performance.now();
  const call = async (name, args = {}) => {
    const t = performance.now(), response = await client.callTool({ name, arguments: args });
    stages[name] = (stages[name] || 0) + performance.now() - t;
    const data = JSON.parse(response.content[0].text); assert(!response.isError, JSON.stringify(data)); return data;
  };
  try {
    await client.connect(transport);
    const diagnosis = await call('ppt_diagnose');
    assert.equal(diagnosis.capabilities.committedOutputPreview, true);
    const { documentId } = await call('ppt_build', { deck: bicycleDeck(), mode: 'file', allowOffice: true, operationId: 'build' });
    const inspected = await call('ppt_inspect', { documentId });
    assert.equal(inspected.objects.filter(o => bicycle().some(i => i.name === o.name)).length, 50);
    const committed = await call('ppt_commit', { documentId, expectedRevision: 0, operationId: 'commit', outputPath: path.join(root, 'bicycle.pptx') });
    const rendered = await call('ppt_render', { documentId, reviewId: committed.reviewId, slides: [1], width: 1600, allowOffice: true });
    assert.equal(rendered.outputHash, committed.outputHash);
    assert.equal(rendered.nativeValidation.readOnly, true);
    assert.equal(rendered.nativeValidation.sha256, committed.outputHash);
    for (const item of bicycle()) {
      const actual = rendered.nativeValidation.snapshot.objects.find(o => o.name === item.name); assert(actual, item.name);
      for (const key of ['left', 'top', 'width', 'height']) assert(Math.abs(actual.geometry[key] - item[key]) < 0.1, `${item.name}:${key}`);
    }
    const image = rendered.images[0]; assert.equal(hash(await fs.readFile(image.path)), image.sha256);
    assert.equal(hash(await fs.readFile(committed.outputPath)), committed.outputHash);
    // Output changes must block acknowledgment even after a successful render.
    const original = await fs.readFile(committed.outputPath);
    await fs.writeFile(committed.outputPath, 'replaced after preview');
    const refused = await client.callTool({ name: 'ppt_finish', arguments: { reviewIds: [committed.reviewId] } });
    assert.equal(JSON.parse(refused.content[0].text).code, 'OUTPUT_CHANGED');
    await fs.writeFile(committed.outputPath, original);
    const finished = await call('ppt_finish', { reviewIds: [committed.reviewId] });
    assert.equal(finished.shutdownReport.workerExited, true); assert.equal(finished.shutdownReport.officeExited, true);
    const status = await call('ppt_status', { reviewId: committed.reviewId });
    assert.equal(status.status, 'finished'); assert.equal(status.reviewBundle.visualReviewed, true);
    assert.equal(status.reviewBundle.visualReviewMethod, 'caller-confirmation');
    assert.equal(hash(await fs.readFile(status.reviewBundle.preview.images[0].path)), image.sha256);
    const summary = { root, stages, totalMs: performance.now() - start, diagnosis, rendered, finished };
    await writeJson(path.join(root, 'report.json'), summary);
    console.log(JSON.stringify({ formalBicycle: root, stages, totalMs: summary.totalMs }));
    // The same actual connection exercises both backends and all background types.
    const imagePath = await imageFixture(root), fills = backgrounds(imagePath);
    for (const mode of ['file', 'native-copy']) {
      await call('ppt_start', { taskId: randomUUID() });
      const built = await call('ppt_build', { deck: { slides: fills.map((_, n) => ({ items: [{ type: 'text', name: `Heading${n}`, text: 'Before 中文', left: 40, top: 40, width: 400, height: 90 }] })) }, mode, allowOffice: true, operationId: 'styles-build' });
      const id = built.documentId, before = await call('ppt_inspect', { documentId: id });
      const heading = before.objects.find(o => o.name === 'Heading0');
      const ops = fills.map((fill, i) => ({ type: 'set_slide_background', targetRef: before.objects.find(o => o.kind === 'background' && o.slide === i + 1).targetRef, fill }));
      ops.push({ type: 'set_style', targetRef: heading.targetRef, style: { fontFace: 'Microsoft YaHei', fontSize: 28, color: '17365D', bold: true, italic: true, underline: true } });
      ops.push({ type: 'set_geometry', targetRef: heading.targetRef, geometry: { left: 70, top: 80, width: 500, height: 100, rotation: 5 } });
      await call('ppt_apply', { documentId: id, expectedRevision: 0, operationId: 'styles', operations: ops });
      let current = await call('ppt_inspect', { documentId: id });
      assert.equal(current.objects.find(o => o.name === 'Heading0').runs[0].style.underline, true);
      await call('ppt_apply', { documentId: id, expectedRevision: 1, operationId: 'underline-off', operations: [{ type: 'set_style', targetRef: current.objects.find(o => o.name === 'Heading0').targetRef, style: { underline: false } }] });
      current = await call('ppt_inspect', { documentId: id });
      assert.equal(current.objects.find(o => o.name === 'Heading0').runs[0].style.underline, false);
      const commit = await call('ppt_commit', { documentId: id, expectedRevision: 2, operationId: 'styles-commit', outputPath: path.join(root, `styles-${mode}.pptx`) });
      const preview = await call('ppt_render', { documentId: id, reviewId: commit.reviewId, slides: [1,2,3,4,5,6], width: 960, allowOffice: true });
      const actual = preview.nativeValidation.snapshot.objects.find(o => o.name === 'Heading0');
      for (const run of actual.runs) { assert.equal(run.style.underline, false); assert.equal(run.style.bold, true); assert.equal(run.style.italic, true); assert.equal(run.style.fontSize, 28); assert.equal(run.style.color, '17365D'); assert(['Microsoft YaHei', '微软雅黑'].includes(run.style.fontFaceFarEast)); }
      for (const [key, value] of Object.entries({ left: 70, top: 80, width: 500, height: 100, rotation: 5 })) assert(Math.abs(actual.geometry[key] - value) < 0.1, key);
      assert.deepEqual(preview.nativeValidation.snapshot.objects.filter(o => o.kind === 'background').map(o => o.fill.type), fills.map(f => f.type));
      const saved = indexPackage(await readPackage(await fs.readFile(commit.outputPath)));
      assert.deepEqual(saved.objects.filter(o => o.kind === 'background').map(o => o.fill.type), fills.map(f => f.type));
      await writeJson(path.join(root, `styles-${mode}-preview.json`), preview);
      const ended = await call('ppt_finish', { reviewIds: [commit.reviewId] });
      assert.equal(ended.shutdownReport.officeExited, true);
    }
    // Template content is preserved and appended content remains editable.
    await call('ppt_start', { taskId: randomUUID() });
    const templatePath = path.join(root, 'template.pptx'), sourcePath = path.join(root, 'old.pptx');
    const templateBytes = await buildDeck({ slides: [{ background: 'EDF4FA', items: [{ type: 'text', name: 'TemplateTitle', text: 'Template retained', left: 30, top: 30, width: 500, height: 70 }, { type: 'shape', shape: 'rect', name: 'Accent', left: 0, top: 0, width: 15, height: 540, fill: '17365D' }] }] });
    const oldBytes = await buildDeck({ slides: [{ background: 'FFFFFF', items: [{ type: 'text', text: 'Existing content 中文', name: 'ImportedText', left: 60, top: 90, width: 700, height: 100 }, { type: 'table', name: 'Data', left: 60, top: 240, width: 400, height: 100, rows: [['A','B'],['1','2']] }], notes: 'Notes retained' }] });
    await fs.writeFile(templatePath, templateBytes); await fs.writeFile(sourcePath, oldBytes);
    const composed = await call('ppt_compose', { templatePath, sourcePath, templateSlide: 1, allowOffice: true, operationId: 'compose' });
    assert.equal(composed.composition.templateSlidesPreserved, 1);
    const combined = await call('ppt_inspect', { documentId: composed.documentId });
    assert.equal(combined.slides.length, 2);
    assert(combined.objects.some(o => o.slide === 1 && o.text === 'Template retained'));
    assert(combined.objects.some(o => o.slide === 2 && o.text === 'Existing content 中文'));
    assert(combined.objects.some(o => o.slide === 2 && o.kind === 'table'));
    const composedCommit = await call('ppt_commit', { documentId: composed.documentId, expectedRevision: 0, outputPath: path.join(root, 'composed.pptx'), operationId: 'compose-commit' });
    const composedPreview = await call('ppt_render', { documentId: composed.documentId, reviewId: composedCommit.reviewId, slides: [1,2], width: 960, allowOffice: true });
    assert.equal(composedPreview.nativeValidation.slidesCount, 2);
    assert(composedPreview.nativeValidation.snapshot.objects.some(o => o.kind === 'notes' && o.text.includes('Notes retained')));
    await writeJson(path.join(root, 'composed-preview.json'), composedPreview);
    await call('ppt_finish', { reviewIds: [composedCommit.reviewId] });
    assert.equal(hash(await fs.readFile(templatePath)), hash(templateBytes)); assert.equal(hash(await fs.readFile(sourcePath)), hash(oldBytes));
    await call('ppt_start', { taskId: randomUUID() });
    const appended = await call('ppt_compose', { templatePath: composedCommit.outputPath, templateSlide: 1, contentDeck: { slides: [{ items: [{ type: 'text', name: 'NewContent', text: 'New page 新内容', left: 60, top: 90, width: 700, height: 100 }] }] }, allowOffice: true, operationId: 'append' });
    assert.equal(appended.composition.templateSlidesPreserved, 2);
    const appendCommit = await call('ppt_commit', { documentId: appended.documentId, expectedRevision: 0, outputPath: path.join(root, 'appended.pptx'), operationId: 'append-commit' });
    const appendPreview = await call('ppt_render', { documentId: appended.documentId, reviewId: appendCommit.reviewId, slides: [3], width: 960, allowOffice: true });
    assert.equal(appendPreview.nativeValidation.slidesCount, 3);
    assert(appendPreview.nativeValidation.snapshot.objects.some(o => o.slide === 3 && o.text === 'New page 新内容'));
    const appendedParts = await readPackage(await fs.readFile(appendCommit.outputPath)), combinedParts = await readPackage(await fs.readFile(composedCommit.outputPath));
    for (const [part, bytes] of combinedParts) if (/^ppt\/slides\/[^/]+\.xml$/.test(part)) assert.deepEqual(appendedParts.get(part), bytes, 'Existing pages stay byte-identical');
    assert.equal(indexPackage(appendedParts).objects.find(o => o.kind === 'background' && o.slide === 3).color, 'EDF4FA');
    await writeJson(path.join(root, 'appended-preview.json'), appendPreview);
    const appendFinished = await call('ppt_finish', { reviewIds: [appendCommit.reviewId] });
    assert.equal(appendFinished.shutdownReport.officeExited, true);
    await writeJson(path.join(root, 'verification.json'), { root, version: diagnosis.version, node: process.version, stages, totalMs: performance.now() - start, workflows: ['formal-bicycle', 'file-backgrounds-and-style', 'native-backgrounds-and-style', 'template-and-old-deck', 'append-new-page'], sourceHashesUnchanged: true, officeExited: true });
  } finally { await client.close(); }
});
