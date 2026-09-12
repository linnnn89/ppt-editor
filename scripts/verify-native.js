import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { buildDeck } from '../src/build.js';
import { readPackage } from '../src/ooxml.js';
import { indexPackage } from '../src/file-engine.js';
import { NativeHost } from '../src/native-host.js';
import { createOwnedJob, observeForeground, restrictAccess } from '../src/windows.js';
import { writeJson, hash } from '../src/storage.js';

const root = path.resolve('work/native-verification', randomUUID());
await fs.mkdir(root, { recursive: true }); restrictAccess(root, { directory: true });
const fixture = await buildDeck({ slides: [{ items: [{ type: 'text', text: 'Original 中文 😀 text', left: 60, top: 60, width: 700, height: 80, name: 'Heading' }, { type: 'table', rows: [['A', 'B'], ['C', '12']], left: 60, top: 180, width: 400, height: 150, name: 'Data' }], notes: 'Original notes' }] });
const source = path.join(root, 'source.pptx'), working = path.join(root, 'working.pptx');
await fs.writeFile(source, fixture); await fs.writeFile(working, fixture);
const job = createOwnedJob(), host = new NativeHost(job), observer = observeForeground(), report = { root, node: process.version, started: new Date().toISOString() };
function verifyEdits(snapshot) {
  const heading = snapshot.objects.find(o => o.name === 'Heading');
  assert.equal(heading.text, 'Revised 中文 😀 text');
  assert.equal(heading.geometry.left, 80);
  assert.ok(heading.runs.length > 0);
  for (const run of heading.runs) {
    assert.equal(run.style.fontSize, 28);
    assert.equal(run.style.color, '224466');
  }
  assert.equal(snapshot.objects.find(o => o.kind === 'table').rows[1][1], '18');
  assert.equal(snapshot.objects.find(o => o.kind === 'notes').text, 'Revised notes');
  const bg = snapshot.objects.find(o => o.kind === 'background');
  assert.ok(bg);
  assert.equal(bg.color, '123456');
}
try {
  await host.start();
  const documentId = randomUUID();
  const opened = await host.request('open', { documentId, path: working, directory: root, visible: false });
  report.lease = opened.lease; report.openedObjects = opened.snapshot.objects.length;
  for (const identity of opened.lease.observedNewProcesses) assert.equal(job.contains(identity.pid), false, 'PowerPoint must not be in the owned Node job');
  const heading = opened.snapshot.objects.find(o => o.name === 'Heading'), table = opened.snapshot.objects.find(o => o.kind === 'table'), notes = opened.snapshot.objects.find(o => o.kind === 'notes'), bg = opened.snapshot.objects.find(o => o.kind === 'background');
  assert.ok(heading && table && notes && bg);
  const applied = await host.request('apply', { documentId, generation: opened.generation, operations: [
    { type: 'replace_text', target: heading, search: 'Original', replacement: 'Revised', expectedMatches: 1, crossRunPolicy: 'reject' },
    { type: 'set_style', target: heading, style: { fontSize: 28, color: '224466' } },
    { type: 'set_geometry', target: heading, geometry: { left: 80 } },
    { type: 'set_table_cell', target: table, row: 2, column: 2, text: '18', formatPolicy: 'first_run' },
    { type: 'replace_text', target: notes, search: 'Original', replacement: 'Revised', expectedMatches: 1, crossRunPolicy: 'reject' },
    { type: 'set_slide_background', target: bg, color: '123456' }
  ] });
  report.apply = applied; assert.equal(applied.outcome, 'completed');
  verifyEdits(applied.snapshot);
  const checkpointIndex = indexPackage(await readPackage(await fs.readFile(applied.checkpoint.path)));
  verifyEdits(checkpointIndex);
  report.render = await host.request('render', { documentId, slides: [1], width: 1280, directory: path.join(root, 'preview') });
  await host.request('close', { documentId });
  const reopened = await host.request('open', { documentId: randomUUID(), path: applied.checkpoint.path, directory: path.join(root, 'reopened'), generation: 2, visible: false });
  verifyEdits(reopened.snapshot);
  report.reopened = true; assert.equal(hash(await fs.readFile(source)), hash(fixture)); report.sourceUnchanged = true;
} catch (error) { report.error = { code: error.code, message: error.message, details: error.details, stack: error.stack }; process.exitCode = 1; }
finally {
  try {
    report.cleanup = await host.shutdown(Boolean(report.error));
    report.cleanup.workerExitCode = host.worker?.exitCode;
    assert.equal(report.cleanup.workerExited, true, 'Native worker must actually exit');
    assert.equal(report.cleanup.workerExitCode, 0, 'Native worker must exit without a native assertion');
    assert.deepEqual(report.cleanup.errors, [], 'Task documents must close successfully');
    if (report.cleanup.lease?.quitRequested) assert.equal(report.cleanup.officeExited, true, 'PowerPoint must actually exit after an owned application Quit');
  } catch (error) { report.cleanup = { ...report.cleanup, error: error.message }; process.exitCode = 1; }
  report.foreground = observer.stop(); job.close();
  await writeJson(path.join(root, 'report.json'), report);
  console.log(JSON.stringify({ root, error: report.error, cleanup: report.cleanup, foreground: report.foreground, preview: report.render?.images }, null, 2));
}
