import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { setTimeout as delay } from 'node:timers/promises';
import { startStdioPeer } from './helpers/stdio-peer.js';
import { buildDeck } from '../src/build.js';
import { FileLock, TaskHost } from '../src/task.js';
import { indexPackage } from '../src/file-engine.js';
import { readPackage } from '../src/ooxml.js';
import { readJson } from '../src/storage.js';
import { listProcesses, processIdentity, waitForExit } from '../src/windows.js';

async function fixture(t) {
  if (listProcesses('POWERPNT.EXE').length) { t.skip('An existing PowerPoint session must remain undisturbed.'); return null; }
  const root = path.resolve('work/tests/mcp-disconnect', randomUUID());
  const baseDir = path.join(root, 'tasks');
  await fs.mkdir(root, { recursive: true });
  const source = path.join(root, 'source.pptx');
  const original = await buildDeck({ slides: [{ items: [{ type: 'text', name: 'Heading', text: 'Before disconnect', left: 40, top: 40, width: 600, height: 80 }] }] });
  await fs.writeFile(source, original);
  const peer = await startStdioPeer(baseDir);
  t.after(() => peer.stop());
  const { taskId } = await peer.call('ppt_status');
  const taskDir = path.join(baseDir, taskId);
  const { documentId } = await peer.call('ppt_open', { path: source, mode: 'native-copy', allowOffice: true, operationId: 'open' });
  const inspected = await peer.call('ppt_inspect', { documentId });
  const target = inspected.objects.find(o => o.name === 'Heading');
  assert.ok(target);
  const operations = [{ type: 'replace_text', targetRef: target.targetRef, search: 'Before disconnect', replacement: 'After disconnect', expectedMatches: 1 }];
  const worker = (await readJson(path.join(taskDir, 'native-state.json'))).nativeWorker;
  const observedOffice = listProcesses('POWERPNT.EXE').map(p => processIdentity(p.pid)).filter(Boolean);
  assert.ok(observedOffice.length, 'Native workflow must really open PowerPoint.');
  return { peer, root, baseDir, taskId, taskDir, documentId, operations, worker, observedOffice, source, original };
}

async function assertDisconnected(state) {
  const exit = await state.peer.stop();
  assert.equal(exit.code, 0, exit.stderr);
  assert.equal(exit.stderr, '');
  assert.equal(await waitForExit(state.worker, 10000), true, 'Owned native worker must exit.');
  for (const identity of state.observedOffice) assert.equal(await waitForExit(identity, 10000), true, 'Observed cold-start Office must exit.');
  assert.deepEqual(await fs.readFile(state.source), state.original);
  const meta = await readJson(path.join(state.taskDir, 'documents', state.documentId, 'meta.json'));
  assert.equal(meta.revision, 1);
  const snapshot = indexPackage(await readPackage(await fs.readFile(meta.checkpointPath)));
  assert.equal(snapshot.objects.find(o => o.name === 'Heading').text, 'After disconnect');
  const status = await TaskHost.status(state.taskDir, { operationId: 'edit' });
  assert.equal(status.status, 'active', 'Disconnect must not fabricate review approval.');
  assert.equal(status.operationReceipt.status, 'completed');
  assert.equal(status.operationReceipt.result.revision, 1);
}

test('MCP stdin EOF closes idle native resources and reconnects from the durable revision', { skip: process.platform !== 'win32', timeout: 120000 }, async t => {
  const started = performance.now();
  const state = await fixture(t);
  if (!state) return;
  const args = { documentId: state.documentId, expectedRevision: 0, operationId: 'edit', operations: state.operations };
  await state.peer.call('ppt_apply', args);
  await assertDisconnected(state);
  const resumed = await startStdioPeer(state.baseDir, state.taskId);
  try {
    const replayed = await resumed.call('ppt_apply', args);
    assert.equal(replayed.revision, 1, 'Retry must replay the durable receipt, not apply twice.');
    const view = await resumed.call('ppt_inspect', { documentId: state.documentId });
    assert.equal(view.revision, 1);
    assert.equal(view.objects.find(o => o.name === 'Heading').text, 'After disconnect');
    const outputPath = path.join(state.root, 'resumed.pptx');
    const committed = await resumed.call('ppt_commit', { documentId: state.documentId, expectedRevision: 1, outputPath, operationId: 'commit' });
    assert.equal(indexPackage(await readPackage(await fs.readFile(outputPath))).objects.find(o => o.name === 'Heading').text, 'After disconnect');
    const finished = await resumed.call('ppt_finish', { reviewIds: [committed.reviewId] });
    assert.equal(finished.shutdownReport.workerExited, true);
    assert.deepEqual(finished.shutdownReport.errors, []);
    if (finished.shutdownReport.lease?.quitRequested) assert.equal(finished.shutdownReport.officeExited, true);
    assert.deepEqual(await fs.readFile(state.source), state.original);
  } finally { assert.equal((await resumed.stop()).code, 0); }
  t.diagnostic(`Idle EOF, reconnect and publish: ${Math.round(performance.now() - started)} ms`);
});

test('MCP stdin EOF drains an accepted native edit before cleanup and retains its receipt', { skip: process.platform !== 'win32', timeout: 120000 }, async t => {
  const started = performance.now();
  const state = await fixture(t);
  if (!state) return;
  const docDir = path.join(state.taskDir, 'documents', state.documentId);
  const barrier = await FileLock.acquire(path.join(docDir, 'mutation.lock'), { context: 'deterministic-eof-barrier' });
  let pending;
  try {
    pending = state.peer.call('ppt_apply', { documentId: state.documentId, expectedRevision: 0, operationId: 'edit', operations: state.operations });
    pending.catch(() => {});
    const opLock = path.join(state.taskDir, 'locks', 'op-edit.lock');
    const deadline = Date.now() + 5000;
    while (true) {
      const held = await readJson(opLock).catch(() => null);
      if (held?.pid === state.peer.child.pid) break;
      assert.ok(Date.now() < deadline, 'Request must reach the real mutation-lock boundary.');
      await delay(20);
    }
    assert.equal((await readJson(path.join(docDir, 'meta.json'))).revision, 0);
    state.peer.endInput();
  } finally { await FileLock.release(barrier); }
  assert.equal((await pending).revision, 1);
  await assertDisconnected(state);
  assert.deepEqual(await fs.readdir(path.join(state.taskDir, 'locks')), []);
  t.diagnostic(`In-flight EOF, durable receipt and cleanup: ${Math.round(performance.now() - started)} ms`);
});
