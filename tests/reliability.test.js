import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { TaskHost } from '../src/task.js';
import { buildDeck } from '../src/build.js';
import { readJson, writeJson, hash } from '../src/storage.js';

const deck = { slides: [{ items: [{ type: 'text', text: 'Before', name: 'Heading', left: 20, top: 20, width: 300, height: 60 }] }] };
async function fixture() {
  const root = path.resolve('work/tests/reliability', randomUUID());
  await fs.mkdir(root, { recursive: true });
  const source = path.join(root, 'source.pptx'), bytes = await buildDeck(deck);
  await fs.writeFile(source, bytes);
  const host = new TaskHost({ baseDir: path.join(root, 'tasks') });
  const opened = await host.open({ path: source, mode: 'file', allowOffice: true, operationId: 'open' });
  return { root, source, bytes, host, documentId: opened.documentId };
}

test('failed native close or shutdown preserves recovery data and never finishes the task', async () => {
  for (const closeFails of [true, false]) {
    const { host, documentId, bytes } = await fixture();
    const session = host.sessions.get(documentId), metaPath = path.join(session.directory, 'meta.json');
    const checkpoint = path.join(session.directory, 'native-1-0.pptx');
    await fs.writeFile(checkpoint, bytes);
    Object.assign(session, { mode: 'native-copy', generation: 1, checkpointPath: checkpoint });
    const meta = { ...await readJson(metaPath), mode: 'native-copy', generation: 1, checkpointPath: checkpoint };
    await writeJson(metaPath, meta);
    host.nativeHost = {
      request: async () => {
        if (closeFails) throw Object.assign(new Error('Synthetic close failure'), { code: 'NATIVE_CLOSE_FAILED' });
        return { closed: true };
      },
      shutdown: async () => ({ workerExited: true, officeExited: false, cleanup: 'document_cleanup_failed', errors: [{ code: 'SYNTHETIC_CLEANUP_FAILURE' }] })
    };
    if (closeFails) {
      await assert.rejects(host.close({ documentId }), { code: 'NATIVE_CLOSE_FAILED' });
      assert.deepEqual(await readJson(metaPath), meta);
      assert.equal(session.closed, false);
    }
    await assert.rejects(host.finish({ reviewIds: [] }), { code: 'CLEANUP_UNCONFIRMED' });
    assert.deepEqual(await fs.readFile(checkpoint), bytes);
    assert.equal((await readJson(metaPath)).checkpointPath, checkpoint);
    assert.equal((await TaskHost.status(host.taskDir)).status, 'active');
    assert.equal(host.closed, false);
    const restored = new TaskHost({ taskId: host.taskId, baseDir: host.baseDir });
    await assert.rejects(restored.finish({ reviewIds: [] }), { code: 'CLEANUP_UNCONFIRMED' });
    assert.deepEqual(await fs.readFile(checkpoint), bytes);
  }
});

test('cleanup rejects unknown outcomes persistently while accepting confirmed exit and preserved external Office', async () => {
  const unused = new TaskHost({ baseDir: path.resolve('work/tests/reliability', randomUUID()) });
  assert.deepEqual(await unused.cleanupCommand(), { workerExited: true, office: 'not_started' });
  await assert.rejects(fs.stat(unused.taskDir), { code: 'ENOENT' });
  const rejected = [
    { workerExited: true, officeExited: null, cleanup: 'worker_not_responsive', outcome: 'outcome_unknown' },
    { workerExited: true, officeExited: false, cleanup: 'office_exit_unconfirmed', lease: { quitRequested: true } },
    { workerExited: true, officeExited: null }
  ];
  for (const report of rejected) {
    const { host } = await fixture();
    host.nativeHost = { shutdown: async () => report };
    await assert.rejects(host.cleanupCommand(), { code: 'CLEANUP_UNCONFIRMED' });
    assert.deepEqual((await TaskHost.status(host.taskDir)).cleanupReport, report);
    await assert.rejects(host.cleanupCommand(), { code: 'CLEANUP_UNCONFIRMED' });
  }
  for (const report of [
    { workerExited: true, office: 'not_started' },
    { workerExited: true, officeExited: true, cleanup: 'completed', errors: [], lease: { quitRequested: true } },
    { workerExited: true, officeExited: false, cleanup: 'application_preserved', errors: [], lease: { quitRequested: false } }
  ]) {
    const { host } = await fixture();
    host.nativeHost = { shutdown: async () => report };
    assert.deepEqual(await host.cleanupCommand(), report);
    assert.equal((await host.finish({ reviewIds: [] })).status, 'finished');
  }
});

test('file edits receive real read-only native readback of current bytes without changing revision or source', { skip: process.platform !== 'win32', timeout: 120000 }, async t => {
  const { listProcesses } = await import('../src/windows.js');
  if (listProcesses('POWERPNT.EXE').length) return t.skip('Requires an isolated cold PowerPoint activation.');
  const { host, source, bytes, documentId, root } = await fixture();
  try {
    await assert.rejects(host.validate({ documentId, nativeReadback: true }), { code: 'OFFICE_NOT_ALLOWED' });
    assert.equal(host.nativeHost, null);
    const denied = await host.open({ path: source, mode: 'file', operationId: 'open_without_office' });
    await assert.rejects(host.validate({ documentId: denied.documentId, nativeReadback: true, allowOffice: true }), { code: 'OFFICE_NOT_ALLOWED' });
    assert.equal(host.nativeHost, null);
    const before = await host.inspect({ documentId });
    const writer = new TaskHost({ taskId: host.taskId, baseDir: host.baseDir });
    await writer.apply({ documentId, expectedRevision: 0, operationId: 'edit', operations: [
      { type: 'replace_text', targetRef: before.objects.find(o => o.name === 'Heading').targetRef, search: 'Before', replacement: 'After', expectedMatches: 1 }
    ] });
    const result = await host.validate({ documentId, nativeReadback: true, allowOffice: true });
    assert.equal(result.nativeReadback, 'passed');
    assert.equal(result.nativeValidation.revision, 1);
    assert.equal(result.nativeValidation.readOnly, true);
    assert.equal(result.nativeValidation.snapshot.objects.find(o => o.name === 'Heading').text, 'After');
    const after = await host.inspect({ documentId });
    assert.equal(after.revision, 1);
    assert.equal(after.generation, before.generation);
    const committed = await host.commit({ documentId, expectedRevision: 1, operationId: 'commit', outputPath: path.join(root, 'output.pptx') });
    assert.equal(result.nativeValidation.sha256, hash(await fs.readFile(committed.outputPath)));
    assert.deepEqual(await fs.readFile(source), bytes);
    const finished = await host.finish({ reviewIds: [committed.reviewId] });
    assert.equal(finished.shutdownReport.workerExited, true);
    assert.equal(finished.shutdownReport.officeExited, true);
    assert.equal(finished.shutdownReport.cleanup, 'completed');
  } finally { await host.cleanupCommand(); }
});
