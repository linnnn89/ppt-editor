import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { Client, InMemoryTransport } from '@modelcontextprotocol/client';
import { createPptMcpServer } from '../src/mcp.js';
import { TaskHost } from '../src/task.js';
import { FileEngine } from '../src/file-engine.js';
import { readPackage } from '../src/ooxml.js';

// Run the published examples, so the tested workflow is the one agents will read.
const manual = await fs.readFile(new URL('../skills/ppt-editor/references/code-mode.md', import.meta.url), 'utf8');
const snippets = [...manual.matchAll(/```javascript\r?\n([\s\S]*?)```/g)].map(match => match[1]);
const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
const EXIT = Symbol('code-mode exit');
const deck = { width: 640, height: 360, slides: Array.from({ length: 3 }, (_, slide) => ({
  items: Array.from({ length: 3 }, (_, item) => ({ type: 'text', name: 'Shared name',
    text: slide === 1 && item === 2 ? 'Long text '.repeat(40) : `Before ${slide + 1}.${item + 1}`,
    left: 30, top: 30 + item * 90, width: 580, height: 70, fontSize: 18 }))
})) };

async function fixture() {
  const baseDir = path.resolve('work/tests/code-mode', randomUUID(), 'tasks');
  const host = new TaskHost({ baseDir });
  host.ensureNativeHost = async () => { throw Error('Code-mode file tests must not start Office'); };
  const instance = createPptMcpServer({ taskHost: host });
  const client = new Client({ name: 'code-mode-regression', version: '1' });
  const [ct, st] = InMemoryTransport.createLinkedPair();
  await instance.server.connect(st); await client.connect(ct);
  const state = new Map(), outputs = [], calls = [];
  const store = (key, value) => state.set(key, structuredClone(value));
  const load = key => structuredClone(state.get(key));
  const call = async (name, args) => {
    const response = await client.callTool({ name, arguments: args });
    assert(!response.isError, response.content[0].text);
    return response.structuredContent;
  };
  const tools = Object.fromEntries(['ppt_inspect', 'ppt_apply'].map(name => [
    `mcp__ppt_editor__${name}`, async args => {
      calls.push({ name, args: structuredClone(args) });
      return client.callTool({ name, arguments: args });
    }
  ]));
  const run = async index => {
    assert.equal(snippets.length, 2);
    try {
      await new AsyncFunction('tools', 'store', 'load', 'text', 'exit', snippets[index])(
        tools, store, load, value => outputs.push(structuredClone(value)), () => { throw EXIT; });
    } catch (error) { if (error !== EXIT) throw error; }
  };
  const { documentId } = await call('ppt_build', { deck, operationId: 'build' });
  store('ppt.inspectArgs', { documentId, kind: 'text', limit: 300 });
  return { host, client, tools, documentId, store, load, outputs, calls, call, run,
    async close() { await instance.cleanup(); await client.close(); await instance.server.close(); } };
}

test('published code-mode examples edit and commit a deck using a compact view and retained exact references', async () => {
  const f = await fixture();
  try {
    const session = await f.host.getSession(f.documentId), original = await fs.readFile(session.originalPath);
    // Clients that provide only MCP's JSON text fallback follow the same path.
    const inspect = f.tools.mcp__ppt_editor__ppt_inspect;
    f.tools.mcp__ppt_editor__ppt_inspect = async args => {
      const { structuredContent, ...response } = await inspect(args); return response;
    };
    await f.run(0);
    const snapshot = f.load('ppt.inspection'), view = f.outputs.at(-1);
    assert.equal(view.totalObjects, 9);
    assert.equal(view.objects.filter(object => object.textTruncated).length, 1);
    assert(view.objects.every(object => !('targetRef' in object)));
    for (let index = 0; index < snapshot.objects.length; index++) {
      const { targetRef, ...expected } = snapshot.objects[index];
      assert(targetRef.startsWith('ref_')); assert.deepEqual(view.objects[index], expected);
    }
    assert(JSON.stringify(view).length < JSON.stringify(snapshot).length * 0.65);
    const targets = ['Before 1.1', 'Before 3.2'];
    f.store('ppt.editPlan', { documentId: f.documentId, revision: 0, generation: 0, operationId: 'edit', edits: targets.map(text => ({
      key: view.objects.find(object => object.text === text).key,
      type: 'replace_text', search: 'Before', replacement: 'After', expectedMatches: 1
    })) });
    await f.run(1);
    const result = f.outputs.at(-1);
    assert.equal(result.revision, 1); assert.equal(f.load('ppt.inspection'), null);
    assert.deepEqual(f.calls.map(call => call.name), ['ppt_inspect', 'ppt_apply']);
    assert.deepEqual(f.load('ppt.applyCall').response.structuredContent, result);
    assert.deepEqual(result.layoutAudit.pendingSlides, []);
    const outputPath = path.join(f.host.taskDir, 'edited.pptx');
    await f.call('ppt_commit', { documentId: f.documentId, expectedRevision: 1, operationId: 'commit', outputPath });
    const bytes = await fs.readFile(outputPath);
    const reopened = new FileEngine(f.host.taskDir, bytes, await readPackage(bytes)).inspect();
    const texts = reopened.objects.filter(object => object.kind === 'text').map(object => object.text);
    assert(texts.includes('After 1.1')); assert(texts.includes('After 3.2'));
    assert(texts.includes('Before 1.2')); assert(texts.includes('Before 3.1'));
    assert.deepEqual(await fs.readFile(session.originalPath), original);
    f.store('ppt.inspectArgs', { documentId: f.documentId, text: 'Long text', detail: 'full' });
    await f.run(0);
    const fullText = f.outputs.at(-1).objects[0];
    assert.equal(fullText.text, 'Long text '.repeat(40), 'An explicit full-detail request must recover the truncated text');
    assert(fullText.runs.length > 0);
  } finally { await f.close(); }
});

test('code-mode target selection rejects missing or ambiguous keys and exposes a stale-revision error', async () => {
  const f = await fixture();
  try {
    await f.run(0);
    const snapshot = f.load('ppt.inspection'), target = snapshot.objects[0];
    const edit = { type: 'set_style', style: { bold: true } };
    const plan = { documentId: f.documentId, revision: 0, generation: 0, operationId: 'edit', edits: [{ key: target.key, ...edit }] };
    for (const mismatch of [{ documentId: randomUUID() }, { revision: 1 }, { generation: 1 }]) {
      f.store('ppt.editPlan', { ...plan, ...mismatch });
      await assert.rejects(f.run(1), /edit plan no longer matches/);
    }
    f.store('ppt.editPlan', { ...plan, edits: [{ key: 'missing', ...edit }] });
    await assert.rejects(f.run(1), /Inspect the exact target/);
    f.store('ppt.inspection', { ...snapshot, objects: [...snapshot.objects, target] });
    f.store('ppt.editPlan', plan);
    await assert.rejects(f.run(1), /Inspect the exact target/);
    assert.equal(f.calls.filter(call => call.name === 'ppt_apply').length, 0);
    f.store('ppt.inspection', snapshot);
    await f.call('ppt_apply', { documentId: f.documentId, expectedRevision: 0, operationId: 'intervening-edit',
      operations: [{ ...edit, targetRef: target.targetRef }] });
    await f.run(1);
    const response = f.outputs.at(-1);
    assert.equal(response.isError, true);
    assert.equal(JSON.parse(response.content[0].text).code, 'REVISION_MISMATCH');
    assert.equal(f.load('ppt.inspection'), null);
    assert.equal((await f.host.getSession(f.documentId)).revision, 1);
    assert.equal(f.calls.filter(call => call.name === 'ppt_apply').length, 1);
  } finally { await f.close(); }
});

test('a lost code-mode reply retains the exact request and recovers its durable receipt without another edit', async () => {
  const f = await fixture();
  try {
    await f.run(0);
    const target = f.load('ppt.inspection').objects.find(object => object.text === 'Before 1.1');
    f.store('ppt.editPlan', { documentId: f.documentId, revision: 0, generation: 0, operationId: 'lost-reply', edits: [{ key: target.key,
      type: 'replace_text', search: 'Before', replacement: 'After', expectedMatches: 1 }] });
    const apply = f.tools.mcp__ppt_editor__ppt_apply;
    f.tools.mcp__ppt_editor__ppt_apply = async args => { await apply(args); throw Error('Reply lost after execution'); };
    await assert.rejects(f.run(1), /Reply lost after execution/);
    const cached = f.load('ppt.applyCall');
    assert.equal(cached.state, 'pending'); assert.equal(f.load('ppt.inspection'), null);
    const recovered = await f.call('ppt_status', { taskId: f.host.taskId, operationId: cached.request.operationId });
    assert.equal(recovered.operationReceipt.result.revision, 1);
    assert.deepEqual(await f.call('ppt_apply', cached.request), recovered.operationReceipt.result);
    const session = await f.host.getSession(f.documentId), restored = await FileEngine.restore(session.engine.directory);
    assert.equal(restored.revision, 1);
    assert.equal(restored.inspect().objects.find(object => object.key === target.key).text, 'After 1.1');
  } finally { await f.close(); }
});
