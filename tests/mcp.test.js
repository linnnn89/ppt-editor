import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { Client } from '@modelcontextprotocol/client';
import { InMemoryTransport as ClientInMemoryTransport } from '@modelcontextprotocol/client';
import { StdioClientTransport } from '@modelcontextprotocol/client/stdio';
import { createPptMcpServer } from '../src/mcp.js';
import { buildDeck } from '../src/build.js';
import { readPackage } from '../src/ooxml.js';

const workDir = () => path.resolve('work/tests/mcp', randomUUID());

test('MCP native preflight preserves the session for editing, output readback and clean exit', { skip: process.platform !== 'win32', timeout: 120000 }, async () => {
  const root = workDir();
  await fs.mkdir(root, { recursive: true });
  const source = path.join(root, 'source.pptx');
  const original = await buildDeck({ slides: [{ items: [{ type: 'text', text: 'Before', left: 10, top: 10, width: 200, height: 40, name: 'Heading' }], notes: 'Synthetic note' }] });
  await fs.writeFile(source, original);
  const instance = createPptMcpServer({ baseDir: path.join(root, 'tasks') });
  const [ct, st] = ClientInMemoryTransport.createLinkedPair();
  const client = new Client({ name: 'native-session-regression', version: '1.0.0' });
  await instance.server.connect(st);
  await client.connect(ct);
  const call = async (name, args) => {
    const result = await client.callTool({ name, arguments: args });
    assert.ok(!result.isError, result.content[0].text);
    return JSON.parse(result.content[0].text);
  };
  try {
    const { documentId } = await call('ppt_open', { path: source, mode: 'native-copy', allowOffice: true, operationId: 'open' });
    const inspected = await call('ppt_inspect', { documentId });
    const heading = inspected.objects.find(o => o.name === 'Heading');
    const operations = [{ type: 'replace_text', targetRef: heading.targetRef, search: 'Before', replacement: 'After', expectedMatches: 1 }];
    const dry = await call('ppt_apply', { documentId, expectedRevision: 0, operationId: 'dry', dryRun: true, operations });
    assert.equal(dry.outcome, 'preflight_passed');
    assert.equal(dry.revision, 0);
    assert.equal(dry.generation, inspected.generation);
    assert.equal((await call('ppt_inspect', { documentId })).objects.find(o => o.name === 'Heading').text, 'Before');
    assert.equal((await call('ppt_apply', { documentId, expectedRevision: 0, operationId: 'apply', operations })).revision, 1);
    const outputPath = path.join(root, 'output.pptx');
    const committed = await call('ppt_commit', { documentId, expectedRevision: 1, operationId: 'commit', outputPath });
    const { indexPackage } = await import('../src/file-engine.js');
    assert.equal(indexPackage(await readPackage(await fs.readFile(outputPath))).objects.find(o => o.name === 'Heading').text, 'After');
    const readback = await call('ppt_open', { path: outputPath, mode: 'native-copy', allowOffice: true, operationId: 'readback' });
    assert.equal((await call('ppt_inspect', { documentId: readback.documentId })).objects.find(o => o.name === 'Heading').text, 'After');
    const finished = await call('ppt_finish', { reviewIds: [committed.reviewId] });
    assert.equal(finished.shutdownReport.workerExited, true);
    assert.deepEqual(finished.shutdownReport.errors, []);
    if (finished.shutdownReport.lease?.quitRequested) assert.equal(finished.shutdownReport.officeExited, true);
    assert.deepEqual(await fs.readFile(source), original);
  } finally { await instance.cleanup(); await client.close(); await instance.server.close(); }
});

test('MCP starts another task explicitly and retains the completed task receipts', async () => {
  const { server } = createPptMcpServer({ baseDir: workDir() });
  const [ct, st] = ClientInMemoryTransport.createLinkedPair();
  const client = new Client({ name: 'sequential-tasks', version: '1.0.0' });
  await server.connect(st);
  await client.connect(ct);
  const call = async (name, args = {}) => {
    const response = await client.callTool({ name, arguments: args });
    assert.ok(!response.isError, response.content[0].text);
    return JSON.parse(response.content[0].text);
  };
  try {
    const first = await call('ppt_status');
    const startArgs = { taskId: randomUUID() };
    const blocked = await client.callTool({ name: 'ppt_start', arguments: startArgs });
    assert.equal(blocked.isError, true);
    const deck = { slides: [{ items: [{ type: 'text', text: 'Synthetic task', left: 10, top: 10, width: 200, height: 40 }] }] };
    await call('ppt_build', { deck, operationId: 'build_first' });
    await call('ppt_finish', { reviewIds: [] });
    assert.equal((await call('ppt_status')).status, 'finished');
    const next = await call('ppt_start', startArgs);
    assert.equal(next.taskId, startArgs.taskId);
    assert.notEqual(next.taskId, first.taskId);
    assert.deepEqual(await call('ppt_start', startArgs), next, 'retry must not create another task');
    const built = await call('ppt_build', { deck, operationId: 'build_second' });
    assert.ok(built.documentId);
    const previous = await call('ppt_status', { taskId: first.taskId, operationId: 'build_first' });
    assert.equal(previous.status, 'finished');
    assert.equal(previous.operationReceipt.status, 'completed');
    await call('ppt_finish', { reviewIds: [] });
  } finally { await client.close(); await server.close(); }
});

test('MCP Server lists all 14 business tools with valid schemas and descriptions', async () => {
  const root = workDir();
  await fs.mkdir(root, { recursive: true });

  const { server } = createPptMcpServer({ baseDir: root });
  const [clientTransport, serverTransport] = ClientInMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);

  const client = new Client({ name: 'test-client', version: '1.0.0' });
  await client.connect(clientTransport);

  try {
    const list = await client.listTools();
    assert.equal(list.tools.length, 14);

    const expectedTools = [
      'ppt_start',
      'ppt_open',
      'ppt_inspect',
      'ppt_apply',
      'ppt_commit',
      'ppt_validate',
      'ppt_close',
      'ppt_finish',
      'ppt_status',
      'ppt_diagnose',
      'ppt_build',
      'ppt_compose',
      'ppt_relayout',
      'ppt_render'
    ];

    const actualToolNames = list.tools.map(t => t.name).sort();
    assert.deepEqual(actualToolNames, [...expectedTools].sort());

    for (const tool of list.tools) {
      assert.ok(tool.description && tool.description.length > 0, `Tool ${tool.name} missing description`);
      assert.equal(tool.inputSchema.type, 'object', `Tool ${tool.name} inputSchema must be object`);
    }
  } finally {
    await client.close();
    await server.close();
  }
});

test('MCP Client executes full file-mode editing lifecycle via tool calls', async () => {
  const root = workDir();
  await fs.mkdir(root, { recursive: true });

  // 1. 构建测试幻灯片
  const deck = {
    title: 'MCP Full Lifecycle Deck',
    slides: [
      {
        items: [
          { type: 'text', text: 'Original Headline for MCP', left: 40, top: 40, width: 500, height: 60, name: 'TitleShape' }
        ],
        notes: 'Original MCP Notes'
      }
    ]
  };
  const sourceFile = path.join(root, 'source.pptx');
  await fs.writeFile(sourceFile, await buildDeck(deck));

  const { server } = createPptMcpServer({ baseDir: root });
  const [clientTransport, serverTransport] = ClientInMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);

  const client = new Client({ name: 'test-client', version: '1.0.0' });
  await client.connect(clientTransport);

  try {
    // 2. ppt_open
    const openRes = await client.callTool({
      name: 'ppt_open',
      arguments: {
        path: sourceFile,
        mode: 'file',
        operationId: 'op_mcp_open_1'
      }
    });
    assert.equal(openRes.isError, undefined);
    const openData = JSON.parse(openRes.content[0].text);
    assert.ok(openData.documentId);
    assert.equal(openData.mode, 'file');
    assert.equal(openData.revision, 0);
    const documentId = openData.documentId;

    // 3. ppt_inspect
    const inspectRes = await client.callTool({
      name: 'ppt_inspect',
      arguments: {
        documentId
      }
    });
    assert.equal(inspectRes.isError, undefined);
    const inspectData = JSON.parse(inspectRes.content[0].text);
    assert.equal(inspectData.revision, 0);
    const titleShape = inspectData.objects.find(o => o.name === 'TitleShape');
    assert.ok(titleShape && titleShape.targetRef);

    // 4. ppt_apply
    const applyRes = await client.callTool({
      name: 'ppt_apply',
      arguments: {
        documentId,
        expectedRevision: 0,
        operationId: 'op_mcp_apply_1',
        operations: [
          {
            type: 'replace_text',
            targetRef: titleShape.targetRef,
            search: 'Original Headline',
            replacement: 'Modified Via MCP',
            expectedMatches: 1,
            crossRunPolicy: 'reject'
          }
        ]
      }
    });
    assert.equal(applyRes.isError, undefined);
    const applyData = JSON.parse(applyRes.content[0].text);
    assert.equal(applyData.outcome, 'completed');
    assert.equal(applyData.revision, 1);

    // 5. ppt_commit
    const outputFile = path.join(root, 'mcp_output.pptx');
    const commitRes = await client.callTool({
      name: 'ppt_commit',
      arguments: {
        documentId,
        expectedRevision: 1,
        outputPath: outputFile,
        operationId: 'op_mcp_commit_1'
      }
    });
    assert.equal(commitRes.isError, undefined);
    const commitData = JSON.parse(commitRes.content[0].text);
    assert.equal(commitData.outputPath, outputFile);
    assert.ok(commitData.reviewId);
    const reviewId = commitData.reviewId;

    // 6. ppt_validate
    const validateRes = await client.callTool({
      name: 'ppt_validate',
      arguments: {
        documentId
      }
    });
    assert.equal(validateRes.isError, undefined);
    const validateData = JSON.parse(validateRes.content[0].text);
    assert.equal(validateData.structural.passed, true);

    // 7. ppt_finish
    const finishRes = await client.callTool({
      name: 'ppt_finish',
      arguments: {
        reviewIds: [reviewId]
      }
    });
    assert.equal(finishRes.isError, undefined);
    const finishData = JSON.parse(finishRes.content[0].text);
    assert.equal(finishData.status, 'finished');

    // 8. 验证磁盘输出文件内容
    const outParts = await readPackage(await fs.readFile(outputFile));
    const s1Text = outParts.get('ppt/slides/slide1.xml').toString('utf8');
    assert.match(s1Text, /Modified Via MCP/);
  } finally {
    await client.close();
    await server.close();
  }
});

test('MCP Server wraps domain errors gracefully with isError flag and error details', async () => {
  const root = workDir();
  await fs.mkdir(root, { recursive: true });

  const deck = {
    title: 'Error Handling Deck',
    slides: [{ items: [{ type: 'text', text: 'Error Test Text', left: 20, top: 20, width: 300, height: 40 }] }]
  };
  const sourceFile = path.join(root, 'source.pptx');
  await fs.writeFile(sourceFile, await buildDeck(deck));

  const { server } = createPptMcpServer({ baseDir: root });
  const [clientTransport, serverTransport] = ClientInMemoryTransport.createLinkedPair();
  await server.connect(serverTransport);

  const client = new Client({ name: 'test-client', version: '1.0.0' });
  await client.connect(clientTransport);

  try {
    const openRes = await client.callTool({
      name: 'ppt_open',
      arguments: { path: sourceFile, mode: 'file', operationId: 'err_open_1' }
    });
    const { documentId } = JSON.parse(openRes.content[0].text);

    // 查询对象获取 targetRef
    const inspectRes = await client.callTool({
      name: 'ppt_inspect',
      arguments: { documentId }
    });
    const { objects } = JSON.parse(inspectRes.content[0].text);
    const targetRef = objects[0].targetRef;

    // 1. 传入错误修订号预期，验证业务级 REVISION_MISMATCH 错误包装
    const applyErrRes = await client.callTool({
      name: 'ppt_apply',
      arguments: {
        documentId,
        expectedRevision: 999, // 错误修订号
        operationId: 'err_apply_stale',
        operations: [
          {
            type: 'replace_text',
            targetRef,
            search: 'Error Test',
            replacement: 'New Text',
            expectedMatches: 1,
            crossRunPolicy: 'reject'
          }
        ]
      }
    });
    assert.equal(applyErrRes.isError, true);
    const errData = JSON.parse(applyErrRes.content[0].text);
    assert.equal(errData.code, 'REVISION_MISMATCH');

    // 2. 覆盖源文件另存拦截
    const commitErrRes = await client.callTool({
      name: 'ppt_commit',
      arguments: {
        documentId,
        expectedRevision: 0,
        outputPath: sourceFile, // 企图覆盖源文件
        operationId: 'err_commit_overwrite'
      }
    });
    assert.equal(commitErrRes.isError, true);
    const overwriteErr = JSON.parse(commitErrRes.content[0].text);
    assert.equal(overwriteErr.code, 'OVERWRITE_SOURCE_PROHIBITED');

    // 3. Schema 级别非法输入拦截（空 operations 列表违反 min(1) 约束）
    const schemaErrRes = await client.callTool({
      name: 'ppt_apply',
      arguments: {
        documentId,
        expectedRevision: 0,
        operationId: 'err_schema_empty',
        operations: []
      }
    });
    assert.equal(schemaErrRes.isError, true);
    assert.match(schemaErrRes.content[0].text, /Input validation error/i);
  } finally {
    await client.close();
    await server.close();
  }
});

test('MCP Client connects to standalone MCP server process via real StdioClientTransport', async () => {
  const root = workDir();
  await fs.mkdir(root, { recursive: true });

  const transport = new StdioClientTransport({
    command: process.execPath,
    args: ['src/mcp.js', '--base-dir', root]
  });

  const client = new Client({ name: 'stdio-test-client', version: '1.0.0' });
  await client.connect(transport);

  try {
    const tools = await client.listTools();
    assert.equal(tools.tools.length, 14);

    const diagnoseRes = await client.callTool({
      name: 'ppt_diagnose',
      arguments: {}
    });
    assert.equal(diagnoseRes.isError, undefined);
    const diag = JSON.parse(diagnoseRes.content[0].text);
    assert.equal(diag.status, 'ok');
    assert.ok(diag.taskId);
    assert.equal(diag.closed, false);

    const statusRes = await client.callTool({
      name: 'ppt_status',
      arguments: {}
    });
    assert.equal(statusRes.isError, undefined);
    const statusData = JSON.parse(statusRes.content[0].text);
    assert.equal(statusData.status, 'active');
    const finished = await client.callTool({ name: 'ppt_finish', arguments: { reviewIds: [] } });
    assert.ok(!finished.isError);
    const nextTaskId = randomUUID();
    const next = await client.callTool({ name: 'ppt_start', arguments: { taskId: nextTaskId } });
    assert.ok(!next.isError, next.content[0].text);
    assert.equal(JSON.parse(next.content[0].text).taskId, nextTaskId);
    const previous = await client.callTool({ name: 'ppt_status', arguments: { taskId: statusData.taskId } });
    assert.equal(JSON.parse(previous.content[0].text).status, 'finished');
  } finally {
    await client.close();
  }
});
