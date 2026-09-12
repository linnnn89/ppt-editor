import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { randomUUID } from 'node:crypto';
import { buildDeck } from '../src/build.js';
import { readPackage, validatePackage } from '../src/ooxml.js';
import { hash } from '../src/storage.js';

const execFileAsync = promisify(execFile);
const cliPath = path.resolve('src/cli.js');
const workDir = () => path.resolve('work/tests', randomUUID());

async function runCli(args, stdin = null) {
  try {
    const child = execFileAsync(process.execPath, [cliPath, ...args], {
      windowsHide: true,
      maxBuffer: 10 * 1024 * 1024
    });
    if (stdin && child.child.stdin) {
      child.child.stdin.write(stdin);
      child.child.stdin.end();
    }
    const { stdout, stderr } = await child;
    return { exitCode: 0, stdout: stdout.trim(), stderr: stderr.trim(), data: stdout.trim() ? JSON.parse(stdout.trim()) : null };
  } catch (error) {
    const stdout = error.stdout?.trim() || '';
    let parsed = null;
    try { parsed = JSON.parse(stdout); } catch {}
    return {
      exitCode: error.code || 1,
      stdout,
      stderr: error.stderr?.trim() || '',
      data: parsed,
      error
    };
  }
}

test('CLI bridge executes complete task lifecycle across separate invocations', async () => {
  const root = workDir();
  await fs.mkdir(root, { recursive: true });
  const tasksDir = path.join(root, 'tasks');

  const deck = {
    title: 'CLI E2E Presentation',
    slides: [
      {
        items: [
          { type: 'text', text: 'CLI Heading Text', left: 40, top: 40, width: 500, height: 60, name: 'Headline' }
        ]
      }
    ]
  };
  const sourceBytes = await buildDeck(deck);
  const sourceFile = path.join(root, 'source.pptx');
  await fs.writeFile(sourceFile, sourceBytes);

  // 1. start 命令：启动并获取 taskId
  const startRes = await runCli(['start', '--base-dir', tasksDir]);
  assert.equal(startRes.exitCode, 0, startRes.stderr);
  assert.equal(startRes.data.status, 'active');
  const taskId = startRes.data.taskId;
  assert.ok(taskId);

  // 2. open 命令：在任务中打开文稿
  const openRes = await runCli(['open', '--task', taskId, '--path', sourceFile, '--mode', 'file', '--op', 'cli_open_1', '--base-dir', tasksDir]);
  assert.equal(openRes.exitCode, 0, openRes.stderr);
  const docId = openRes.data.documentId;
  assert.ok(docId);
  assert.equal(openRes.data.revision, 0);

  // 3. inspect 命令：查询对象并获得 targetRef
  const inspectRes = await runCli(['inspect', '--task', taskId, '--doc', docId, '--base-dir', tasksDir]);
  assert.equal(inspectRes.exitCode, 0, inspectRes.stderr);
  assert.equal(inspectRes.data.revision, 0);
  const headline = inspectRes.data.objects.find(o => o.name === 'Headline');
  assert.ok(headline);
  assert.ok(headline.targetRef);

  // 4. apply 命令：批量预检并修改（支持 --operations 传递 JSON 字符串）
  const ops = [
    {
      type: 'replace_text',
      targetRef: headline.targetRef,
      search: 'CLI Heading',
      replacement: 'CLI Modified',
      expectedMatches: 1,
      crossRunPolicy: 'reject'
    }
  ];
  const applyRes = await runCli(['apply', '--task', taskId, '--doc', docId, '--rev', '0', '--op', 'cli_apply_1', '--operations', JSON.stringify(ops), '--base-dir', tasksDir]);
  assert.equal(applyRes.exitCode, 0, applyRes.stderr);
  assert.equal(applyRes.data.outcome, 'completed');
  assert.equal(applyRes.data.revision, 1);

  // 5. validate 命令：验证当前包结构
  const validateRes = await runCli(['validate', '--task', taskId, '--doc', docId, '--base-dir', tasksDir]);
  assert.equal(validateRes.exitCode, 0, validateRes.stderr);
  assert.equal(validateRes.data.structural.passed, true);

  // 6. commit 命令：发布到新文件
  const outputFile = path.join(root, 'cli_output.pptx');
  const commitRes = await runCli(['commit', '--task', taskId, '--doc', docId, '--rev', '1', '--out', outputFile, '--op', 'cli_commit_1', '--base-dir', tasksDir]);
  assert.equal(commitRes.exitCode, 0, commitRes.stderr);
  const reviewId = commitRes.data.reviewId;
  assert.ok(reviewId);
  assert.equal(commitRes.data.outputPath, outputFile);

  // 验证输出文件内容
  const outParts = await readPackage(await fs.readFile(outputFile));
  const s1Xml = outParts.get('ppt/slides/slide1.xml').toString('utf8');
  assert.match(s1Xml, /CLI Modified/);

  // 7. status 命令：离线查询任务状态与操作回执
  const statusRes = await runCli(['status', '--task', taskId, '--op', 'cli_commit_1', '--base-dir', tasksDir]);
  assert.equal(statusRes.exitCode, 0, statusRes.stderr);
  assert.equal(statusRes.data.status, 'active');
  assert.equal(statusRes.data.operationReceipt?.status, 'completed');

  // 8. finish 命令：核验 reviewId 并退出
  const finishRes = await runCli(['finish', '--task', taskId, '--reviews', JSON.stringify([reviewId]), '--base-dir', tasksDir]);
  assert.equal(finishRes.exitCode, 0, finishRes.stderr);
  assert.equal(finishRes.data.status, 'finished');
});

test('CLI bridge handles input via stdin for complex JSON payloads', async () => {
  const root = workDir();
  await fs.mkdir(root, { recursive: true });
  const tasksDir = path.join(root, 'tasks');

  const deck = {
    title: 'Stdin Test Deck',
    slides: [{ items: [{ type: 'text', text: 'Before Stdin Edit', left: 30, top: 30, width: 400, height: 50, name: 'T1' }] }]
  };
  const sourceFile = path.join(root, 'stdin_test.pptx');
  await fs.writeFile(sourceFile, await buildDeck(deck));

  const startRes = await runCli(['start', '--base-dir', tasksDir]);
  const taskId = startRes.data.taskId;

  const openRes = await runCli(['open', '--task', taskId, '--path', sourceFile, '--mode', 'file', '--op', 'stdin_open', '--base-dir', tasksDir]);
  const docId = openRes.data.documentId;

  const inspectRes = await runCli(['inspect', '--task', taskId, '--doc', docId, '--base-dir', tasksDir]);
  const t1 = inspectRes.data.objects.find(o => o.name === 'T1');

  // 通过 stdin 传入完整的 apply payload JSON
  const payload = {
    documentId: docId,
    expectedRevision: 0,
    operationId: 'stdin_apply_op',
    operations: [
      {
        type: 'replace_text',
        targetRef: t1.targetRef,
        search: 'Before Stdin Edit',
        replacement: 'After Stdin Edit',
        expectedMatches: 1,
        crossRunPolicy: 'reject'
      }
    ]
  };

  const applyRes = await runCli(['apply', '--task', taskId, '--base-dir', tasksDir, '--stdin'], JSON.stringify(payload));
  assert.equal(applyRes.exitCode, 0, applyRes.stderr);
  assert.equal(applyRes.data.outcome, 'completed');
  assert.equal(applyRes.data.revision, 1);
});

test('Scenario 3: CLI executes complete native-copy lifecycle across separate invocations with clean process exit', { skip: process.platform !== 'win32' }, async () => {
  const root = workDir();
  await fs.mkdir(root, { recursive: true });
  const tasksDir = path.join(root, 'tasks');

  // 记录调用前系统中的 PowerPoint 进程
  const { listProcesses } = await import('../src/windows.js');
  const initialOfficePids = new Set(listProcesses('POWERPNT.EXE').map(p => p.pid));

  const deck = {
    title: 'Native CLI Deck',
    slides: [
      {
        items: [
          { type: 'text', text: 'Original Native Headline', left: 50, top: 50, width: 500, height: 60, name: 'Headline' }
        ],
        notes: 'Original Native Notes'
      }
    ]
  };
  const sourceBytes = await buildDeck(deck);
  const sourceFile = path.join(root, 'native_source.pptx');
  await fs.writeFile(sourceFile, sourceBytes);
  const sourceHash = hash(sourceBytes);

  // 1. start 命令
  const startRes = await runCli(['start', '--base-dir', tasksDir]);
  assert.equal(startRes.exitCode, 0, startRes.stderr);
  assert.equal(startRes.data.status, 'active');
  const taskId = startRes.data.taskId;

  // 2. open 命令（原生副本模式 native-copy）
  const openRes = await runCli([
    'open',
    '--task', taskId,
    '--path', sourceFile,
    '--mode', 'native-copy',
    '--allow-office', 'true',
    '--op', 'cli_native_open',
    '--base-dir', tasksDir
  ]);
  assert.equal(openRes.exitCode, 0, openRes.stderr);
  assert.equal(openRes.data.mode, 'native-copy');
  assert.equal(openRes.data.revision, 0);
  const docId = openRes.data.documentId;
  assert.ok(docId);

  // 确认普通命令执行后，任务仍然为 active，且本命令启动的 Office 已退出
  const postOpenStatus = await runCli(['status', '--task', taskId, '--base-dir', tasksDir]);
  assert.equal(postOpenStatus.data.status, 'active');
  const openPids = new Set(listProcesses('POWERPNT.EXE').map(p => p.pid));
  assert.deepEqual(Array.from(openPids).sort(), Array.from(initialOfficePids).sort());

  // 3. inspect 命令（跨进程由检查点自动恢复）
  const inspectRes = await runCli([
    'inspect',
    '--task', taskId,
    '--doc', docId,
    '--base-dir', tasksDir
  ]);
  assert.equal(inspectRes.exitCode, 0, inspectRes.stderr);
  assert.equal(inspectRes.data.revision, 0);
  const headline = inspectRes.data.objects.find(o => o.name === 'Headline');
  assert.ok(headline);
  assert.ok(headline.targetRef);
  const notesObj = inspectRes.data.objects.find(o => o.kind === 'notes');
  assert.ok(notesObj);
  assert.ok(notesObj.targetRef);
  const bgObj = inspectRes.data.objects.find(o => o.kind === 'background');
  assert.ok(bgObj);
  assert.ok(bgObj.targetRef);
  const inspectPids = new Set(listProcesses('POWERPNT.EXE').map(p => p.pid));
  assert.deepEqual(Array.from(inspectPids).sort(), Array.from(initialOfficePids).sort());

  // 4. apply 命令（由检查点恢复，修改文本、备注和不对称背景色，保存新检查点）
  const ops = [
    {
      type: 'replace_text',
      targetRef: headline.targetRef,
      search: 'Original Native Headline',
      replacement: 'Modified Native Headline',
      expectedMatches: 1,
      crossRunPolicy: 'reject'
    },
    {
      type: 'replace_text',
      targetRef: notesObj.targetRef,
      search: 'Original Native Notes',
      replacement: 'Modified Native Notes',
      expectedMatches: 1,
      crossRunPolicy: 'reject'
    },
    {
      type: 'set_slide_background',
      targetRef: bgObj.targetRef,
      color: '123456'
    }
  ];
  const applyRes = await runCli([
    'apply',
    '--task', taskId,
    '--doc', docId,
    '--rev', '0',
    '--op', 'cli_native_apply',
    '--operations', JSON.stringify(ops),
    '--base-dir', tasksDir
  ]);
  assert.equal(applyRes.exitCode, 0, applyRes.stderr);
  assert.equal(applyRes.data.outcome, 'completed');
  assert.equal(applyRes.data.revision, 1);
  const applyPids = new Set(listProcesses('POWERPNT.EXE').map(p => p.pid));
  assert.deepEqual(Array.from(applyPids).sort(), Array.from(initialOfficePids).sort());

  // 5. validate 命令（真实验证检查点及原生读回）
  const validateRes = await runCli([
    'validate',
    '--task', taskId,
    '--doc', docId,
    '--base-dir', tasksDir
  ]);
  assert.equal(validateRes.exitCode, 0, validateRes.stderr);
  assert.equal(validateRes.data.structural.passed, true);
  assert.equal(validateRes.data.nativeReadback, 'passed');
  const validatePids = new Set(listProcesses('POWERPNT.EXE').map(p => p.pid));
  assert.deepEqual(Array.from(validatePids).sort(), Array.from(initialOfficePids).sort());

  // 6. commit 命令（另存到新文件并验证）
  const outputFile = path.join(root, 'native_committed_output.pptx');
  const commitRes = await runCli([
    'commit',
    '--task', taskId,
    '--doc', docId,
    '--rev', '1',
    '--out', outputFile,
    '--op', 'cli_native_commit',
    '--base-dir', tasksDir
  ]);
  assert.equal(commitRes.exitCode, 0, commitRes.stderr);
  const reviewId = commitRes.data.reviewId;
  assert.ok(reviewId);
  assert.equal(commitRes.data.outputPath, outputFile);
  const commitPids = new Set(listProcesses('POWERPNT.EXE').map(p => p.pid));
  assert.deepEqual(Array.from(commitPids).sort(), Array.from(initialOfficePids).sort());

  // 验证输出文件内容与原稿未修改
  const outParts = await readPackage(await fs.readFile(outputFile));
  const s1Xml = outParts.get('ppt/slides/slide1.xml').toString('utf8');
  assert.match(s1Xml, /Modified Native Headline/);
  assert.doesNotMatch(s1Xml, /Original Native Headline/);
  assert.match(s1Xml, /123456/);

  const notesXml = outParts.get('ppt/notesSlides/notesSlide1.xml').toString('utf8');
  assert.match(notesXml, /Modified Native Notes/);
  assert.doesNotMatch(notesXml, /Original Native Notes/);
  assert.equal(hash(await fs.readFile(sourceFile)), sourceHash);

  // 验证输出文件包结构有效
  const outValidation = validatePackage(outParts);
  assert.equal(outValidation.passed, true);

  // 7. finish 命令（核验 reviewId 并退出）
  const finishRes = await runCli([
    'finish',
    '--task', taskId,
    '--reviews', JSON.stringify([reviewId]),
    '--base-dir', tasksDir
  ]);
  assert.equal(finishRes.exitCode, 0, finishRes.stderr);
  assert.equal(finishRes.data.status, 'finished');

  // 8. 最终核实：所有自有 Office 与 Node 资源彻底释放，系统进程无任何泄露
  const finalPids = new Set(listProcesses('POWERPNT.EXE').map(p => p.pid));
  assert.deepEqual(Array.from(finalPids).sort(), Array.from(initialOfficePids).sort());

  // 离线查询任务状态
  const finalStatus = await runCli(['status', '--task', taskId, '--op', 'cli_native_commit', '--base-dir', tasksDir]);
  assert.equal(finalStatus.data.status, 'finished');
  assert.equal(finalStatus.data.operationReceipt?.status, 'completed');
});

test('Scenario 4: TaskHost/CLI in native-copy mode strictly preserves pre-existing external PowerPoint sessions and presentations', { skip: process.platform !== 'win32' }, async (t) => {
  const { listProcesses, processIdentity, waitForExit, pumpMessages } = await import('../src/windows.js');
  if (listProcesses('POWERPNT.EXE').length) {
    t.skip('An actual external Office session is present; do not create or clean up a shared test application.');
    return;
  }
  const root = workDir();
  assert.equal(typeof globalThis.gc, 'function', 'Run via npm test to release temporary fixture COM references.');
  await fs.mkdir(root, { recursive: true });
  const tasksDir = path.join(root, 'tasks');

  // 1. 模拟用户预先存在且正在运行的外部 PowerPoint 进程与未保存文稿
  const { createRequire } = await import('node:module');
  const require = createRequire(import.meta.url);
  const winax = require('winax');
  const externalApp = new winax.Object('PowerPoint.Application', { activate: false });
  const externalPresentations = externalApp.Presentations;
  const externalPres = externalPresentations.Add(0);
  assert.equal(Number(externalPresentations.Count), 1);

  const externalPids = new Set(listProcesses('POWERPNT.EXE').map(p => p.pid));
  const externalIdentities = [...externalPids].map(processIdentity).filter(Boolean);
  assert.ok(externalPids.size >= 1, 'Pre-existing PowerPoint process must be present');

  try {
    const deck = {
      title: 'Task Deck',
      slides: [{ items: [{ type: 'text', text: 'Task Slide Text', left: 50, top: 50, width: 400, height: 50, name: 'T1' }] }]
    };
    const sourceBytes = await buildDeck(deck);
    const sourceFile = path.join(root, 'task_source.pptx');
    await fs.writeFile(sourceFile, sourceBytes);

    // 2. 执行 CLI start & open（原生副本模式）
    const startRes = await runCli(['start', '--base-dir', tasksDir]);
    assert.equal(startRes.exitCode, 0, startRes.stderr);
    const taskId = startRes.data.taskId;

    const openRes = await runCli([
      'open',
      '--task', taskId,
      '--path', sourceFile,
      '--mode', 'native-copy',
      '--allow-office', 'true',
      '--op', 'ext_test_open',
      '--base-dir', tasksDir
    ]);
    assert.equal(openRes.exitCode, 0, openRes.stderr);
    const docId = openRes.data.documentId;

    // 3. inspect
    const inspectRes = await runCli(['inspect', '--task', taskId, '--doc', docId, '--base-dir', tasksDir]);
    assert.equal(inspectRes.exitCode, 0, inspectRes.stderr);
    const t1 = inspectRes.data.objects.find(o => o.name === 'T1');
    assert.ok(t1 && t1.targetRef);

    // 4. apply
    const applyRes = await runCli([
      'apply',
      '--task', taskId,
      '--doc', docId,
      '--rev', '0',
      '--op', 'ext_test_apply',
      '--operations', JSON.stringify([{ type: 'replace_text', targetRef: t1.targetRef, search: 'Task Slide Text', replacement: 'Updated Text', expectedMatches: 1, crossRunPolicy: 'reject' }]),
      '--base-dir', tasksDir
    ]);
    assert.equal(applyRes.exitCode, 0, applyRes.stderr);

    // 5. commit
    const outputFile = path.join(root, 'task_committed.pptx');
    const commitRes = await runCli([
      'commit',
      '--task', taskId,
      '--doc', docId,
      '--rev', '1',
      '--out', outputFile,
      '--op', 'ext_test_commit',
      '--base-dir', tasksDir
    ]);
    assert.equal(commitRes.exitCode, 0, commitRes.stderr);

    // 6. finish
    const finishRes = await runCli([
      'finish',
      '--task', taskId,
      '--reviews', JSON.stringify([commitRes.data.reviewId]),
      '--base-dir', tasksDir
    ]);
    assert.equal(finishRes.exitCode, 0, finishRes.stderr);

    // 7. 严格核实：外部 PowerPoint 会话完全存活，外部文稿未被关闭！
    assert.equal(Number(externalPresentations.Count), 1, 'External presentation must remain open');
    const currentPids = new Set(listProcesses('POWERPNT.EXE').map(p => p.pid));
    for (const pid of externalPids) {
      assert.ok(currentPids.has(pid), `External PowerPoint process ${pid} must not be terminated`);
    }
  } finally {
    // 测试结束后显式释放并彻底关闭外部模拟会话，避免污染后续测试环境
    try {
      externalPres.Close();
      externalApp.Quit();
    } catch {}
    winax.release(externalPres, externalPresentations, externalApp);
    // Match the existing lifecycle harness: temporary winax wrappers also hold COM references.
    const pump = setInterval(pumpMessages, 10);
    try {
      await new Promise(resolve => setImmediate(resolve));
      globalThis.gc();
      for (const identity of externalIdentities) assert.equal(await waitForExit(identity, 5000), true, 'Fixture Office must exit through COM cleanup; never kill processes by enumeration.');
    } finally { clearInterval(pump); }
  }
});
