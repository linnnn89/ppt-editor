import test from 'node:test';
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { fork } from 'node:child_process';
import { buildDeck } from '../src/build.js';
import { TaskHost, encodeTargetRef, decodeTargetRef, assertTaskDirBoundary } from '../src/task.js';
import { readPackage } from '../src/ooxml.js';
import { hash, writeJson, stableJson } from '../src/storage.js';

const workDir = () => path.resolve('work/tests', randomUUID());

const sampleDeck = {
  title: 'TaskHost Test Deck',
  slides: [
    {
      items: [
        { type: 'text', text: 'Original Title Heading', left: 50, top: 50, width: 600, height: 80, name: 'Title' },
        { type: 'text', text: 'Second Paragraph Item', left: 50, top: 150, width: 600, height: 80, name: 'Body' }
      ],
      notes: 'Presenter notes'
    }
  ]
};

test('targetRef encoding and decoding binds full context', () => {
  const context = {
    taskId: randomUUID(),
    documentId: randomUUID(),
    backend: 'file',
    revision: 2,
    generation: 0,
    slideId: 256,
    part: 'ppt/slides/slide1.xml',
    shapeId: 3,
    groupPath: [10, 20],
    fingerprint: 'a1b2c3d4e5f6'
  };

  const ref = encodeTargetRef(context);
  assert.ok(typeof ref === 'string');
  assert.ok(ref.length >= 20 && ref.length <= 6000);

  const decoded = decodeTargetRef(ref);
  assert.equal(decoded.taskId, context.taskId);
  assert.equal(decoded.documentId, context.documentId);
  assert.equal(decoded.backend, context.backend);
  assert.equal(decoded.revision, context.revision);
  assert.equal(decoded.generation, context.generation);
  assert.equal(decoded.slideId, context.slideId);
  assert.equal(decoded.part, context.part);
  assert.equal(decoded.shapeId, context.shapeId);
  assert.deepEqual(decoded.groupPath, context.groupPath);
  assert.equal(decoded.fingerprint, context.fingerprint);
  assert.equal(decoded.key, 'slide:256:10.20:3');

  // 验证原生备注对象（part 为 null，kind 为 notes）的编解码与 key 生成正确性
  const nativeNotesContext = {
    taskId: context.taskId,
    documentId: context.documentId,
    backend: 'native-copy',
    revision: 1,
    generation: 1,
    slideId: 256,
    part: null,
    shapeId: 2,
    groupPath: [],
    fingerprint: 'notes_fp_123',
    kind: 'notes'
  };
  const nativeNotesRef = encodeTargetRef(nativeNotesContext);
  const decodedNativeNotes = decodeTargetRef(nativeNotesRef);
  assert.equal(decodedNativeNotes.kind, 'notes');
  assert.equal(decodedNativeNotes.key, 'notes:256::2');

  // 验证背景对象（kind 为 background）的编解码与 key 生成正确性
  const bgContext = {
    taskId: context.taskId,
    documentId: context.documentId,
    backend: 'native-copy',
    revision: 1,
    generation: 1,
    slideId: 256,
    part: null,
    shapeId: 0,
    groupPath: [],
    fingerprint: 'bg_fp_456',
    kind: 'background'
  };
  const bgRef = encodeTargetRef(bgContext);
  const decodedBg = decodeTargetRef(bgRef);
  assert.equal(decodedBg.kind, 'background');
  assert.equal(decodedBg.key, 'background:256::0');
});

test('TaskHost executes full file-mode edit lifecycle with source protection', async () => {
  const root = workDir();
  const sourceBytes = await buildDeck(sampleDeck);
  const sourceHash = hash(sourceBytes);
  const sourceFile = path.join(root, 'input.pptx');
  await fs.mkdir(root, { recursive: true });
  await fs.writeFile(sourceFile, sourceBytes);

  const host = new TaskHost({ baseDir: path.join(root, 'tasks') });

  // 1. 打开文稿
  const opened = await host.open({
    path: sourceFile,
    mode: 'file',
    operationId: 'op_open_deck'
  });
  assert.equal(opened.mode, 'file');
  assert.equal(opened.revision, 0);
  assert.equal(opened.sourceHash, sourceHash);

  // 2. 查询对象并获取 targetRef
  const inspected = await host.inspect({
    documentId: opened.documentId
  });
  assert.equal(inspected.revision, 0);
  const titleObj = inspected.objects.find(o => o.name === 'Title');
  assert.ok(titleObj);
  assert.ok(titleObj.targetRef);

  // 3. 执行修改（预检 + 修改）
  const applied = await host.apply({
    documentId: opened.documentId,
    expectedRevision: 0,
    operationId: 'op_apply_title',
    operations: [
      {
        type: 'replace_text',
        targetRef: titleObj.targetRef,
        search: 'Original Title',
        replacement: 'Revised Title',
        expectedMatches: 1,
        crossRunPolicy: 'reject'
      }
    ]
  });
  assert.equal(applied.outcome, 'completed');
  assert.equal(applied.revision, 1);

  // 验证原稿绝对不受影响
  assert.equal(hash(await fs.readFile(sourceFile)), sourceHash);

  // 4. 验证结构
  const validation = await host.validate({
    documentId: opened.documentId
  });
  assert.equal(validation.structural.passed, true);

  // 5. 提交另存到新文件
  const outputFile = path.join(root, 'output.pptx');
  const committed = await host.commit({
    documentId: opened.documentId,
    expectedRevision: 1,
    outputPath: outputFile,
    operationId: 'op_commit_revised'
  });
  assert.ok(committed.reviewId);
  assert.equal(committed.revision, 1);
  assert.equal(committed.outputPath, outputFile);
  assert.equal(committed.reviewBundle.validation.basis, 'generated-output-bytes');

  // 验证输出文件可被正常读取且内容已被修改
  const outParts = await readPackage(await fs.readFile(outputFile));
  const slide1Xml = outParts.get('ppt/slides/slide1.xml').toString('utf8');
  assert.match(slide1Xml, /Revised Title/);
  assert.doesNotMatch(slide1Xml, /Original Title/);

  // 原稿依然保持不变
  assert.equal(hash(await fs.readFile(sourceFile)), sourceHash);

  // 6. 完工并退出
  const finished = await host.finish({
    reviewIds: [committed.reviewId],
    preserveCheckpoints: true
  });
  assert.equal(finished.status, 'finished');
  assert.equal(finished.reviewedCount, 1);
});

test('TaskHost enforces revision checks, rejects stale references and prohibits source overwrite', async () => {
  const root = workDir();
  const sourceBytes = await buildDeck(sampleDeck);
  const sourceFile = path.join(root, 'input.pptx');
  await fs.mkdir(root, { recursive: true });
  await fs.writeFile(sourceFile, sourceBytes);

  const host = new TaskHost({ baseDir: path.join(root, 'tasks') });
  const opened = await host.open({ path: sourceFile, mode: 'file', operationId: 'op_open_check' });
  const inspectedRev0 = await host.inspect({ documentId: opened.documentId });
  const titleObjRev0 = inspectedRev0.objects.find(o => o.name === 'Title');

  // 第一次修改：从 revision 0 -> 1
  await host.apply({
    documentId: opened.documentId,
    expectedRevision: 0,
    operationId: 'op_mod_1',
    operations: [
      {
        type: 'replace_text',
        targetRef: titleObjRev0.targetRef,
        search: 'Original Title',
        replacement: 'First Edit',
        expectedMatches: 1,
        crossRunPolicy: 'reject'
      }
    ]
  });

  // 尝试用错误的 expectedRevision 修改 -> 预期 REVISION_MISMATCH
  await assert.rejects(
    host.apply({
      documentId: opened.documentId,
      expectedRevision: 0, // 当前应为 1
      operationId: 'op_mod_bad_rev',
      operations: [
        {
          type: 'replace_text',
          targetRef: titleObjRev0.targetRef,
          search: 'First Edit',
          replacement: 'Second Edit',
          expectedMatches: 1,
          crossRunPolicy: 'reject'
        }
      ]
    }),
    { code: 'REVISION_MISMATCH' }
  );

  // 尝试用基于 revision 0 签发的 stale targetRef 修改（即使 expectedRevision 传 1） -> 预期 TARGET_STALE
  await assert.rejects(
    host.apply({
      documentId: opened.documentId,
      expectedRevision: 1,
      operationId: 'op_mod_stale_ref',
      operations: [
        {
          type: 'replace_text',
          targetRef: titleObjRev0.targetRef, // 过期的引用
          search: 'First Edit',
          replacement: 'Second Edit',
          expectedMatches: 1,
          crossRunPolicy: 'reject'
        }
      ]
    }),
    { code: 'TARGET_STALE' }
  );

  // 尝试将输出覆盖回原稿 -> 预期 OVERWRITE_SOURCE_PROHIBITED
  await assert.rejects(
    host.commit({
      documentId: opened.documentId,
      expectedRevision: 1,
      outputPath: sourceFile,
      operationId: 'op_commit_overwrite'
    }),
    { code: 'OVERWRITE_SOURCE_PROHIBITED' }
  );
});

test('TaskHost handles idempotency and conflicts on operationId', async () => {
  const root = workDir();
  const sourceBytes = await buildDeck(sampleDeck);
  const sourceFile = path.join(root, 'input.pptx');
  await fs.mkdir(root, { recursive: true });
  await fs.writeFile(sourceFile, sourceBytes);

  const host = new TaskHost({ baseDir: path.join(root, 'tasks') });
  const openParams = { path: sourceFile, mode: 'file', operationId: 'op_idempotent_open' };

  // 第一次调用 open
  const first = await host.open(openParams);
  // 第二次以相同参数调用 open，应直接返回相同的 documentId，不产生重复创建
  const second = await host.open(openParams);
  assert.equal(first.documentId, second.documentId);

  // 以相同 operationId 但不同参数调用 open，应拒绝 IDEMPOTENCY_CONFLICT
  await assert.rejects(
    host.open({ ...openParams, visible: true }),
    { code: 'IDEMPOTENCY_CONFLICT' }
  );

  // 离线查询状态
  const status = await TaskHost.status(host.taskDir, { operationId: 'op_idempotent_open' });
  assert.equal(status.status, 'active');
  assert.equal(status.operationReceipt.status, 'completed');
  assert.equal(status.operationReceipt.result.documentId, first.documentId);
});

test('TaskHost hydrates session across separate instances and checks source existence', async () => {
  const root = workDir();
  const sourceBytes = await buildDeck(sampleDeck);
  const sourceFile = path.join(root, 'input.pptx');
  await fs.mkdir(root, { recursive: true });
  await fs.writeFile(sourceFile, sourceBytes);

  const taskId = randomUUID();
  const baseDir = path.join(root, 'tasks');

  // 第一个 TaskHost 实例：执行打开与第一次编辑
  const host1 = new TaskHost({ taskId, baseDir });
  const opened = await host1.open({ path: sourceFile, mode: 'file', operationId: 'op_h1_open' });
  const docId = opened.documentId;
  const inspected1 = await host1.inspect({ documentId: docId });
  const titleObj = inspected1.objects.find(o => o.name === 'Title');

  await host1.apply({
    documentId: docId,
    expectedRevision: 0,
    operationId: 'op_h1_mod',
    operations: [{ type: 'replace_text', targetRef: titleObj.targetRef, search: 'Original Title', replacement: 'Instance1 Edit', expectedMatches: 1, crossRunPolicy: 'reject' }]
  });

  // 第二个全新的 TaskHost 实例（模拟下一个短命 CLI 进程）：以相同 taskId 启动
  const host2 = new TaskHost({ taskId, baseDir });
  // 内存中最初没有该 session，应从磁盘检查点与 meta.json 自动恢复 (hydrate)
  const inspected2 = await host2.inspect({ documentId: docId });
  assert.equal(inspected2.revision, 1);
  const titleObj2 = inspected2.objects.find(o => o.name === 'Title');
  assert.match(titleObj2.text, /Instance1 Edit/);

  // 在 host2 继续进行第二次编辑
  const applied2 = await host2.apply({
    documentId: docId,
    expectedRevision: 1,
    operationId: 'op_h2_mod',
    operations: [{ type: 'replace_text', targetRef: titleObj2.targetRef, search: 'Instance1 Edit', replacement: 'Instance2 Edit', expectedMatches: 1, crossRunPolicy: 'reject' }]
  });
  assert.equal(applied2.revision, 2);

  // 校验不存在的源文件路径抛出 SOURCE_NOT_FOUND
  await assert.rejects(
    host2.open({ path: path.join(root, 'nonexistent.pptx'), mode: 'file', operationId: 'op_nonexistent' }),
    { code: 'SOURCE_NOT_FOUND' }
  );
});

test('TaskHost preserves reviews, unclosed documents and task metadata across separate instances during finish', async () => {
  const root = workDir();
  const sourceBytes = await buildDeck(sampleDeck);
  const sourceFile = path.join(root, 'input.pptx');
  await fs.mkdir(root, { recursive: true });
  await fs.writeFile(sourceFile, sourceBytes);

  const taskId = randomUUID();
  const baseDir = path.join(root, 'tasks');

  // 实例 1：打开、修改、commit 产生 reviewBundle
  const host1 = new TaskHost({ taskId, baseDir });
  const opened = await host1.open({ path: sourceFile, mode: 'file', operationId: 'op_inst1_open' });
  const docId = opened.documentId;
  const inspected = await host1.inspect({ documentId: docId });
  const titleObj = inspected.objects.find(o => o.name === 'Title');

  await host1.apply({
    documentId: docId,
    expectedRevision: 0,
    operationId: 'op_inst1_apply',
    operations: [{ type: 'replace_text', targetRef: titleObj.targetRef, search: 'Original Title', replacement: 'Committed Edit', expectedMatches: 1, crossRunPolicy: 'reject' }]
  });

  const committed = await host1.commit({
    documentId: docId,
    expectedRevision: 1,
    outputPath: path.join(root, 'inst1_out.pptx'),
    operationId: 'op_inst1_commit'
  });
  const reviewId = committed.reviewId;
  assert.ok(reviewId);

  // 读取此时的 task.json 创建时间
  const taskMetaBefore = await host1.constructor.status(host1.taskDir);
  const createdTimestamp = taskMetaBefore.createdAt;

  // 模拟实例 1 进程退出（host1 销毁）
  // 实例 2 启动：以相同 taskId 初始化
  const host2 = new TaskHost({ taskId, baseDir });
  await host2.init();

  // 验证 task.json 原有创建时间未被实例 2 的 init 覆盖冲刷
  const taskMetaAfter = await host2.constructor.status(host2.taskDir);
  assert.equal(taskMetaAfter.createdAt, createdTimestamp);

  // 尝试不提供 reviewId 直接调用 finish -> 预期 UNREVIEWED_OUTPUTS（实例 2 必须能感知到磁盘上的待审阅项）
  await assert.rejects(
    host2.finish({ reviewIds: [] }),
    { code: 'UNREVIEWED_OUTPUTS' }
  );

  // 提供正确的 reviewId 调用 finish -> 预期成功，且残留的会话被正常清理
  const finishRes = await host2.finish({ reviewIds: [reviewId], preserveCheckpoints: true });
  assert.equal(finishRes.status, 'finished');
  assert.equal(finishRes.reviewedCount, 1);
});

test('Scenario 1: TaskHost enforces directory boundary safety, symlink rejection, and closed/finished lifecycle state', async () => {
  const root = workDir();
  await fs.mkdir(root, { recursive: true });
  const baseDir = path.join(root, 'tasks');

  // 1. Task ID 格式校验与路径字符防御
  assert.throws(() => new TaskHost({ taskId: '../evil-task', baseDir }), { code: 'INVALID_TASK_ID' });
  assert.throws(() => new TaskHost({ taskId: '12345', baseDir }), { code: 'INVALID_TASK_ID' });
  assert.throws(() => new TaskHost({ taskId: 'c:\\windows\\system32', baseDir }), { code: 'INVALID_TASK_ID' });

  // 2. 目录边界防御（不允许等于 baseDir，不允许越出 baseDir）
  const validId = randomUUID();
  await assert.rejects(
    assertTaskDirBoundary(baseDir, baseDir, validId),
    { code: 'TASK_DIR_OUT_OF_BOUNDS' }
  );

  // 3. Junction / Symlink 探测与拒绝（Windows 下使用 mklink /J 创建目录重解析点）
  const junctionTaskDir = path.join(baseDir, validId);
  const targetDir = path.join(root, 'target_folder');
  await fs.mkdir(targetDir, { recursive: true });
  await fs.mkdir(baseDir, { recursive: true });
  try {
    const { execSync } = await import('node:child_process');
    execSync(`cmd /c mklink /J "${junctionTaskDir}" "${targetDir}"`, { stdio: 'ignore' });
    const hostWithJunction = new TaskHost({ taskId: validId, baseDir });
    await assert.rejects(hostWithJunction.init(), { code: 'SYMLINK_NOT_ALLOWED' });
  } finally {
    try { await fs.rmdir(junctionTaskDir); } catch {}
    try { await fs.rm(targetDir, { recursive: true, force: true }); } catch {}
  }

  // 4. 关闭文稿会话持久性
  const sourceBytes = await buildDeck(sampleDeck);
  const sourceFile = path.join(root, 'source.pptx');
  await fs.writeFile(sourceFile, sourceBytes);

  const normalTaskId = randomUUID();
  const host = new TaskHost({ taskId: normalTaskId, baseDir });
  const opened = await host.open({ path: sourceFile, mode: 'file', operationId: 'op_open_close_test' });
  const docId = opened.documentId;

  // 显式关闭文稿，即使保留检查点
  const closeRes1 = await host.close({ documentId: docId, preserveCheckpoint: true });
  assert.equal(closeRes1.status, 'closed');

  // 重复关闭幂等返回 already_closed
  const closeRes2 = await host.close({ documentId: docId });
  assert.equal(closeRes2.status, 'already_closed');

  // 已关闭文稿禁止隐式重新加载
  const reloadedSession = await host.getSession(docId);
  assert.equal(reloadedSession, null);
  await assert.rejects(host.inspect({ documentId: docId }), { code: 'SESSION_NOT_FOUND' });

  // 5. 完成任务状态与回执回放
  // 打开新文稿并 commit
  const opened2 = await host.open({ path: sourceFile, mode: 'file', operationId: 'op_open_finish_test' });
  const committed2 = await host.commit({
    documentId: opened2.documentId,
    expectedRevision: 0,
    outputPath: path.join(root, 'committed2.pptx'),
    operationId: 'op_commit_finish_test'
  });

  const finishRes = await host.finish({ reviewIds: [committed2.reviewId], preserveCheckpoints: false });
  assert.equal(finishRes.status, 'finished');

  // 已完成任务：尝试进行新修改 -> 严格拒绝 TASK_ALREADY_FINISHED
  await assert.rejects(
    host.open({ path: sourceFile, mode: 'file', operationId: 'op_new_on_finished' }),
    { code: 'TASK_ALREADY_FINISHED' }
  );

  // 已完成任务：重新实例化后调用 init() 不会被隐式恢复为 active
  const hostAgain = new TaskHost({ taskId: normalTaskId, baseDir });
  await hostAgain.init();
  assert.equal(hostAgain.closed, true);
  await assert.rejects(
    hostAgain.open({ path: sourceFile, mode: 'file', operationId: 'op_new_on_finished_2' }),
    { code: 'TASK_ALREADY_FINISHED' }
  );

  // 已完成任务：按约定回放既有已完成回执，直接返回既有结果，不产生报错也不产生新写入
  const replayed = await hostAgain.open({ path: sourceFile, mode: 'file', operationId: 'op_open_finish_test' });
  assert.equal(replayed.documentId, opened2.documentId);

  // 离线状态查询如实报告 finished 状态
  const statusInfo = await TaskHost.status(hostAgain.taskDir, { operationId: 'op_commit_finish_test' });
  assert.equal(statusInfo.status, 'finished');
  assert.equal(statusInfo.operationReceipt?.status, 'completed');
});

test('Scenario 2: TaskHost enforces cross-process mutual exclusion, idempotency and revision conflict under real child processes', async () => {
  const root = workDir();
  await fs.mkdir(root, { recursive: true });
  const baseDir = path.join(root, 'tasks');
  const sourceBytes = await buildDeck(sampleDeck);
  const sourceFile = path.join(root, 'concurrency_source.pptx');
  await fs.writeFile(sourceFile, sourceBytes);

  const taskId = randomUUID();
  const host = new TaskHost({ taskId, baseDir });
  const opened = await host.open({ path: sourceFile, mode: 'file', operationId: 'op_init_open' });
  const docId = opened.documentId;
  const inspected0 = await host.inspect({ documentId: docId });
  const titleObj = inspected0.objects.find(o => o.name === 'Title');

  const workerScript = path.resolve('tests/helpers/concurrent-worker.js');

  // Helper 函数：启动一个就绪的独立 Node 子进程
  function spawnWorker() {
    const child = fork(workerScript, [], { stdio: ['inherit', 'inherit', 'inherit', 'ipc'] });
    return new Promise((resolve) => {
      child.on('message', msg => {
        if (msg.ready) resolve(child);
      });
      child.send({ action: 'init' });
    });
  }

  // --- Part A: 两个真实 Node 子进程使用相同 operationId 并发竞态 ---
  const [worker1, worker2] = await Promise.all([spawnWorker(), spawnWorker()]);

  const raceOps = [
    {
      type: 'replace_text',
      targetRef: titleObj.targetRef,
      search: 'Original Title',
      replacement: 'Race Winner',
      expectedMatches: 1,
      crossRunPolicy: 'reject'
    }
  ];

  // 同时向两子进程发送同一 operationId 的 apply 请求
  const p1 = new Promise(resolve => worker1.once('message', resolve));
  const p2 = new Promise(resolve => worker2.once('message', resolve));

  worker1.send({
    action: 'apply',
    taskId,
    baseDir,
    documentId: docId,
    expectedRevision: 0,
    operationId: 'op_concurrent_same_id',
    operations: raceOps
  });

  worker2.send({
    action: 'apply',
    taskId,
    baseDir,
    documentId: docId,
    expectedRevision: 0,
    operationId: 'op_concurrent_same_id',
    operations: raceOps
  });

  const [res1, res2] = await Promise.all([p1, p2]);
  worker1.kill();
  worker2.kill();

  // 两者都必须成功，且返回相同的 outcome 和 revision
  assert.equal(res1.success, true);
  assert.equal(res2.success, true);
  assert.equal(res1.result.outcome, 'completed');
  assert.equal(res2.result.outcome, 'completed');
  assert.equal(res1.result.revision, 1);
  assert.equal(res2.result.revision, 1);

  // 验证文稿最新 revision 仅增加了一次（依然是 1）
  const inspected1 = await host.inspect({ documentId: docId });
  assert.equal(inspected1.revision, 1);
  const titleObj1 = inspected1.objects.find(o => o.name === 'Title');
  assert.match(titleObj1.text, /Race Winner/);

  // --- Part B: 两个真实 Node 子进程使用不同 operationId、相同 expectedRevision 并发修改 ---
  const [workerA, workerB] = await Promise.all([spawnWorker(), spawnWorker()]);

  const opsA = [
    {
      type: 'replace_text',
      targetRef: titleObj1.targetRef,
      search: 'Race Winner',
      replacement: 'Winner A',
      expectedMatches: 1,
      crossRunPolicy: 'reject'
    }
  ];
  const opsB = [
    {
      type: 'replace_text',
      targetRef: titleObj1.targetRef,
      search: 'Race Winner',
      replacement: 'Winner B',
      expectedMatches: 1,
      crossRunPolicy: 'reject'
    }
  ];

  const pA = new Promise(resolve => workerA.once('message', resolve));
  const pB = new Promise(resolve => workerB.once('message', resolve));

  // 同时触发两个不同 operationId 但相同 expectedRevision: 1 的并发写请求
  workerA.send({
    action: 'apply',
    taskId,
    baseDir,
    documentId: docId,
    expectedRevision: 1,
    operationId: 'op_concurrent_diff_A',
    operations: opsA
  });

  workerB.send({
    action: 'apply',
    taskId,
    baseDir,
    documentId: docId,
    expectedRevision: 1,
    operationId: 'op_concurrent_diff_B',
    operations: opsB
  });

  const [resA, resB] = await Promise.all([pA, pB]);
  workerA.kill();
  workerB.kill();

  // 严格互斥与冲突保护：必须恰好一个成功（revision -> 2），另一个因锁内重新读取持久状态被拦截为 REVISION_MISMATCH
  const successes = [resA, resB].filter(r => r.success);
  const failures = [resA, resB].filter(r => !r.success);

  assert.equal(successes.length, 1, `Expected exactly 1 success, got ${successes.length}`);
  assert.equal(failures.length, 1, `Expected exactly 1 failure, got ${failures.length}`);
  assert.equal(successes[0].result.revision, 2);
  assert.equal(failures[0].error.code, 'REVISION_MISMATCH');

  // --- Part C: 中断残留 in_progress 回执的安全防范 ---
  // 人为写入一个来自已死亡进程的 in_progress 回执
  const receiptsDir = path.join(host.taskDir, 'receipts');
  const deadOpId = 'op_interrupted_dead_holder';
  await writeJson(path.join(receiptsDir, `${deadOpId}.json`), {
    operationId: deadOpId,
    paramsHash: hash(stableJson({ dummy: 123 })),
    status: 'in_progress',
    pid: 999999, // 假定不存在的僵尸 PID
    startedAt: new Date().toISOString()
  });

  // 尝试以相同 operationId 执行操作，预期拒绝 OPERATION_INTERRUPTED，绝不盲目自动重跑
  await assert.rejects(
    host.executeIdempotent(deadOpId, { dummy: 123 }, async () => {
      throw new Error('Should never be called');
    }),
    { code: 'OPERATION_INTERRUPTED' }
  );
});

test('TaskHost inspects and modifies slide background with revision integrity', async () => {
  const root = workDir();
  await fs.mkdir(root, { recursive: true });
  const host = new TaskHost({ taskId: randomUUID(), baseDir: root });
  await host.init();

  const deck = {
    title: 'Background Deck',
    slides: [
      { background: '112233', items: [{ type: 'text', text: 'Slide 1', left: 10, top: 10, width: 200, height: 50 }] },
      { background: '445566', items: [{ type: 'text', text: 'Slide 2', left: 10, top: 10, width: 200, height: 50 }] }
    ]
  };
  const sourceBytes = await buildDeck(deck);
  const sourceFile = path.join(root, 'bg_source.pptx');
  await fs.writeFile(sourceFile, sourceBytes);

  const opened = await host.open({
    path: sourceFile,
    mode: 'file',
    operationId: 'op_bg_open'
  });

  // inspect 过滤 kind: background
  const inspected = await host.inspect({
    documentId: opened.documentId,
    kind: 'background'
  });
  assert.equal(inspected.objects.length, 2);
  const bg1 = inspected.objects.find(o => o.slide === 1);
  const bg2 = inspected.objects.find(o => o.slide === 2);
  assert.ok(bg1 && bg2);
  assert.equal(bg1.color, '112233');
  assert.equal(bg2.color, '445566');

  // apply set_slide_background
  const applied = await host.apply({
    documentId: opened.documentId,
    expectedRevision: 0,
    operationId: 'op_bg_apply',
    operations: [
      {
        type: 'set_slide_background',
        targetRef: bg1.targetRef,
        color: 'AABBCC'
      }
    ]
  });
  assert.equal(applied.outcome, 'completed');
  assert.equal(applied.revision, 1);

  // 再次 inspect 验证
  const reinspected = await host.inspect({
    documentId: opened.documentId,
    kind: 'background'
  });
  const newBg1 = reinspected.objects.find(o => o.slide === 1);
  assert.equal(newBg1.color, 'AABBCC');

  // 过期 targetRef 拒绝
  await assert.rejects(
    host.apply({
      documentId: opened.documentId,
      expectedRevision: 1,
      operationId: 'op_bg_stale',
      operations: [
        {
          type: 'set_slide_background',
          targetRef: bg1.targetRef,
          color: 'FFFFFF'
        }
      ]
    }),
    { code: 'TARGET_STALE' }
  );

  await host.cleanupCommand();
});
