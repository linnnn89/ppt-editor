import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { contracts } from './contracts.js';
import { PptError, check } from './errors.js';
import { FileEngine } from './file-engine.js';
import { NativeHost, assertCleanupConfirmed } from './native-host.js';
import { hash, stableJson, writeJson, readJson, atomicWrite, publishNew, exists } from './storage.js';
import { createOwnedJob, restrictAccess, processIdentity } from './windows.js';
import { buildDeck } from './build.js';
import { readPackage, validatePackage } from './ooxml.js';
import packageInfo from '../package.json' with { type: 'json' };
import { readBackgroundImage, backgroundPatterns } from './background.js';
import { composeDeck } from './compose.js';
import { applyDesignTheme, designThemes, planLayout } from './design.js';
import { recordAudit, auditSummary, readOptional, summarizeNative, assessReview } from './audit-workflow.js';

const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

/**
 * 严格校验 TaskId 是否为合法的 UUID 格式，杜绝路径穿越、绝对路径或非法字符。
 */
export function validateTaskId(taskId) {
  check(typeof taskId === 'string' && UUID_REGEX.test(taskId), 'INVALID_TASK_ID', 'Task ID must be a valid UUID.');
  check(!taskId.includes('/') && !taskId.includes('\\') && !taskId.includes('..'), 'INVALID_TASK_ID', 'Task ID cannot contain path characters.');
}

/**
 * 断言任务目录严格处于 baseDir 内部，且不可与 baseDir 相同，并防范符号链接/Junction 越界。
 */
export async function assertTaskDirBoundary(baseDir, taskDir, taskId) {
  validateTaskId(taskId);
  const resolvedBase = path.resolve(baseDir);
  const resolvedTask = path.resolve(taskDir);
  const relative = path.relative(resolvedBase, resolvedTask);
  check(
    relative !== '' && !relative.startsWith('..' + path.sep) && relative !== '..' && !path.isAbsolute(relative),
    'TASK_DIR_OUT_OF_BOUNDS',
    'Task directory must be strictly inside baseDir and cannot equal baseDir.'
  );

  // 检查已有任务目录是否为符号链接或 Junction 等重解析点，若是则拒绝访问
  try {
    const stat = await fs.lstat(resolvedTask);
    check(!stat.isSymbolicLink(), 'SYMLINK_NOT_ALLOWED', 'Task directory cannot be a symbolic link or junction.');
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
}

/**
 * 跨进程内核级原子文件锁管理器（基于 fs.open 'wx' 排他创建）
 */
export class FileLock {
  /**
   * 获取排他文件锁。若已被占用且持有进程仍存活，则按重试等待或报告冲突；若持有者已死亡则判定为僵尸锁并清理恢复。
   */
  static async acquire(lockPath, { context = '', maxWaitMs = 10000, retryIntervalMs = 40 } = {}) {
    const nonce = randomUUID();
    const myPid = process.pid;
    const myIdentity = processIdentity(myPid);
    const myCreated = myIdentity?.created || '0';
    const deadline = Date.now() + maxWaitMs;

    while (true) {
      try {
        await fs.mkdir(path.dirname(lockPath), { recursive: true });
        const handle = await fs.open(lockPath, 'wx', 0o600);
        const payload = JSON.stringify({
          pid: myPid,
          created: myCreated,
          nonce,
          context,
          acquiredAt: new Date().toISOString(),
          at: Date.now()
        });
        await handle.writeFile(payload);
        await handle.sync();
        await handle.close();
        return { lockPath, nonce };
      } catch (error) {
        if (error.code !== 'EEXIST') throw error;

        // 锁文件已存在，检测持有者进程生死
        try {
          const content = await fs.readFile(lockPath, 'utf8');
          const data = JSON.parse(content);
          const holderIdentity = processIdentity(data.pid);
          const isAlive = holderIdentity && holderIdentity.created === data.created;
          if (!isAlive) {
            // 持有者进程已退出或 PID 被复用，属于崩溃残留锁，安全清理后重试
            await fs.unlink(lockPath).catch(() => {});
            continue;
          }
        } catch (readError) {
          if (readError.code === 'ENOENT') continue;
        }

        if (Date.now() >= deadline) {
          throw new PptError('LOCK_TIMEOUT', `Timed out waiting for lock: ${path.basename(lockPath)}`, { lockPath, context });
        }
        await new Promise(r => setTimeout(r, retryIntervalMs));
      }
    }
  }

  /**
   * 安全释放锁：必须核验持有者 token（nonce），避免错误删除他人锁。
   */
  static async release(lockHandle) {
    if (!lockHandle || !lockHandle.lockPath || !lockHandle.nonce) return;
    try {
      const content = await fs.readFile(lockHandle.lockPath, 'utf8');
      const data = JSON.parse(content);
      if (data.nonce === lockHandle.nonce) {
        await fs.unlink(lockHandle.lockPath).catch(() => {});
      }
    } catch {}
  }
}

/**
 * 将对象上下文编码为不透明但可校验的 targetRef 字符串。
 */
export function encodeTargetRef({ taskId, documentId, backend, revision, generation = 0, slideId, part = null, shapeId, groupPath = [], fingerprint, kind = 'shape' }) {
  const payload = [
    1, // 编码版本
    taskId,
    documentId,
    backend,
    revision,
    generation,
    slideId,
    part,
    shapeId,
    groupPath.join(','),
    fingerprint,
    kind
  ];
  return 'ref_' + Buffer.from(JSON.stringify(payload)).toString('base64url');
}

/**
 * 解码并还原 targetRef 中的上下文信息。
 */
export function decodeTargetRef(targetRef) {
  check(typeof targetRef === 'string' && targetRef.startsWith('ref_'), 'INVALID_TARGET_REF', 'Malformed target reference.');
  try {
    const raw = Buffer.from(targetRef.slice(4), 'base64url').toString('utf8');
    const payload = JSON.parse(raw);
    check(Array.isArray(payload) && payload[0] === 1, 'INVALID_TARGET_REF', 'Unsupported target reference version.');
    const [_, taskId, documentId, backend, revision, generation, slideId, part, shapeId, groupPathStr, fingerprint, rawKind] = payload;
    const groupPath = groupPathStr ? groupPathStr.split(',').map(Number) : [];
    const isNotes = Boolean(part && typeof part === 'string' && part.includes('notes'));
    const kind = rawKind || (isNotes ? 'notes' : 'shape');
    const prefix = kind === 'background' ? 'background' : kind === 'notes' ? 'notes' : 'slide';
    return {
      taskId,
      documentId,
      backend,
      revision: Number(revision),
      generation: Number(generation),
      slideId: Number(slideId),
      part,
      shapeId: Number(shapeId),
      groupPath,
      fingerprint,
      kind,
      key: `${prefix}:${slideId}:${groupPath.join('.')}:${shapeId}`
    };
  } catch (error) {
    if (error instanceof PptError) throw error;
    throw new PptError('INVALID_TARGET_REF', 'Failed to decode target reference.', { reason: error.message });
  }
}

/**
 * 文稿编辑会话，维护单份文稿的当前后端状态。
 */
class DocumentSession {
  constructor({ documentId, taskId, originalPath, mode, allowOffice, visible, directory, engine = null, revision = 0, generation = 0, checkpointPath = null, closed = false }) {
    this.documentId = documentId;
    this.taskId = taskId;
    this.originalPath = originalPath;
    this.mode = mode; // 'file' | 'native-copy'
    this.allowOffice = allowOffice;
    this.visible = visible;
    this.directory = directory;
    this.engine = engine;
    this.revision = revision;
    this.generation = generation;
    this.checkpointPath = checkpointPath;
    this.closed = closed;
  }
}

/**
 * 统一任务宿主（TaskHost）：
 * 落实目录安全、状态一致性、锁内状态刷新、原生生命周期及跨进程可靠退出。
 */
export class TaskHost {
  constructor({ taskId = randomUUID(), baseDir = 'work/tasks' } = {}) {
    validateTaskId(taskId);
    this.taskId = taskId;
    this.baseDir = path.resolve(baseDir);
    this.taskDir = path.resolve(this.baseDir, taskId);
    this.sessions = new Map();
    this.reviews = new Map();
    this.closed = false;
    this.job = null;
    this.nativeHost = null;
    this.initialized = false;
  }

  /**
   * 初始化任务目录，设置 Windows 访问控制，防范重解析点越界。
   */
  async init() {
    if (this.initialized) return;
    await assertTaskDirBoundary(this.baseDir, this.taskDir, this.taskId);
    await fs.mkdir(this.taskDir, { recursive: true });
    restrictAccess(this.taskDir, { directory: true });
    await fs.mkdir(path.join(this.taskDir, 'receipts'), { recursive: true });
    await fs.mkdir(path.join(this.taskDir, 'reviews'), { recursive: true });
    await fs.mkdir(path.join(this.taskDir, 'documents'), { recursive: true });
    await fs.mkdir(path.join(this.taskDir, 'locks'), { recursive: true });
    await fs.mkdir(path.join(this.taskDir, 'bindings'), { recursive: true });

    const taskJsonPath = path.join(this.taskDir, 'task.json');
    if (!(await exists(taskJsonPath))) {
      await writeJson(taskJsonPath, {
        taskId: this.taskId,
        status: 'active',
        createdAt: new Date().toISOString()
      });
    } else {
      const info = await readJson(taskJsonPath);
      if (info.status === 'finished') {
        this.closed = true;
      }
    }
    this.initialized = true;
  }

  /**
   * 检查任务是否已关闭/完成，已完成任务严禁重新写入。
   */
  async assertTaskActive() {
    await this.init();
    const taskJsonPath = path.join(this.taskDir, 'task.json');
    if (await exists(taskJsonPath)) {
      const info = await readJson(taskJsonPath);
      if (info.status === 'finished') {
        this.closed = true;
        throw new PptError('TASK_ALREADY_FINISHED', 'Task is already finished and cannot accept modifications.');
      }
    }
    check(!this.closed, 'TASK_CLOSED', 'Task is already closed.');
  }

  /**
   * 确保拉起受本任务管辖的原生工作进程。
   */
  async ensureNativeHost() {
    await this.init();
    await this.assertPreviousCleanup();
    if (this.nativeHost && !this.nativeHost.broken && !this.nativeHost.exited) return this.nativeHost;
    if (!this.job) this.job = createOwnedJob();
    this.nativeHost = new NativeHost(this.job, async state => {
      await writeJson(path.join(this.taskDir, 'native-state.json'), state).catch(() => {});
    });
    await this.nativeHost.start();
    return this.nativeHost;
  }

  /**
   * 从磁盘恢复所有已持久化的审阅包。
   */
  async loadReviews() {
    const reviewsDir = path.join(this.taskDir, 'reviews');
    if (!(await exists(reviewsDir))) return;
    try {
      const files = await fs.readdir(reviewsDir);
      for (const file of files) {
        if (file.endsWith('.json')) {
          const review = await readJson(path.join(reviewsDir, file)).catch(() => null);
          if (review && review.reviewId) {
            this.reviews.set(review.reviewId, review);
          }
        }
      }
    } catch {}
  }

  /**
   * 获取或从磁盘持久检查点恢复文稿会话。
   * 关键约束：已关闭的文稿禁止隐式重新加载。
   */
  async getSession(documentId) {
    const docDir = path.join(this.taskDir, 'documents', documentId);
    const metaPath = path.join(docDir, 'meta.json');
    if (!(await exists(metaPath))) return null;

    const meta = await readJson(metaPath);
    if (meta.closed) return null; // 已关闭文稿不得隐式恢复
    if (meta.unusable) {
      throw new PptError('SESSION_UNUSABLE', 'Document session entered outcome_unknown state and is unusable.');
    }

    if (this.sessions.has(documentId)) {
      const session = this.sessions.get(documentId);
      if (session.closed) return null;
      if (session.revision === meta.revision && session.generation === meta.generation) {
        return session;
      }
      if (session.mode === 'native-copy' && this.nativeHost && !this.nativeHost.exited) {
        await this.nativeHost.request('close', { documentId });
      }
      this.sessions.delete(documentId);
    }

    let engine = null;
    let checkpointPath = meta.checkpointPath;

    if (meta.mode === 'file') {
      const engineDir = path.join(docDir, 'engine');
      engine = await FileEngine.restore(engineDir);
    } else if (meta.mode === 'native-copy') {
      // 原生文稿跨进程恢复：通过持久化检查点重新在原生执行器中打开
      const nativeHost = await this.ensureNativeHost();
      const openTarget = checkpointPath || path.join(docDir, 'working.pptx');
      const opened = await nativeHost.request('open', {
        documentId,
        path: openTarget,
        directory: docDir,
        visible: meta.visible,
        revision: meta.revision,
        generation: meta.generation
      });
      checkpointPath = opened.checkpoint?.path || openTarget;
      meta.revision = opened.revision;
      meta.generation = opened.generation;
      await this.recordBinding(documentId, { originalPath: meta.originalPath, mode: meta.mode, native: opened.binding });
    }

    const session = new DocumentSession({
      documentId: meta.documentId,
      taskId: this.taskId,
      originalPath: meta.originalPath,
      mode: meta.mode,
      allowOffice: meta.allowOffice,
      visible: meta.visible,
      directory: docDir,
      engine,
      revision: meta.revision,
      generation: meta.generation,
      checkpointPath,
      closed: false
    });

    this.sessions.set(documentId, session);
    return session;
  }

  /**
   * 通用跨进程幂等执行包装器：
   * 顺序获取 operation 锁与 document 锁，并在取得 document 锁后强制重新读取持久状态。
   */
  async executeIdempotent(operationId, rawParams, executeFn, { documentId = null } = {}) {
    await this.init();
    const receiptPath = path.join(this.taskDir, 'receipts', `${operationId}.json`);
    const opLockPath = path.join(this.taskDir, 'locks', `op-${operationId}.lock`);
    const paramsHash = hash(stableJson(rawParams));

    // 1. 获取 operation 锁
    const opLock = await FileLock.acquire(opLockPath, { context: `op:${operationId}` });
    let docLock = null;

    try {
      // 检查是否已有完成回执（即使任务已完成，也允许按约定回放既有回执）
      if (await exists(receiptPath)) {
        const receipt = await readJson(receiptPath);
        const { layoutCheck, ...legacyParams } = rawParams;
        const legacyReplay = !receipt.receiptVersion && layoutCheck === true && receipt.paramsHash === hash(stableJson(legacyParams));
        check(receipt.paramsHash === paramsHash || legacyReplay, 'IDEMPOTENCY_CONFLICT', 'Operation ID was already used with different parameters.', { operationId });
        if (receipt.status === 'completed') {
          return receipt.result;
        }
        if (receipt.status === 'failed') {
          throw new PptError(receipt.error?.code || 'OPERATION_FAILED', receipt.error?.message || 'Previous operation failed.', receipt.error?.details);
        }
        if (receipt.status === 'in_progress') {
          // 该操作之前已被启动，但持有者进程中断崩溃，留下未完结的 in_progress 回执
          throw new PptError('OPERATION_INTERRUPTED', 'Previous execution of this operation was interrupted in progress; outcome unknown. Manual inspection or recovery required.', {
            operationId,
            startedAt: receipt.startedAt
          });
        }
      }

      // 未见已完成回执，需要执行写入；此时必须保证任务处于 active 状态
      await this.assertTaskActive();

      // 2. 若涉及特定文稿修改，按序获取 document 锁
      if (documentId) {
        const docDir = path.join(this.taskDir, 'documents', documentId);
        const docLockPath = path.join(docDir, 'mutation.lock');
        docLock = await FileLock.acquire(docLockPath, { context: `doc:${documentId}` });
      }

      // 记录 in_progress 回执
      await writeJson(receiptPath, {
        receiptVersion: 2,
        operationId,
        paramsHash,
        status: 'in_progress',
        pid: process.pid,
        startedAt: new Date().toISOString()
      });

      const result = await executeFn();

      // 记录 completed 回执
      await writeJson(receiptPath, {
        receiptVersion: 2,
        operationId,
        paramsHash,
        status: 'completed',
        result,
        completedAt: new Date().toISOString()
      });

      return result;
    } catch (error) {
      if (error.code !== 'IDEMPOTENCY_CONFLICT' && error.code !== 'OPERATION_INTERRUPTED') {
        await writeJson(receiptPath, {
          receiptVersion: 2,
          operationId,
          paramsHash,
          status: 'failed',
          error: { code: error.code || 'UNKNOWN_ERROR', message: error.message, details: error.details },
          failedAt: new Date().toISOString()
        }).catch(() => {});
      }
      throw error;
    } finally {
      // 逆序释放锁
      if (docLock) await FileLock.release(docLock);
      if (opLock) await FileLock.release(opLock);
    }
  }

  /**
   * ppt_open: 打开文稿，建立安全任务会话。
   */
  async open(params) {
    const valid = contracts.ppt_open.parse(params);
    return this.executeIdempotent(valid.operationId, valid, async () => {
      await this.assertTaskActive();
      const originalPath = path.resolve(valid.path);
      check(await exists(originalPath), 'SOURCE_NOT_FOUND', 'Source presentation file does not exist.', { path: originalPath });
      const bytes = await fs.readFile(originalPath);
      const sourceHash = hash(bytes);

      if (valid.mode === 'native-copy') {
        check(valid.allowOffice, 'OFFICE_NOT_ALLOWED', 'Native mode requires explicit allowOffice permission.');
      }

      const documentId = randomUUID();
      const docDir = path.join(this.taskDir, 'documents', documentId);
      await fs.mkdir(docDir, { recursive: true });
      restrictAccess(docDir, { directory: true });

      let engine = null;
      let stateHash = '';
      let revision = 0;
      let generation = 0;
      let checkpointPath = null;
      let nativeBinding = null;

      if (valid.mode === 'file') {
        engine = await FileEngine.create(path.join(docDir, 'engine'), bytes);
        stateHash = hash(stableJson(engine.inspect().objects.map(o => [o.key, o.fingerprint])));
      } else {
        const nativeHost = await this.ensureNativeHost();
        const workingPath = path.join(docDir, 'working.pptx');
        await fs.writeFile(workingPath, bytes);
        const opened = await nativeHost.request('open', {
          documentId,
          path: workingPath,
          directory: docDir,
          visible: valid.visible
        });
        stateHash = opened.snapshot?.stateHash || '';
        // 严格同步工作进程真实返回的 generation 与 revision
        generation = opened.generation;
        revision = opened.revision;
        checkpointPath = opened.checkpoint?.path || workingPath;
        nativeBinding = opened.binding;
      }

      const session = new DocumentSession({
        documentId,
        taskId: this.taskId,
        originalPath,
        mode: valid.mode,
        allowOffice: valid.allowOffice,
        visible: valid.visible,
        directory: docDir,
        engine,
        revision,
        generation,
        checkpointPath,
        closed: false
      });

      this.sessions.set(documentId, session);
      await this.recordBinding(documentId, { originalPath, sourceHash, mode: valid.mode, native: nativeBinding });

      await writeJson(path.join(docDir, 'meta.json'), {
        documentId,
        taskId: this.taskId,
        originalPath,
        sourceHash,
        mode: valid.mode,
        allowOffice: valid.allowOffice,
        visible: valid.visible,
        revision,
        generation,
        checkpointPath,
        closed: false,
        createdAt: new Date().toISOString()
      });

      return {
        documentId,
        mode: valid.mode,
        revision,
        generation,
        sourceHash,
        stateHash
      };
    });
  }

  /**
   * ppt_build: 从声明式结构新建文稿。
   */
  async build(params) {
    const valid = contracts.ppt_build.parse({ ...params, deck: applyDesignTheme(params.deck) });
    return this.executeIdempotent(valid.operationId, valid, async () => {
      await this.assertTaskActive();
      const bytes = await buildDeck(valid.deck);
      const tempOriginal = path.join(this.taskDir, `built-${randomUUID()}.pptx`);
      await fs.writeFile(tempOriginal, bytes);

      const buildOpenOpId = ('bo_' + hash(valid.operationId)).slice(0, 40);
      const opened = await this.open({
        path: tempOriginal,
        mode: valid.mode,
        allowOffice: valid.allowOffice,
        visible: valid.visible,
        operationId: buildOpenOpId
      });

      return {
        ...opened,
        operationId: valid.operationId,
        deckTitle: valid.deck.title,
        ...(valid.layoutCheck ? { layoutAudit: await this.batchAudit(opened.documentId, opened.revision) } : {})
      };
    });
  }

  async relayout(params) {
    const valid = contracts.ppt_relayout.parse(params);
    return this.executeIdempotent(valid.operationId, valid, async () => {
      // Reinspect durable state; apply performs the final revision check under its lock.
      if (this.sessions.get(valid.documentId)?.mode === 'file') this.sessions.delete(valid.documentId);
      const snapshot = await this.inspect({ documentId: valid.documentId, slide: valid.slide, limit: 300 });
      check(snapshot.revision === valid.expectedRevision, 'REVISION_MISMATCH', 'Layout revision is stale.');
      check(snapshot.totalObjects <= 300, 'LAYOUT_TOO_MANY_OBJECTS', 'Layout inspection is incomplete for this slide.');
      const plan = planLayout(snapshot, valid);
      if (valid.dryRun) return { outcome: 'layout_planned', revision: snapshot.revision, ...plan };
      const applied = await this.apply({ documentId: valid.documentId, expectedRevision: valid.expectedRevision,
        operationId: ('layout_' + hash(valid.operationId)).slice(0, 60), operations: plan.operations });
      return { ...applied, layout: plan };
    });
  }

  async recordBinding(documentId, data) {
    const destination = path.join(this.taskDir, 'bindings', `${documentId}.json`);
    const previous = await readJson(destination).catch(error => { if (error.code === 'ENOENT') return {}; throw error; });
    await writeJson(destination, { ...previous, ...data, taskId: this.taskId, documentId,
      recordedAt: new Date().toISOString(), note: 'Historical binding; live native requests recheck process identity and document path.' });
  }

  async compose(params) {
    const valid = contracts.ppt_compose.parse({ ...params, ...(params.contentDeck ? { contentDeck: applyDesignTheme(params.contentDeck) } : {}) });
    return this.executeIdempotent(valid.operationId, valid, async () => {
      await this.assertTaskActive();
      const templateBytes = await fs.readFile(path.resolve(valid.templatePath));
      const contentBytes = valid.sourcePath ? await fs.readFile(path.resolve(valid.sourcePath)) : await buildDeck(valid.contentDeck);
      const { bytes, report } = await composeDeck(templateBytes, contentBytes, valid);
      const generated = path.join(this.taskDir, `composed-${randomUUID()}.pptx`);
      await publishNew(generated, bytes);
      const opened = await this.open({ path: generated, mode: 'file', allowOffice: valid.allowOffice, operationId: ('co_' + hash(valid.operationId)).slice(0, 40) });
      return { ...opened, composition: report, sourceHashes: { template: hash(templateBytes), content: hash(contentBytes) },
        ...(valid.layoutCheck ? { layoutAudit: await this.batchAudit(opened.documentId, opened.revision) } : {}) };
    });
  }

  /**
   * ppt_inspect: 查询文稿中的对象树并派发安全绑定的 targetRef。
   */
  async inspect(params) {
    const valid = contracts.ppt_inspect.parse(params);
    await this.assertTaskActive();
    const session = await this.getSession(valid.documentId);
    check(session && !session.closed, 'SESSION_NOT_FOUND', 'Document session not found or already closed.');

    let rawIndex;
    if (session.mode === 'file') {
      rawIndex = session.engine.inspect({ slides: valid.slide ? [valid.slide] : undefined });
    } else {
      const nativeHost = await this.ensureNativeHost();
      const res = await nativeHost.request('inspect', { documentId: session.documentId });
      rawIndex = res.snapshot;
      session.revision = res.revision;
      session.generation = res.generation;
    }

    let filteredObjects = rawIndex.objects;
    if (valid.slide !== undefined) filteredObjects = filteredObjects.filter(o => o.slide === valid.slide);
    if (valid.kind !== undefined) filteredObjects = filteredObjects.filter(o => o.kind === valid.kind);
    if (valid.text !== undefined) filteredObjects = filteredObjects.filter(o => o.text && o.text.includes(valid.text));

    const totalObjects = filteredObjects.length;
    const paginated = filteredObjects.slice(valid.offset, valid.offset + valid.limit);

    const objects = paginated.map(o => ({
      ...(valid.detail === 'summary' ? { key: o.key, slide: o.slide, name: o.name, kind: o.kind,
        text: o.text?.slice(0, 240), textTruncated: (o.text?.length || 0) > 240, geometry: o.geometry, capabilities: o.capabilities } : o),
      targetRef: encodeTargetRef({
        taskId: this.taskId,
        documentId: session.documentId,
        backend: session.mode,
        revision: session.revision,
        generation: session.generation,
        slideId: o.slideId,
        part: o.part,
        shapeId: o.shapeId,
        groupPath: o.groupPath || [],
        fingerprint: o.fingerprint,
        kind: o.kind
      })
    }));

    return {
      documentId: session.documentId,
      revision: session.revision,
      generation: session.generation,
      totalObjects,
      width: rawIndex.width, height: rawIndex.height,
      offset: valid.offset,
      limit: valid.limit,
      slides: rawIndex.slides,
      objects
    };
  }

  /**
   * ppt_apply: 批量执行修改操作。
   * 关键约束：取得 document 锁后强制重新读取磁盘持久状态，严格核对 expectedRevision。
   */
  async apply(params) {
    const valid = contracts.ppt_apply.parse(params);
    return this.executeIdempotent(valid.operationId, valid, async () => {
      await this.assertTaskActive();
      // 在持锁状态下重新加载最新会话与磁盘状态
      if (this.sessions.get(valid.documentId)?.mode === 'file') this.sessions.delete(valid.documentId);
      const session = await this.getSession(valid.documentId);
      check(session && !session.closed, 'SESSION_NOT_FOUND', 'Document session not found or already closed.');

      // 锁内检查最新 revision
      check(
        session.revision === valid.expectedRevision,
        'REVISION_MISMATCH',
        'Document revision mismatch. Refresh inspection before applying modifications.',
        { expected: valid.expectedRevision, actual: session.revision }
      );

      const normalizedOperations = valid.operations.map(op => {
        const ref = decodeTargetRef(op.targetRef);
        check(ref.taskId === this.taskId, 'CROSS_TASK_REFERENCE', 'Target does not belong to this task.');
        check(ref.documentId === session.documentId, 'CROSS_DOCUMENT_REFERENCE', 'Target does not belong to this document.');
        check(ref.backend === session.mode, 'BACKEND_MISMATCH', 'Target backend does not match current session mode.');
        check(
          ref.revision === session.revision,
          'TARGET_STALE',
          'Target reference is stale (was obtained from a different revision). Re-inspect to obtain fresh target references.',
          { targetRevision: ref.revision, currentRevision: session.revision }
        );
        check(
          ref.generation === session.generation,
          'TARGET_STALE',
          'Target reference generation mismatch.',
          { targetGeneration: ref.generation, currentGeneration: session.generation }
        );

        const { targetRef, ...rest } = op;
        return {
          ...rest,
          target: {
            key: ref.key,
            fingerprint: ref.fingerprint,
            slideId: ref.slideId,
            shapeId: ref.shapeId,
            groupPath: ref.groupPath,
            kind: ref.kind || (ref.key.startsWith('notes') ? 'notes' : ref.key.startsWith('background') ? 'background' : 'shape')
          }
        };
      });

      let outcome = '';
      if (session.mode === 'native-copy') for (const op of normalizedOperations) {
        if (op.type !== 'set_slide_background' || !['image', 'texture'].includes(op.fill?.type)) continue;
        const image = await readBackgroundImage(op.fill.path), assetDir = path.join(session.directory, 'assets');
        await fs.mkdir(assetDir, { recursive: true });
        const assetPath = path.join(assetDir, `${image.sha256}.${image.extension}`);
        if (!await exists(assetPath)) await fs.writeFile(assetPath, image.bytes, { flag: 'wx', mode: 0o600 });
        op.fill = { ...op.fill, path: assetPath }; op.imageHash = image.sha256;
      }
      let changes = [];
      let stateHash = '';
      let postSnapshot;

      if (session.mode === 'file') {
        const result = await session.engine.apply(normalizedOperations, { dryRun: valid.dryRun });
        outcome = result.outcome;
        changes = result.changes;
        stateHash = result.stateHash || '';
        if (!valid.dryRun) {
          session.revision = result.revision;
        }
      } else {
        const nativeHost = await this.ensureNativeHost();
        const result = await nativeHost.request('apply', {
          documentId: session.documentId,
          generation: session.generation,
          operations: normalizedOperations,
          dryRun: valid.dryRun
        });

        outcome = result.outcome;
        changes = result.changes || [];
        stateHash = result.snapshot?.stateHash || '';
        postSnapshot = result.snapshot;

        // 真实处理原生执行结果
        if (valid.dryRun && result.outcome === 'preflight_passed') {
          return { outcome: result.outcome, documentId: session.documentId, revision: result.revision,
            generation: result.generation, changes: [], stateHash };
        } else if (result.outcome === 'completed') {
          if (!valid.dryRun) {
            session.revision = result.revision;
            session.generation = result.generation;
            session.checkpointPath = result.checkpoint?.path || session.checkpointPath;
          }
        } else if (result.outcome === 'restored_from_checkpoint') {
          // 从检查点恢复了，同步新的 generation，但 revision 未增加
          session.generation = result.generation;
          session.revision = result.revision;
          session.checkpointPath = result.checkpoint?.path || session.checkpointPath;
          const metaPath = path.join(session.directory, 'meta.json');
          await writeJson(metaPath, { ...await readJson(metaPath), revision: session.revision,
            generation: session.generation, checkpointPath: session.checkpointPath, updatedAt: new Date().toISOString() });
          throw new PptError('NATIVE_APPLY_FAILED', 'Native operations failed; restored from last durable checkpoint.', {
            outcome: result.outcome,
            completedBeforeFailure: result.completedBeforeFailure,
            error: result.error
          });
        } else {
          // outcome_unknown 状态，禁止继续操作
          session.unusable = true;
          await writeJson(path.join(session.directory, 'meta.json'), {
            ...await readJson(path.join(session.directory, 'meta.json')),
            documentId: session.documentId,
            taskId: this.taskId,
            originalPath: session.originalPath,
            mode: session.mode,
            unusable: true,
            updatedAt: new Date().toISOString()
          }).catch(() => {});
          throw new PptError('NATIVE_OUTCOME_UNKNOWN', 'Native execution outcome is unknown; session is unusable.', {
            outcome: result.outcome,
            error: result.error,
            recoveryError: result.recoveryError
          });
        }
      }

      if (!valid.dryRun) {
        await this.loadReviews();
        for (const review of this.reviews.values()) {
          if (review.documentId === session.documentId) {
            review.stale = true;
            await writeJson(path.join(this.taskDir, 'reviews', `${review.reviewId}.json`), review);
          }
        }
      }

      await writeJson(path.join(session.directory, 'meta.json'), {
        documentId: session.documentId,
        taskId: this.taskId,
        originalPath: session.originalPath,
        mode: session.mode,
        allowOffice: session.allowOffice,
        visible: session.visible,
        revision: session.revision,
        generation: session.generation,
        checkpointPath: session.checkpointPath,
        closed: false,
        updatedAt: new Date().toISOString()
      });

      let layoutAudit;
      if (!valid.dryRun && valid.layoutCheck) {
        const index = postSnapshot || session.engine.inspect({ slides: [] });
        const ids = new Set(normalizedOperations.map(op => op.target.slideId));
        const slides = index.slides.filter(s => ids.has(s.slideId)).map(s => s.slide);
        try { layoutAudit = await this.recordLayoutAudit(session, postSnapshot || session.engine.inspect({ slides }), { slides, detail: 'summary' }); }
        catch (error) { layoutAudit = { status: 'failed', error: error.code || 'LAYOUT_CHECK_FAILED', message: error.message, nextAction: 'run_layout_check' }; }
      }
      return {
        outcome,
        documentId: session.documentId,
        revision: session.revision,
        generation: session.generation,
        changes,
        stateHash,
        ...(layoutAudit ? { layoutAudit } : {})
      };
    }, { documentId: valid.documentId });
  }

  /**
   * ppt_validate: 执行真实包验证与原生状态核验。
   */
  async validate(params) {
    const valid = contracts.ppt_validate.parse(params);
    await this.assertTaskActive();
    const docDir = path.join(this.taskDir, 'documents', valid.documentId);
    const lock = await FileLock.acquire(path.join(docDir, 'mutation.lock'), { context: `validate:${valid.documentId}` });
    try {
      await this.assertTaskActive();
      if (valid.checks === 'full' && this.sessions.get(valid.documentId)?.mode === 'file') this.sessions.delete(valid.documentId);
      const session = await this.getSession(valid.documentId);
      check(session && !session.closed, 'SESSION_NOT_FOUND', 'Document session not found or already closed.');

      if (session.mode === 'file') {
        if (valid.expectedRevision !== undefined) check(valid.expectedRevision === session.revision, 'REVISION_MISMATCH', 'Refresh layout references before declaring design overlaps.');
        const report = valid.checks === 'layout' ? { structural: 'not_run', nativeReadback: 'not_run', visualReview: 'not_run' } : session.engine.validate();
        if (valid.layoutCheck && !valid.nativeReadback) report.layoutAudit = await this.recordLayoutAudit(session, session.engine.inspect({ slides: valid.slides }), valid);
        if (valid.nativeReadback) {
          check(valid.allowOffice && session.allowOffice, 'OFFICE_NOT_ALLOWED', 'Native readback requires explicit allowOffice permission.');
          const nativeHost = await this.ensureNativeHost();
          const bytes = await session.engine.bytes();
          const parts = await readPackage(bytes);
          check(parts.size === session.engine.parts.size && [...session.engine.parts].every(([name, data]) => parts.get(name)?.equals(data)),
            'OUTPUT_MISMATCH', 'Generated readback package does not match the edited document.');
          const candidate = path.join(session.directory, `readback-${randomUUID()}.pptx`);
          await publishNew(candidate, bytes);
          // On failure leave the candidate available for native cleanup and diagnosis.
          const readbackResult = await nativeHost.request('validateFile', { path: candidate, sha256: hash(bytes) });
          check(readbackResult.nativeReadback === 'passed' && readbackResult.sha256 === hash(bytes),
            'NATIVE_READBACK_FAILED', 'Native readback did not confirm the generated candidate.');
          report.nativeReadback = 'passed';
          report.nativeValidation = { ...readbackResult, revision: session.revision, generation: session.generation, basis: 'generated-output-bytes' };
          if (valid.layoutCheck) report.layoutAudit = await this.recordLayoutAudit(session, readbackResult.snapshot, valid);
          if (valid.detail === 'summary') report.nativeValidation = summarizeNative(report.nativeValidation);
          await fs.unlink(candidate);
        }
        return report;
      } else {
        const nativeHost = await this.ensureNativeHost();
        const report = await nativeHost.request('validate', { documentId: session.documentId, layoutCheck: valid.layoutCheck,
          checks: valid.checks, expectedRevision: valid.expectedRevision });
        session.revision = report.revision; session.generation = report.generation;
        if (report.snapshot) { report.layoutAudit = await this.recordLayoutAudit(session, report.snapshot, valid); delete report.snapshot; }
        return report;
      }
    } finally { await FileLock.release(lock); }
  }

  async recordLayoutAudit(session, snapshot, options) {
    return recordAudit(this.taskDir, session, snapshot, options);
  }

  async batchAudit(documentId, expectedRevision) {
    // Generation/editing already succeeded. Report a failed follow-up check
    // explicitly without presenting the completed mutation as safe to retry.
    try { return (await this.validate({ documentId, expectedRevision, layoutCheck: true, checks: 'layout', detail: 'summary' })).layoutAudit; }
    catch (error) { return { status: 'failed', error: error.code || 'LAYOUT_CHECK_FAILED', message: error.message, nextAction: 'run_layout_check' }; }
  }

  /**
   * ppt_commit: 另存发布到新路径并生成真实验证的 ReviewBundle。
   */
  async commit(params) {
    const valid = contracts.ppt_commit.parse(params);
    return this.executeIdempotent(valid.operationId, valid, async () => {
      await this.assertTaskActive();
      if (this.sessions.get(valid.documentId)?.mode === 'file') this.sessions.delete(valid.documentId);
      const session = await this.getSession(valid.documentId);
      check(session && !session.closed, 'SESSION_NOT_FOUND', 'Document session not found or already closed.');

      check(
        session.revision === valid.expectedRevision,
        'REVISION_MISMATCH',
        'Document revision mismatch at commit.',
        { expected: valid.expectedRevision, actual: session.revision }
      );

      const targetPath = path.resolve(valid.outputPath);
      check(
        targetPath.toLowerCase() !== session.originalPath.toLowerCase(),
        'OVERWRITE_SOURCE_PROHIBITED',
        'Cannot overwrite original source presentation. Please choose a new destination path.'
      );

      let published;
      let validationReport;

      if (session.mode === 'file') {
        const bytes = await session.engine.bytes();
        const outputParts = await readPackage(bytes);
        check(outputParts.size === session.engine.parts.size &&
          [...session.engine.parts].every(([name, data]) => outputParts.get(name)?.equals(data)),
        'OUTPUT_MISMATCH', 'Generated package does not match the edited document.');
        validationReport = { ...session.engine.validate(), structural: validatePackage(outputParts), basis: 'generated-output-bytes' };
        published = await publishNew(targetPath, bytes);
      } else {
        const nativeHost = await this.ensureNativeHost();
        const commitRes = await nativeHost.request('commit', {
          documentId: session.documentId,
          generation: session.generation,
          outputPath: targetPath
        });
        published = {
          path: commitRes.outputPath,
          sha256: commitRes.sha256,
          bytes: commitRes.bytes
        };
        validationReport = commitRes.validation;
      }

      const reviewId = randomUUID();
      const reviewBundle = {
        reviewId,
        documentId: session.documentId,
        revision: session.revision,
        mode: session.mode,
        outputPath: published.path,
        generation: session.generation,
        outputHash: published.sha256,
        outputBytes: published.bytes,
        validation: validationReport,
        visualReviewRequired: true,
        visualReviewed: false,
        stale: false,
        createdAt: new Date().toISOString()
      };

      this.reviews.set(reviewId, reviewBundle);
      try {
        await writeJson(path.join(this.taskDir, 'reviews', `${reviewId}.json`), reviewBundle);
      } catch (err) {
        err.details = { ...(err.details || {}), published: true, outputPath: published.path };
        throw err;
      }

      return {
        reviewId,
        documentId: session.documentId,
        revision: session.revision,
        outputPath: published.path,
        outputHash: published.sha256,
        outputBytes: published.bytes,
        reviewBundle
      };
    }, { documentId: valid.documentId });
  }

  /**
   * ppt_close: 明确关闭文稿会话，持久化 closed 标记防止后续被错误恢复。
   */
  async close(params) {
    const valid = contracts.ppt_close.parse(params);
    await this.assertTaskActive();
    await this.assertPreviousCleanup();
    const docDir = path.join(this.taskDir, 'documents', valid.documentId);
    if (!(await exists(docDir))) return { documentId: valid.documentId, status: 'already_closed' };

    const docLockPath = path.join(docDir, 'mutation.lock');
    const docLock = await FileLock.acquire(docLockPath, { context: `close:${valid.documentId}` });

    try {
      if (this.sessions.get(valid.documentId)?.mode === 'file') this.sessions.delete(valid.documentId);
      const session = await this.getSession(valid.documentId);
      if (!session || session.closed) return { documentId: valid.documentId, status: 'already_closed' };

      if (session.mode === 'native-copy') {
        check(this.nativeHost && !this.nativeHost.exited && !this.nativeHost.broken,
          'CLEANUP_UNCONFIRMED', 'Native document close cannot be confirmed; checkpoint retained.');
        const result = await this.nativeHost.request('close', {
          documentId: valid.documentId,
          preserveCheckpoint: valid.preserveCheckpoint
        });
        check(result?.closed === true, 'CLEANUP_UNCONFIRMED', 'Native document close was not confirmed; checkpoint retained.');
      }

      await writeJson(path.join(session.directory, 'meta.json'), {
        ...await readJson(path.join(session.directory, 'meta.json')),
        closed: true,
        closedAt: new Date().toISOString()
      });
      session.closed = true;
      await this.recordBinding(valid.documentId, { closed: true });

      if (!valid.preserveCheckpoint) {
        // 清理中间大体积检查点
        if (session.mode === 'file') {
          await fs.rm(path.join(session.directory, 'engine', 'parts'), { recursive: true, force: true }).catch(() => {});
        } else {
          const files = await fs.readdir(session.directory).catch(() => []);
          for (const f of files) {
            if (f.startsWith('native-') && f.endsWith('.pptx')) {
              await fs.unlink(path.join(session.directory, f)).catch(() => {});
            }
          }
        }
      }

      this.sessions.delete(valid.documentId);
      return { documentId: valid.documentId, status: 'closed' };
    } finally {
      await FileLock.release(docLock);
    }
  }

  /**
   * ppt_finish: 任务验收完成。清理大体积临时文件，保留最小状态与操作回执。
   */
  async finish(params) {
    const valid = contracts.ppt_finish.parse(params);
    await this.init();
    await this.loadReviews();

    const taskLockPath = path.join(this.taskDir, 'locks', 'task.lock');
    const taskLock = await FileLock.acquire(taskLockPath, { context: 'finish' });

    try {
      await this.assertPreviousCleanup();
      await this.loadReviews();
      check(new Set(valid.reviewIds).size === valid.reviewIds.length && valid.reviewIds.every(id => this.reviews.has(id)),
        'INVALID_REVIEW_ID', 'Review IDs must be unique and belong to this task.');
      // Acknowledgment must still refer to the bytes whose preview was reviewed.
      for (const review of this.reviews.values()) {
        if (!review.stale && valid.reviewIds.includes(review.reviewId)) {
          await this.assertReviewOutput(review);
          for (const image of review.preview?.images || []) {
            check(await exists(image.path) && hash(await fs.readFile(image.path)) === image.sha256,
              'PREVIEW_CHANGED', 'Preview image changed or is missing. Render and review the output again.');
          }
        }
      }
      const unreviewed = Array.from(this.reviews.values())
        .filter(r => !r.stale && !valid.reviewIds.includes(r.reviewId))
        .map(r => r.reviewId);
      check(
        unreviewed.length === 0,
        'UNREVIEWED_OUTPUTS',
        'There are unreviewed committed outputs in this task. Provide all active reviewIds to finish.',
        { unreviewed }
      );

      const collectAcceptance = async () => {
        const reviews = await Promise.all([...this.reviews.values()].filter(r => !r.stale && valid.reviewIds.includes(r.reviewId))
          .map(async review => ({ reviewId: review.reviewId, ...await assessReview(this.taskDir, review) })));
        return { status: !reviews.length ? 'no_outputs' : reviews.every(r => r.status === 'accepted') ? 'accepted' : 'incomplete', reviews };
      };
      let acceptance = await collectAcceptance();
      check(!valid.requireAccepted || acceptance.status === 'accepted', 'ACCEPTANCE_INCOMPLETE',
        'Complete the final layout check and review previews for every output page before accepting. Resources are still available for repair.', { acceptance });

      const docsDir = path.join(this.taskDir, 'documents');
      const closeErrors = [];
      if (await exists(docsDir)) {
        const docEntries = await fs.readdir(docsDir).catch(() => []);
        for (const docId of docEntries) {
          // Retain every checkpoint until both document close and host shutdown are confirmed.
          try { await this.close({ documentId: docId, preserveCheckpoint: true }); }
          catch (error) { closeErrors.push({ documentId: docId, code: error.code || 'DOCUMENT_CLOSE_FAILED', message: error.message }); }
        }
      }

      if (closeErrors.length) {
        const info = await readJson(path.join(this.taskDir, 'task.json'));
        await writeJson(path.join(this.taskDir, 'task.json'), { ...info, closeErrors });
      }
      const shutdownReport = await this.cleanupCommand();
      check(closeErrors.length === 0, 'CLEANUP_UNCONFIRMED', 'Some documents could not be closed; recovery data retained.', { closeErrors, shutdownReport });

      // Shutdown can take seconds; do not acknowledge bytes changed during it.
      await this.loadReviews();
      check([...this.reviews.values()].every(r => r.stale || valid.reviewIds.includes(r.reviewId)),
        'UNREVIEWED_OUTPUTS', 'Committed outputs changed during cleanup. Review them before finishing.');
      for (const review of this.reviews.values()) if (!review.stale && valid.reviewIds.includes(review.reviewId)) {
        await this.assertReviewOutput(review);
        for (const image of review.preview?.images || []) check(await exists(image.path) && hash(await fs.readFile(image.path)) === image.sha256,
          'PREVIEW_CHANGED', 'Preview changed during cleanup. Render and review again.');
      }
      acceptance = await collectAcceptance();
      check(!valid.requireAccepted || acceptance.status === 'accepted', 'ACCEPTANCE_INCOMPLETE', 'Output acceptance changed during cleanup; retained checkpoints require review.', { acceptance });

      for (const review of this.reviews.values()) {
        if (!review.stale && valid.reviewIds.includes(review.reviewId)) {
          // This records the caller's confirmation; rendering alone never approves.
          review.acceptance = await assessReview(this.taskDir, review);
          review.callerConfirmed = true;
          review.visualReviewed = review.acceptance.visualEvidenceComplete;
          review.visualReviewMethod = review.visualReviewed ? 'caller-confirmation' : 'not_verified';
          review.reviewedAt = new Date().toISOString();
          await writeJson(path.join(this.taskDir, 'reviews', `${review.reviewId}.json`), review);
        }
      }

      let checkpointCleanupError = null;
      if (!valid.preserveCheckpoints) {
        // 清理大体积中间检查点，保留 task.json、receipts、reviews 最小回执
        try {
          await fs.rm(docsDir, { recursive: true, force: true });
        } catch (err) {
          checkpointCleanupError = err.message;
        }
      }

      this.closed = true;
      await writeJson(path.join(this.taskDir, 'task.json'), {
        taskId: this.taskId,
        status: 'finished',
        acceptance,
        finishedAt: new Date().toISOString(),
        shutdownReport,
        checkpointCleanupError
      });

      return {
        status: 'finished',
        taskId: this.taskId,
        reviewedCount: valid.reviewIds.length,
        acceptance,
        shutdownReport,
        checkpointCleanupError
      };
    } finally {
      await FileLock.release(taskLock);
    }
  }

  /**
   * ppt_status: 实例状态查询快捷方法（包装静态 status）
   */
  async status(params = {}) {
    await this.init();
    const valid = contracts.ppt_status.parse(params);
    const id = valid.taskId || this.taskId;
    const directory = path.resolve(this.baseDir, id);
    await assertTaskDirBoundary(this.baseDir, directory, id);
    return TaskHost.status(directory, valid);
  }

  /**
   * ppt_diagnose: 诊断环境状态、任务活跃度与 Office 进程信息
   */
  async diagnose(params = {}) {
    await this.init();
    contracts.ppt_diagnose.parse(params);
    const { listProcesses } = await import('./windows.js');
    const officeProcesses = listProcesses('POWERPNT.EXE');
    const diskVersion = (await readJson(new URL('../package.json', import.meta.url))).version;
    return {
      status: 'ok',
      version: packageInfo.version,
      diskVersion,
      restartRequired: diskVersion !== packageInfo.version,
      nodeVersion: process.version,
      hostProcess: processIdentity(process.pid),
      capabilities: { shapeTransforms: ['flipV', 'flipH', 'rotation'], textStyles: ['fontFace', 'fontSize', 'color', 'bold', 'italic', 'underline'],
        charts: ['bar','line','area','pie','doughnut','radar','combo'], smartArt: ['native-conversion','native-node-text','template-preservation'],
        visualLayouts: ['two-column','grid','stack'], designThemes: Object.keys(designThemes),
        pageLayoutAudit: { codeOnly: true, overlapReview: 'immediate-unless-declared-design', finalWholeDeckCheck: true, textFit: false, nativeTextBounds: 'ordinary-unrotated-text' },
        batchPageReports: true, pageDependencyFreshness: true, compactInspection: true,
        inheritedPlaceholderGeometry: 'complete-transform', renderLayoutReadback: true,
        draftCheckpointPreview: true, cumulativePreviewEvidence: true, explicitAcceptance: true,
        backgroundFills: ['solid', 'gradient', 'image', 'texture', 'pattern', 'inherit'], backgroundPatterns: Object.keys(backgroundPatterns),
      templateComposition: true,
        taskDocumentProcessBinding: true,
        committedOutputPreview: true, durablePreviewRecords: true },
      taskId: this.taskId,
      taskDir: this.taskDir,
      closed: this.closed,
      activeSessions: this.sessions.size,
      officeProcesses: officeProcesses.map(p => ({ pid: p.pid }))
    };
  }

  /**
   * ppt_render: 原生副本模式下的单页/多页幻灯片高清导出渲染
   */
  async render(params) {
    const valid = contracts.ppt_render.parse(params);
    await this.assertTaskActive();
    if (valid.reviewId) return this.renderCommitted(valid);
    const session = await this.getSession(valid.documentId);
    check(session && !session.closed, 'SESSION_NOT_FOUND', 'Document session not found or already closed.');
    if (session.mode === 'file') return this.renderFileDraft(valid);
    if (valid.expectedRevision !== undefined) check(valid.expectedRevision === session.revision, 'REVISION_MISMATCH', 'Inspect the current revision before previewing.');
    check(session.mode === 'native-copy', 'RENDER_REQUIRES_NATIVE', 'Slide rendering is only available in native-copy mode.');
    check(valid.allowOffice && session.allowOffice, 'OFFICE_NOT_ALLOWED', 'Rendering requires explicit allowOffice permission.');
    const nativeHost = await this.ensureNativeHost();
    const renderDir = path.join(session.directory, 'renders', `rev-${session.revision}`);
    return nativeHost.request('render', {
      documentId: session.documentId,
      directory: renderDir,
      slides: valid.slides,
      width: valid.width
    });
  }

  async renderFileDraft(valid) {
    check(valid.expectedRevision !== undefined, 'REVISION_REQUIRED', 'Draft previews require expectedRevision.');
    const taskLock = await FileLock.acquire(path.join(this.taskDir, 'locks', 'task.lock'), { context: 'draft-preview' });
    let docLock;
    try {
      docLock = await FileLock.acquire(path.join(this.taskDir, 'documents', valid.documentId, 'mutation.lock'), { context: 'draft-preview' });
      await this.assertTaskActive();
      this.sessions.delete(valid.documentId);
      const session = await this.getSession(valid.documentId);
      check(session && session.mode === 'file', 'SESSION_NOT_FOUND', 'File session is unavailable.');
      check(session.revision === valid.expectedRevision, 'REVISION_MISMATCH', 'Draft preview revision is stale.');
      check(session.allowOffice && valid.allowOffice, 'OFFICE_NOT_ALLOWED', 'Draft preview requires Office permission at open and render.');
      const directory = path.join(this.taskDir, 'draft-previews', valid.documentId, randomUUID());
      const bytes = await session.engine.bytes(), candidate = path.join(directory, 'candidate.pptx');
      await publishNew(candidate, bytes);
      const host = await this.ensureNativeHost();
      const native = await host.request('validateFile', { path: candidate, sha256: hash(bytes), render: { slides: valid.slides, width: valid.width, directory } });
      check(native.nativeReadback === 'passed' && native.readOnly === true && native.sha256 === hash(bytes), 'NATIVE_READBACK_FAILED', 'Draft bytes were not confirmed.');
      check(native.images?.length === new Set(valid.slides).size, 'RENDER_INCOMPLETE', 'Some draft pages were not rendered.');
      const { images, ...nativeValidation } = native;
      const preview = { documentId: session.documentId, revision: session.revision, generation: session.generation,
        basis: 'current-checkpoint-bytes', checkpointHash: hash(bytes), images, nativeValidation, visualReview: 'not_run' };
      await writeJson(path.join(directory, 'preview.json'), preview);
      await fs.unlink(candidate);
      return valid.detail === 'summary' ? { ...preview, nativeValidation: summarizeNative(nativeValidation) } : preview;
    } finally { if (docLock) await FileLock.release(docLock); await FileLock.release(taskLock); }
  }

  async assertReviewOutput(review) {
    check(await exists(review.outputPath), 'OUTPUT_CHANGED', 'Committed output is missing. Preserve recovery data and publish a new output.');
    const bytes = await fs.readFile(review.outputPath);
    check(hash(bytes) === review.outputHash, 'OUTPUT_CHANGED', 'Committed output bytes changed. Publish and review a new output.');
    return bytes;
  }

  async renderCommitted(valid) {
    // Same task -> document lock order as finish; edits cannot change this revision
    // while its evidence is being generated. External file writes are hash-checked.
    const taskLock = await FileLock.acquire(path.join(this.taskDir, 'locks', 'task.lock'), { context: 'render-committed' });
    let docLock;
    try {
      await this.assertTaskActive();
      const directory = path.join(this.taskDir, 'documents', valid.documentId);
      docLock = await FileLock.acquire(path.join(directory, 'mutation.lock'), { context: 'render-committed' });
      check(await exists(path.join(directory, 'meta.json')), 'SESSION_NOT_FOUND', 'Document session not found.');
      const meta = await readJson(path.join(directory, 'meta.json'));
      check(valid.allowOffice && meta.allowOffice, 'OFFICE_NOT_ALLOWED', 'Rendering requires explicit Office permission on the document and request.');
      const reviewPath = path.join(this.taskDir, 'reviews', `${valid.reviewId}.json`);
      check(await exists(reviewPath), 'REVIEW_NOT_FOUND', 'Committed review not found in this task.');
      const review = await readJson(reviewPath);
      check(review.documentId === valid.documentId, 'REVIEW_DOCUMENT_MISMATCH', 'Review belongs to a different document.');
      check(!review.stale && review.revision === meta.revision && (review.generation === undefined || review.generation === meta.generation),
        'REVIEW_STALE', 'Document changed since commit. Commit the current revision before rendering.');
      const bytes = await this.assertReviewOutput(review);
      validatePackage(await readPackage(bytes));
      // Assets live under reviews so finish can remove checkpoints without losing previews.
      const renderDir = path.join(this.taskDir, 'reviews', valid.reviewId, `render-${randomUUID()}`);
      await fs.mkdir(renderDir, { recursive: true });
      const candidate = path.join(renderDir, 'input.pptx');
      await publishNew(candidate, bytes);
      const host = await this.ensureNativeHost();
      const native = await host.request('validateFile', { path: candidate, sha256: review.outputHash,
        render: { slides: valid.slides, width: valid.width, directory: renderDir } });
      check(native.nativeReadback === 'passed' && native.readOnly === true && native.sha256 === review.outputHash,
        'NATIVE_READBACK_FAILED', 'PowerPoint did not confirm the committed bytes.');
      await this.assertReviewOutput(review);
      const { images, ...nativeValidation } = native;
      check(images?.length === new Set(valid.slides).size, 'RENDER_INCOMPLETE', 'Some requested preview pages were not rendered.');
      const preview = { reviewId: review.reviewId, documentId: valid.documentId, revision: review.revision,
        outputPath: review.outputPath, outputHash: review.outputHash, basis: 'committed-output-bytes',
        images, nativeValidation, width: valid.width, visualReview: 'not_run', verifiedAt: new Date().toISOString() };
      if (valid.layoutCheck) {
        const session = await this.getSession(valid.documentId);
        check(session.revision === review.revision && session.generation === review.generation,
          'REVIEW_STALE', 'Document changed before recording render layout evidence.');
        // validateFile already reads every object from these exact committed
        // bytes, even when only some pages are rendered. Reuse that snapshot.
        preview.layoutAudit = await this.recordLayoutAudit(session, nativeValidation.snapshot, { detail: valid.detail });
        preview.reviewBundlePath = reviewPath;
      }
      await fs.unlink(candidate); // Only a confirmed closed read-only candidate is disposable.
      const merged = new Map((review.preview?.images || []).map(image => [image.slide, image]));
      for (const image of images) merged.set(image.slide, image);
      review.preview = { ...preview, images: [...merged.values()].sort((a,b) => a.slide-b.slide) }; review.visualReviewed = false;
      delete review.reviewedAt; delete review.visualReviewMethod;
      await writeJson(reviewPath, review); this.reviews.set(review.reviewId, review);
      return valid.detail === 'summary' ? { ...preview, nativeValidation: summarizeNative(nativeValidation) } : preview;
    } finally { await FileLock.release(docLock); await FileLock.release(taskLock); }
  }

  /**
   * 每次 CLI 命令执行结束后的资源清理（释放当前进程的原生 host 与临时锁，不把任务标记为 finished）
   */
  async cleanupCommand() {
    let report;
    try {
      if (this.nativeHost) {
        try { report = await this.nativeHost.shutdown(); }
        catch (error) {
          report = { workerExited: false, cleanup: 'worker_not_responsive', outcome: 'outcome_unknown',
            errors: [{ code: error.code || 'NATIVE_SHUTDOWN_FAILED', message: error.message }] };
        }
        const info = await readJson(path.join(this.taskDir, 'task.json'));
        await writeJson(path.join(this.taskDir, 'task.json'), { ...info, cleanupReport: report });
        assertCleanupConfirmed(report);
      } else {
        report = await this.assertPreviousCleanup();
      }
      return report || { workerExited: true, office: 'not_started' };
    } finally {
      this.nativeHost = null;
      if (this.job) { this.job.close(); this.job = null; }
      for (const [id, session] of this.sessions) if (session.mode === 'native-copy') this.sessions.delete(id);
    }
  }

  async assertPreviousCleanup() {
    const taskPath = path.join(this.taskDir, 'task.json');
    if (!this.initialized) await assertTaskDirBoundary(this.baseDir, this.taskDir, this.taskId);
    if (!await exists(taskPath)) return;
    const info = await readJson(taskPath);
    if (info.cleanupReport) assertCleanupConfirmed(info.cleanupReport);
    return info.cleanupReport;
  }

  /**
   * 离线状态查询：纯磁盘读取，不启动任何后台进程或 Office。
   */
  static async status(taskDir, { operationId, reviewId } = {}) {
    contracts.ppt_status.parse({ operationId, reviewId });
    const taskJsonPath = path.join(taskDir, 'task.json');
    if (!(await exists(taskJsonPath))) {
      return { status: 'not_found', taskDir };
    }
    const taskInfo = await readJson(taskJsonPath);
    let operationReceipt = null;
    if (operationId) {
      const receiptPath = path.join(taskDir, 'receipts', `${operationId}.json`);
      if (await exists(receiptPath)) {
        operationReceipt = await readJson(receiptPath);
      }
    }
    const bindings = await Promise.all((await fs.readdir(path.join(taskDir, 'bindings')).catch(error => { if (error.code === 'ENOENT') return []; throw error; }))
      .filter(name => /^[0-9a-f-]{36}\.json$/i.test(name)).map(name => readJson(path.join(taskDir, 'bindings', name))));
    const documents = await Promise.all(bindings.map(async binding => {
      check(UUID_REGEX.test(binding.documentId), 'INVALID_BINDING', 'Document binding ID is invalid.');
      const meta = await readOptional(path.join(taskDir, 'documents', binding.documentId, 'meta.json'));
      const audit = await readOptional(path.join(taskDir, 'layout-audits', `${binding.documentId}.json`));
      return { documentId: binding.documentId, mode: binding.mode, closed: meta?.closed ?? true, historical: !meta,
        revision: meta?.revision ?? audit?.revision, generation: meta?.generation ?? audit?.generation,
        layout: auditSummary(audit, meta?.revision ?? audit?.revision, meta?.generation ?? audit?.generation) };
    }));
    return {
      ...taskInfo, bindings, documents,
      operationReceipt,
      ...(reviewId ? { reviewBundle: await readJson(path.join(taskDir, 'reviews', `${reviewId}.json`)).catch(error => { if (error.code === 'ENOENT') return null; throw error; }) } : {})
    };
  }
}
