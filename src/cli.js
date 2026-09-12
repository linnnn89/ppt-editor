#!/usr/bin/env node
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { TaskHost, validateTaskId, assertTaskDirBoundary } from './task.js';
import { contracts } from './contracts.js';
import { PptError, check } from './errors.js';

/**
 * 读取完整的标准输入（用于安全接收复杂大体积或多行 JSON payload，避免 Windows shell 转义灾难）
 */
async function readStdin() {
  const chunks = [];
  for await (const chunk of process.stdin) {
    chunks.push(chunk);
  }
  return Buffer.concat(chunks).toString('utf8');
}

/**
 * 极简、鲁棒的命令行参数解析器（遵循零外部依赖与原生能力优先原则）
 */
function parseArgs(args) {
  const flags = {};
  let command = null;
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (!command && !arg.startsWith('-')) {
      command = arg;
      continue;
    }
    if (arg.startsWith('--')) {
      const key = arg.slice(2);
      if (i + 1 < args.length && !args[i + 1].startsWith('-')) {
        flags[key] = args[++i];
      } else {
        flags[key] = true;
      }
    }
  }
  return { command, flags };
}

function toBoolean(val, fallback = false) {
  if (val === undefined || val === null) return fallback;
  if (typeof val === 'boolean') return val;
  if (typeof val === 'string') {
    const s = val.trim().toLowerCase();
    if (s === 'true' || s === '1') return true;
    if (s === 'false' || s === '0') return false;
  }
  return Boolean(val);
}

/**
 * CLI 核心调度入口
 */
async function main() {
  const { command, flags } = parseArgs(process.argv.slice(2));

  if (!command || flags.help) {
    console.log(JSON.stringify({
      usage: 'ppt-task <command> [options]',
      commands: ['start', 'open', 'build', 'compose', 'relayout', 'inspect', 'apply', 'validate', 'commit', 'render', 'diagnose', 'close', 'finish', 'status'],
      examples: [
        'ppt-task start',
        'ppt-task open --task <id> --path <file> --mode file --op op1',
        'ppt-task inspect --task <id> --doc <docId>',
        'ppt-task apply --task <id> --doc <docId> --rev 0 --op op2 --operations \'<json>\'',
        'ppt-task commit --task <id> --doc <docId> --rev 1 --out <output> --op op3',
        'ppt-task finish --task <id> --reviews \'["<reviewId>"]\'',
        'ppt-task status --task <id>'
      ]
    }, null, 2));
    process.exit(0);
  }

  // 基础参数解析
  const baseDir = flags['base-dir'] ? path.resolve(flags['base-dir']) : path.resolve('work/tasks');
  const taskId = flags.task || flags['task-id'];

  // 如果指定了 --stdin，读取完整标准输入 JSON 并与 flags 合并
  let stdinData = {};
  if (flags.stdin) {
    const rawStdin = await readStdin();
    if (rawStdin.trim()) {
      try {
        stdinData = JSON.parse(rawStdin.trim());
      } catch (err) {
        throw new PptError('INVALID_INPUT', 'Failed to parse JSON from stdin.', { reason: err.message });
      }
    }
  }

  // 1. status 命令：离线查询，不依赖 TaskHost 实例，不启动任何后台进程
  if (command === 'status') {
    check(taskId, 'MISSING_ARGUMENT', 'Task ID (--task) is required for status.');
    const taskDir = path.join(baseDir, taskId);
    validateTaskId(taskId);
    await assertTaskDirBoundary(baseDir, taskDir, taskId);
    const operationId = flags.op || flags['operation-id'] || stdinData.operationId;
    const statusParams = contracts.ppt_status.parse({ operationId, reviewId: flags.review || stdinData.reviewId });
    const result = await TaskHost.status(taskDir, statusParams);
    console.log(JSON.stringify(result));
    return;
  }

  // 2. start 命令：创建并初始化一个新任务
  if (command === 'start') {
    const host = new TaskHost({ taskId, baseDir });
    try {
      await host.init();
      console.log(JSON.stringify({
        status: 'active',
        taskId: host.taskId,
        taskDir: host.taskDir
      }));
    } finally {
      await host.cleanupCommand().catch(() => {});
    }
    return;
  }

  // 后续命令均需要 taskId
  check(taskId || stdinData.taskId, 'MISSING_ARGUMENT', 'Task ID (--task) is required.');
  const resolvedTaskId = taskId || stdinData.taskId;
  const host = new TaskHost({ taskId: resolvedTaskId, baseDir });
  let outputResult;
  try {
    await host.init();

  switch (command) {
    case 'compose': {
      const { taskId: inputTaskId, ...composeInput } = stdinData;
      outputResult = await host.compose({ ...composeInput, templatePath: flags.template || stdinData.templatePath,
        sourcePath: flags.source || stdinData.sourcePath, operationId: flags.op || stdinData.operationId,
        layoutCheck: toBoolean(flags['layout-check'] ?? stdinData.layoutCheck, true),
        allowOffice: toBoolean(flags['allow-office'] ?? stdinData.allowOffice) });
      break;
    }
    case 'diagnose': {
      outputResult = await host.diagnose();
      break;
    }
    case 'render': {
      outputResult = await host.render({
        documentId: flags.doc || flags['document-id'] || stdinData.documentId,
        reviewId: flags.review || stdinData.reviewId,
        slides: flags.slides ? JSON.parse(flags.slides) : stdinData.slides,
        width: flags.width ? Number(flags.width) : stdinData.width,
        allowOffice: toBoolean(flags['allow-office'] ?? stdinData.allowOffice),
        expectedRevision: flags.rev !== undefined ? Number(flags.rev) : stdinData.expectedRevision,
        detail: flags.detail || stdinData.detail,
        layoutCheck: toBoolean(flags['layout-check'] ?? stdinData.layoutCheck)
      });
      break;
    }
    case 'open': {
      const openParams = {
        path: flags.path || stdinData.path,
        mode: flags.mode || stdinData.mode || 'file',
        allowOffice: toBoolean(flags['allow-office'] ?? stdinData.allowOffice),
        visible: toBoolean(flags.visible ?? stdinData.visible),
        operationId: flags.op || flags['operation-id'] || stdinData.operationId
      };
      outputResult = await host.open(openParams);
      break;
    }

    case 'build': {
      let deck = stdinData.deck;
      if (!deck && flags.deck) {
        deck = JSON.parse(flags.deck);
      } else if (!deck && flags['deck-file']) {
        deck = JSON.parse(await fs.readFile(path.resolve(flags['deck-file']), 'utf8'));
      }
      check(deck, 'MISSING_ARGUMENT', 'Deck specification (--deck or --deck-file or stdin) is required.');
      const buildParams = {
        deck,
        layoutCheck: toBoolean(flags['layout-check'] ?? stdinData.layoutCheck, true),
        mode: flags.mode || stdinData.mode || 'file',
        allowOffice: toBoolean(flags['allow-office'] ?? stdinData.allowOffice),
        visible: toBoolean(flags.visible ?? stdinData.visible),
        operationId: flags.op || flags['operation-id'] || stdinData.operationId
      };
      outputResult = await host.build(buildParams);
      break;
    }

    case 'inspect': {
      const inspectParams = {
        documentId: flags.doc || flags['document-id'] || stdinData.documentId,
        text: flags.text || stdinData.text,
        slide: flags.slide ? Number(flags.slide) : stdinData.slide,
        kind: flags.kind || stdinData.kind,
        offset: flags.offset ? Number(flags.offset) : stdinData.offset ?? 0,
        limit: flags.limit ? Number(flags.limit) : stdinData.limit ?? 100,
        detail: flags.detail || stdinData.detail
      };
      outputResult = await host.inspect(inspectParams);
      break;
    }

    case 'relayout': {
      outputResult = await host.relayout({ ...stdinData,
        documentId: flags.doc || flags['document-id'] || stdinData.documentId,
        expectedRevision: flags.rev !== undefined ? Number(flags.rev) : stdinData.expectedRevision,
        operationId: flags.op || flags['operation-id'] || stdinData.operationId,
        dryRun: toBoolean(flags['dry-run'] ?? stdinData.dryRun, true)
      });
      break;
    }

    case 'apply': {
      let operations = stdinData.operations;
      if (!operations && flags.operations) {
        operations = JSON.parse(flags.operations);
      } else if (!operations && flags['operations-file']) {
        operations = JSON.parse(await fs.readFile(path.resolve(flags['operations-file']), 'utf8'));
      }
      check(Array.isArray(operations), 'MISSING_ARGUMENT', 'Operations list (--operations, --operations-file or stdin) is required.');
      const applyParams = {
        documentId: flags.doc || flags['document-id'] || stdinData.documentId,
        expectedRevision: flags.rev !== undefined ? Number(flags.rev) : stdinData.expectedRevision,
        operationId: flags.op || flags['operation-id'] || stdinData.operationId,
        operations,
        layoutCheck: toBoolean(flags['layout-check'] ?? stdinData.layoutCheck, true),
        dryRun: toBoolean(flags['dry-run'] ?? stdinData.dryRun)
      };
      outputResult = await host.apply(applyParams);
      break;
    }

    case 'validate': {
      const validateParams = {
        documentId: flags.doc || flags['document-id'] || stdinData.documentId,
        layoutCheck: toBoolean(flags['layout-check'] ?? stdinData.layoutCheck),
        slides: flags.slides ? JSON.parse(flags.slides) : stdinData.slides,
        expectedRevision: flags.rev !== undefined ? Number(flags.rev) : stdinData.expectedRevision,
        allowedOverlapPairs: stdinData.allowedOverlapPairs,
        checks: flags.checks || stdinData.checks,
        detail: flags.detail || stdinData.detail,
        nativeReadback: toBoolean(flags['native-readback'] ?? stdinData.nativeReadback),
        allowOffice: toBoolean(flags['allow-office'] ?? stdinData.allowOffice)
      };
      outputResult = await host.validate(validateParams);
      break;
    }

    case 'commit': {
      const commitParams = {
        documentId: flags.doc || flags['document-id'] || stdinData.documentId,
        expectedRevision: flags.rev !== undefined ? Number(flags.rev) : stdinData.expectedRevision,
        outputPath: flags.out || flags['output-path'] || stdinData.outputPath,
        operationId: flags.op || flags['operation-id'] || stdinData.operationId
      };
      outputResult = await host.commit(commitParams);
      break;
    }

    case 'close': {
      const closeParams = {
        documentId: flags.doc || flags['document-id'] || stdinData.documentId,
        preserveCheckpoint: toBoolean(flags['preserve-checkpoint'] ?? stdinData.preserveCheckpoint)
      };
      outputResult = await host.close(closeParams);
      break;
    }

    case 'finish': {
      let reviewIds = stdinData.reviewIds;
      if (!reviewIds && flags.reviews) {
        reviewIds = JSON.parse(flags.reviews);
      }
      check(Array.isArray(reviewIds), 'MISSING_ARGUMENT', 'Review IDs list (--reviews or stdin) is required.');
      const finishParams = {
        reviewIds,
        requireAccepted: toBoolean(flags['require-accepted'] ?? stdinData.requireAccepted),
        preserveCheckpoints: toBoolean(flags['preserve-checkpoints'] ?? stdinData.preserveCheckpoints)
      };
      outputResult = await host.finish(finishParams);
      break;
    }

    default:
      throw new PptError('UNKNOWN_COMMAND', `Unknown command: ${command}`);
  }

  } finally {
    await host.cleanupCommand();
  }
  // Report success only after command-owned resources have been released.
  console.log(JSON.stringify(outputResult));
}

// 统一异常捕获与结构化错误输出
main().catch(error => {
  const errorPayload = {
    error: true,
    code: error.code || 'UNKNOWN_ERROR',
    message: error.message,
    details: error.details || null
  };
  process.stderr.write(JSON.stringify(errorPayload, null, 2) + '\n');
  process.exit(1);
});
