#!/usr/bin/env node
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { McpServer } from '@modelcontextprotocol/server';
import { StdioServerTransport } from '@modelcontextprotocol/server/stdio';
import { TaskHost } from './task.js';
import { contracts, descriptions } from './contracts.js';
import { check } from './errors.js';
import { exists } from './storage.js';
import { validateTaskId, assertTaskDirBoundary } from './task.js';
import packageInfo from '../package.json' with { type: 'json' };
import { createRequestProgress } from './progress.js';

/**
 * 创建配置完备的 PowerPoint MCP 服务器实例。
 * 一条连接依次服务多个任务；任务切换与业务调用共用串行队列。
 */
export function createPptMcpServer({ taskHost, taskId, baseDir = 'work/tasks' } = {}) {
  let host = taskHost || new TaskHost({ taskId, baseDir });
  let queue = Promise.resolve();
  let stopping = false;
  const server = new McpServer({
    name: 'ppt-editor',
    version: packageInfo.version
  });

  const toolMappings = [
    ['ppt_start', async args => {
      const { taskId: nextId } = contracts.ppt_start.parse(args);
      validateTaskId(nextId);
      const current = await host.status();
      if (nextId === host.taskId) return current;
      check(current.status === 'finished', 'TASK_STILL_ACTIVE', 'Finish the current task before starting another.');
      const nextDir = path.resolve(host.baseDir, nextId);
      await assertTaskDirBoundary(host.baseDir, nextDir, nextId);
      check(!await exists(nextDir), 'TASK_ALREADY_EXISTS', 'Use a fresh task UUID; historical tasks are available through ppt_status.');
      const next = new TaskHost({ taskId: nextId, baseDir: host.baseDir });
      await next.init();
      host = next;
      return host.status();
    }],
    ['ppt_open', async args => host.open(args)],
    ['ppt_inspect', async args => host.inspect(args)],
    ['ppt_apply', async args => host.apply(args)],
    ['ppt_commit', async args => host.commit(args)],
    ['ppt_validate', async args => host.validate(args)],
    ['ppt_close', async args => host.close(args)],
    ['ppt_finish', async args => host.finish(args)],
    ['ppt_status', async args => host.status(args)],
    ['ppt_diagnose', async args => host.diagnose(args)],
    ['ppt_build', async args => host.build(args)],
    ['ppt_compose', async args => host.compose(args)],
    ['ppt_relayout', async args => host.relayout(args)],
    ['ppt_render', async args => host.render(args)]
  ];

  for (const [name, handler] of toolMappings) {
    server.registerTool(
      name,
      {
        description: descriptions[name] || `Execute ${name} on PowerPoint presentation`,
        inputSchema: contracts[name]
      },
      async (args, context) => {
        const receivedAt = performance.now();
        let timing;
        try {
          check(!stopping, 'SERVER_STOPPING', 'The MCP connection is closing.');
          const pending = queue.then(async () => {
            check(!context.mcpReq.signal.aborted, 'REQUEST_CANCELLED', 'Request cancelled before execution.');
            // Once started, let the operation retain its durable result or recovery
            // state. Cancelling the caller's wait must not interrupt persistence.
            const progress = createRequestProgress(context.mcpReq, name, receivedAt);
            let completed = false;
            try {
              const result = await progress.run(() => handler(args)); completed = true; return result;
            } finally { timing = await progress.finish(completed ? 'completed' : 'failed'); }
          });
          queue = pending.catch(() => {});
          const result = await pending;
          return {
            content: [{ type: 'text', text: JSON.stringify(result) }],
            structuredContent: result,
            _meta: { pptEditorTiming: timing }
          };
        } catch (error) {
          return {
            isError: true,
            ...(timing ? { _meta: { pptEditorTiming: timing } } : {}),
            content: [{
              type: 'text',
              text: JSON.stringify({
                error: error.message || String(error),
                code: error.code || 'INTERNAL_ERROR',
                details: error.details || null
              })
            }]
          };
        }
      }
    );
  }

  let cleanupPromise;
  const cleanup = () => {
    stopping = true;
    cleanupPromise ??= queue.then(() => host.cleanupCommand());
    return cleanupPromise;
  };
  return { server, get host() { return host; }, cleanup };
}

/**
 * 启动基于标准标准输入输出流（STDIO）的 MCP 服务守护入口。
 */
export async function servePptMcpStdio({ taskId, baseDir = 'work/tasks' } = {}) {
  const instance = createPptMcpServer({ taskId, baseDir });
  const { server } = instance;
  const transport = new StdioServerTransport();

  const cleanup = async () => {
    try {
      await instance.cleanup();
    } catch (error) {
      process.stderr.write(`MCP cleanup failed: ${error.code || 'CLEANUP_FAILED'}\n`);
      process.exitCode = 1;
    }
  };

  process.once('SIGINT', async () => {
    await cleanup();
    process.exit(process.exitCode || 0);
  });

  process.once('SIGTERM', async () => {
    await cleanup();
    process.exit(process.exitCode || 0);
  });

  process.stdin.once('close', async () => {
    await cleanup();
  });
  process.stdin.once('end', cleanup);

  await server.connect(transport);
  return { server, get host() { return instance.host; }, transport };
}

// 支持直接作为 CLI 服务启动
const isDirectRun = process.argv[1] && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url));
if (isDirectRun) {
  const args = process.argv.slice(2);
  const getFlag = (name, fallback) => {
    const idx = args.indexOf(name);
    return idx !== -1 && idx + 1 < args.length ? args[idx + 1] : fallback;
  };
  const taskId = getFlag('--task') || getFlag('--task-id') || process.env.PPT_TASK_ID;
  const baseDir = getFlag('--base-dir', 'work/tasks');

  servePptMcpStdio({ taskId, baseDir }).catch(err => {
    process.stderr.write(`Failed to start MCP server: ${err.message}\n`);
    process.exit(1);
  });
}
