import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import path from 'node:path';

// Keep stdout readable after stdin EOF so tests can observe the real server
// draining an accepted operation, without a client's forced-kill timeout.
export async function startStdioPeer(baseDir, taskId) {
  const args = [path.resolve('src/mcp.js'), '--base-dir', baseDir];
  if (taskId) args.push('--task', taskId);
  const child = spawn(process.execPath, args, { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] });
  let sequence = 0;
  let stderr = '';
  const pending = new Map();
  child.stderr.on('data', chunk => { stderr = (stderr + chunk).slice(-3000); });
  const lines = createInterface({ input: child.stdout });
  lines.on('line', line => {
    let message;
    try { message = JSON.parse(line); } catch { return; }
    const waiter = pending.get(message.id);
    if (!waiter) return;
    pending.delete(message.id);
    clearTimeout(waiter.timer);
    if (message.error) waiter.reject(Object.assign(new Error(message.error.message), { code: message.error.code }));
    else waiter.resolve(message.result);
  });
  const failPending = error => {
    for (const waiter of pending.values()) { clearTimeout(waiter.timer); waiter.reject(error); }
    pending.clear();
  };
  child.on('error', failPending);
  child.stdin.on('error', failPending);
  const exited = new Promise(resolve => child.once('close', (code, signal) => {
    lines.close();
    failPending(new Error(`MCP exited before responding: ${code}/${signal}`));
    resolve({ code, signal, stderr });
  }));
  const request = (method, params) => new Promise((resolve, reject) => {
    const id = ++sequence;
    const timer = setTimeout(() => { pending.delete(id); reject(new Error(`MCP request timed out: ${method}`)); }, 60000);
    pending.set(id, { resolve, reject, timer });
    child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
  });
  const endInput = () => child.stdin.end();
  const stop = async () => {
    endInput();
    let timer;
    try {
      return await Promise.race([exited, new Promise((_, reject) => {
        timer = setTimeout(() => {
          child.kill(); // This exact test child only; never enumerate/kill Office.
          reject(new Error('MCP did not drain and exit within 45 seconds.'));
        }, 45000);
      })]);
    } finally { clearTimeout(timer); }
  };
  try {
    const initialized = await request('initialize', { protocolVersion: '2025-11-25', capabilities: {}, clientInfo: { name: 'ppt-eof-regression', version: '1.0.0' } });
    child.stdin.write(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' }) + '\n');
    const call = async (name, args = {}) => {
      const result = await request('tools/call', { name, arguments: args });
      if (result.isError) {
        const error = JSON.parse(result.content[0].text);
        throw Object.assign(new Error(error.error), { code: error.code, details: error.details });
      }
      return JSON.parse(result.content[0].text);
    };
    return { child, initialized, request, call, endInput, stop };
  } catch (error) { await stop(); throw error; }
}
