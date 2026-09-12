import net from 'node:net';
import { mkdir } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import path from 'node:path';
import { createRequire } from 'node:module';
import { currentUserSid, processIdentity, listProcesses, createOwnedJob, restrictAccess } from '../src/windows.js';

const require = createRequire(import.meta.url);
const winax = require('winax');
const directory = path.resolve('work/platform-probe');
await mkdir(directory, { recursive: true });
restrictAccess(directory, { directory: true });
const pipe = `\\\\.\\pipe\\ppt-editor-probe-${randomUUID()}`;
const server = net.createServer(socket => socket.end());
await new Promise((resolve, reject) => { server.once('error', reject); server.listen(pipe, resolve); });
try {
  restrictAccess(pipe);
  const job = createOwnedJob();
  try { console.log(JSON.stringify({ node: process.version, winaxLoaded: Boolean(winax.Object), userSidAvailable: Boolean(currentUserSid()), process: processIdentity(process.pid), powerpointProcessCount: listProcesses('POWERPNT.EXE').length, pipeAcl: true, jobCreated: true })); }
  finally { job.close(); }
} finally { await new Promise(resolve => server.close(resolve)); }
