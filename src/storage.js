import { promises as fs } from 'node:fs';
import path from 'node:path';
import { createHash, randomUUID } from 'node:crypto';
import { check } from './errors.js';

export const hash = value => createHash('sha256').update(value).digest('hex');
export const canonicalPath = value => path.resolve(value).replaceAll('\\', '/').toLowerCase();
export async function exists(file) { try { await fs.access(file); return true; } catch { return false; } }

export function stableJson(value) {
  if (Array.isArray(value)) return `[${value.map(stableJson).join(',')}]`;
  if (value && typeof value === 'object') return `{${Object.keys(value).sort().filter(k => value[k] !== undefined).map(k => `${JSON.stringify(k)}:${stableJson(value[k])}`).join(',')}}`;
  return JSON.stringify(value);
}

export async function readJson(file) { return JSON.parse(await fs.readFile(file, 'utf8')); }

export async function atomicWrite(file, data) {
  await fs.mkdir(path.dirname(file), { recursive: true });
  const temp = `${file}.${randomUUID()}.tmp`;
  const handle = await fs.open(temp, 'wx', 0o600);
  try { await handle.writeFile(data); await handle.sync(); } finally { await handle.close(); }
  try { await fs.rename(temp, file); }
  catch (error) { await fs.unlink(temp).catch(() => {}); throw error; }
}

export async function writeJson(file, value) { await atomicWrite(file, JSON.stringify(value, null, 2) + '\n'); }

// Preflight before invoking Office. This does not replace publishNew's atomic
// no-overwrite guarantee: another process can still appear after this check.
export async function prepareNewOutput(file) {
  check(path.extname(file).toLowerCase() === '.pptx', 'UNSUPPORTED_FORMAT', 'Output must be a .pptx file.');
  let occupied = false;
  try { await fs.lstat(file); occupied = true; }
  catch (error) { if (error.code !== 'ENOENT') throw error; }
  check(!occupied, 'OUTPUT_ALREADY_EXISTS', 'Output path is occupied; choose a new path before saving.');
  const directory = path.dirname(file);
  await fs.mkdir(directory, { recursive: true });
  const probe = path.join(directory, `.ppt-write-probe-${randomUUID()}`);
  const handle = await fs.open(probe, 'wx', 0o600);
  try { await handle.writeFile(''); }
  finally { await handle.close(); await fs.unlink(probe); }
}

export async function publishNew(file, bytes) {
  check(path.extname(file).toLowerCase() === '.pptx', 'UNSUPPORTED_FORMAT', 'Output must be a .pptx file.');
  await fs.mkdir(path.dirname(file), { recursive: true });
  const temporary = path.join(path.dirname(file), `.${path.basename(file)}.${randomUUID()}.tmp`);
  const handle = await fs.open(temporary, 'wx', 0o600);
  try { await handle.writeFile(bytes); await handle.sync(); } finally { await handle.close(); }
  // A hard link publishes without ever replacing an existing destination.
  // Unsupported filesystems fail safely and keep the candidate for recovery.
  try { await fs.link(temporary, file); }
  catch (error) {
    error.details = { ...(error.details || {}), candidatePath: temporary };
    if (error.code === 'EEXIST') error.message = 'Output already exists; choose a new output path.';
    throw error;
  }
  await fs.unlink(temporary);
  return { path: path.resolve(file), sha256: hash(bytes), bytes: bytes.length };
}

export function inside(root, candidate) {
  const relative = path.relative(path.resolve(root), path.resolve(candidate));
  return relative === '' || (!relative.startsWith('..' + path.sep) && relative !== '..' && !path.isAbsolute(relative));
}
