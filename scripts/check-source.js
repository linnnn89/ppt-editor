import { readdir } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import path from 'node:path';

let count = 0;
async function walk(directory) {
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const file = path.join(directory, entry.name);
    if (entry.isDirectory()) await walk(file);
    else if (/\.(?:c|m)?js$/.test(file)) { execFileSync(process.execPath, ['--check', file], { stdio: 'pipe', windowsHide: true }); count++; }
  }
}
for (const directory of ['src', 'scripts', 'tests']) await walk(directory).catch(error => { if (error.code !== 'ENOENT') throw error; });
console.log(`Syntax checked ${count} JavaScript files.`);
