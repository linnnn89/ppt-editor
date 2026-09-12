import { spawnSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

// nodejs/node#65446: ObjectWrap compiled with 24.19.0 headers can abort
// during allocation-driven GC. The 24.18.1 headers keep the Node 24 ABI
// without that incomplete backport; the installed runtime stays unchanged.
const target = '24.18.1';
if (process.platform !== 'win32' || process.arch !== 'x64' || process.versions.node.split('.')[0] !== '24') {
  throw new Error('The native build requires Windows x64 and Node 24.');
}
if (!process.env.npm_execpath) throw new Error('Run this build with npm run build:native.');

const require = createRequire(import.meta.url);
const npmRequire = createRequire(process.env.npm_execpath);
const nodeGyp = npmRequire.resolve('node-gyp/bin/node-gyp.js');
const directory = path.dirname(require.resolve('winax/package.json'));
const cache = fileURLToPath(new URL('../node_modules/.cache/node-gyp', import.meta.url));
console.log(`Building winax with Node ${target} headers for runtime ${process.version}.`);
const build = spawnSync(process.execPath, [nodeGyp, 'rebuild', `--directory=${directory}`, `--target=${target}`, '--arch=x64', `--devdir=${cache}`], {
  stdio: 'inherit', windowsHide: true
});
if (build.error) throw build.error;
if (build.status !== 0) throw new Error(`winax build failed (${build.status ?? build.signal}).`);

// Test in a disposable process so a broken addon cannot abort the installer.
const probe = fileURLToPath(new URL('./diagnose-native.cjs', import.meta.url));
const smoke = spawnSync(process.execPath, [probe, '--child', 'allocation-gc'], {
  stdio: 'inherit', windowsHide: true, timeout: 15000
});
if (smoke.error) throw smoke.error;
if (smoke.status !== 0) throw new Error(`winax allocation-GC check failed (${smoke.status ?? smoke.signal}).`);
