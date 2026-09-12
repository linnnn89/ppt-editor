// Isolate native-GC crashes from Office and from the presentation contents.
const { spawnSync } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');

if (process.argv[2] === '--child') {
  const winax = require('winax');
  const mode = process.argv[3];
  if (mode === 'allocation-gc') {
    // Explicit global.gc() can miss nodejs/node#65446. Let ordinary
    // allocations reclaim ObjectWrap instances while the process is running.
    let retained = [];
    for (let i = 0; i < 300000; i++) {
      new winax.Variant(i, 'int');
      retained.push({ value: i });
      if (retained.length > 1000) retained = [];
    }
    console.log(JSON.stringify({ phase: 'allocation_gc_passed', mode }));
  }
  for (let i = 0; i < 100; i++) {
    let dictionary = new winax.Object('Scripting.Dictionary');
    dictionary.Add('example', 'value');
    if (String(dictionary.Item('example')) !== 'value') throw new Error('COM readback failed');
    winax.release(dictionary);
    dictionary = null;
  }
  console.log(JSON.stringify({ phase: 'com_readback_and_release_passed', mode }));
  if (mode === 'synchronous-gc') {
    global.gc();
    console.log(JSON.stringify({ phase: 'gc_passed', mode }));
  } else if (mode === 'asynchronous-gc') {
    global.gc({ execution: 'async', type: 'major' }).then(() => {
      console.log(JSON.stringify({ phase: 'gc_passed', mode }));
    });
  } else if (mode === 'native-callback-gc') {
    const zlib = require('node:zlib');
    const input = zlib.gzipSync(Buffer.alloc(16 * 1024 * 1024, 65));
    let pending = 8;
    for (let i = 0; i < pending; i++) {
      zlib.gunzip(input, (error, bytes) => {
        if (error || bytes.length !== 16 * 1024 * 1024) throw error || new Error('Decompression failed');
        if (--pending === 0) console.log(JSON.stringify({ phase: 'native_callbacks_passed', mode }));
      });
    }
  }
} else {
  const root = path.resolve('work', 'native-diagnosis', new Date().toISOString().replace(/[:.]/g, '-'));
  fs.mkdirSync(root, { recursive: true });
  const report = { node: process.version, executable: process.execPath, winax: require('winax/package.json').version, officeStarted: false, results: [] };
  for (const mode of process.argv.length > 2 ? process.argv.slice(2) : ['no-forced-gc', 'synchronous-gc', 'asynchronous-gc', 'native-callback-gc', 'allocation-gc']) {
    const result = spawnSync(process.execPath, ['--expose-gc', __filename, '--child', mode], { encoding: 'utf8', windowsHide: true, timeout: 15000 });
    report.results.push({ mode, exitCode: result.status, signal: result.signal, error: result.error?.message, stdout: result.stdout, stderr: result.stderr });
  }
  fs.writeFileSync(path.join(root, 'gc-report.json'), JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify({ root, ...report }, null, 2));
  if (report.results.some(result => result.exitCode !== 0 || result.error)) process.exitCode = 1;
}
