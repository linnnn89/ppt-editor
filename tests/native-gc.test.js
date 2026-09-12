import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

test('winax survives allocation-driven GC and process exit without starting Office', { skip: process.platform !== 'win32' }, () => {
  const script = fileURLToPath(new URL('../scripts/diagnose-native.cjs', import.meta.url));
  const result = spawnSync(process.execPath, [script, '--child', 'allocation-gc'], {
    encoding: 'utf8', windowsHide: true, timeout: 15000
  });
  assert.ifError(result.error);
  assert.equal(result.status, 0, `Native child failed (${result.signal}):\n${result.stderr}`);
  const phases = result.stdout.trim().split(/\r?\n/).map(line => JSON.parse(line).phase);
  assert.ok(phases.includes('allocation_gc_passed'));
  assert.ok(phases.includes('com_readback_and_release_passed'));
});
