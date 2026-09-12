import test from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { buildDeck } from '../src/build.js';
import { NativeHost } from '../src/native-host.js';
import { createOwnedJob, listProcesses, processIdentity, sameProcess, waitForExit, observeForeground, restrictAccess } from '../src/windows.js';
import { hash, inside, writeJson } from '../src/storage.js';

const winax = createRequire(import.meta.url)('winax');
const release = (...refs) => winax.release(...refs.filter(Boolean));
assert.equal(typeof globalThis.gc, 'function', 'Use npm run test:lifecycle so test-owned temporary COM references can be released.');

async function scenario(kind) {
  assert.equal(listProcesses('POWERPNT.EXE').length, 0, 'Lifecycle tests need a desktop without an existing PowerPoint; existing applications are never closed by setup.');
  const root = path.resolve('work/lifecycle-verification', randomUUID());
  await fs.mkdir(root, { recursive: true }); restrictAccess(root, { directory: true });
  const source = await buildDeck({ slides: [{ items: [{ type: 'text', text: 'Task original', left: 60, top: 60, width: 500, height: 80, name: 'Heading' }] }] });
  const working = path.join(root, 'working.pptx'), externalPath = path.join(root, 'external.pptx');
  await fs.writeFile(working, source); await fs.writeFile(externalPath, source);
  const job = createOwnedJob(), host = new NativeHost(job), observer = observeForeground();
  const report = { kind, root, node: process.version };
  let application, external, externalText, officeIdentity, shutdown;
  const openExternal = () => {
    // Activate an independent automation client in PowerPoint's shared process.
    // GetActiveObject borrows the task's automation object and its lifetime.
    application = new winax.Object('PowerPoint.Application', { activate: false });
    const presentations = application.Presentations;
    try { external = presentations.Open(externalPath, 0, 0, 0); }
    finally { release(presentations); }
    const slide = external.Slides.Item(1), shape = slide.Shapes.Item(1), frame = shape.TextFrame;
    externalText = frame.TextRange;
    externalText.Text = 'Unsaved external edit';
    release(frame, shape, slide);
  };
  try {
    if (kind === 'preexisting') openExternal();
    await host.start();
    const opened = await host.request('open', { documentId: 'task', path: working, directory: root, visible: false });
    report.lease = opened.lease;
    const processes = listProcesses('POWERPNT.EXE');
    assert.equal(processes.length, 1);
    officeIdentity = processIdentity(processes[0].pid);
    report.officeIdentity = officeIdentity;
    assert.equal(job.contains(officeIdentity.pid), false);
    if (kind === 'adopted') openExternal();
    const target = opened.snapshot.objects.find(object => object.name === 'Heading');
    const applied = await host.request('apply', { documentId: 'task', generation: opened.generation, operations: [
      { type: 'replace_text', target, search: 'original', replacement: 'revised', expectedMatches: 1, crossRunPolicy: 'reject' }
    ] });
    assert.equal(applied.outcome, 'completed');
    assert.equal(applied.snapshot.objects.find(object => object.name === 'Heading').text, 'Task revised');
    shutdown = await host.shutdown(); report.shutdown = shutdown;
    assert.equal(shutdown.workerExited, true);
    assert.equal(host.worker.exitCode, 0);
    assert.deepEqual(shutdown.errors, []);
    if (kind === 'cold') {
      assert.equal(shutdown.lease.createdByTask, true, 'Cold-start application ownership must be established');
      assert.equal(shutdown.lease.quitRequested, true);
      assert.equal(shutdown.officeExited, true, 'Cold-start PowerPoint must actually exit');
    } else {
      assert.equal(shutdown.lease.quitRequested, false);
      assert.equal(sameProcess(officeIdentity), true, 'External presentation must retain its application');
      assert.equal(String(external.FullName).toLowerCase(), externalPath.toLowerCase());
      assert.equal(String(externalText.Text), 'Unsaved external edit');
      assert.equal(Number(external.Saved), 0, 'Unsaved external state must be preserved');
      assert.equal(hash(await fs.readFile(externalPath)), hash(source));
      if (kind === 'preexisting') assert.equal(shutdown.lease.createdByTask, false);
      else assert.equal(shutdown.lease.createdByTask, true);
      report.externalPreserved = true;
    }
    report.behaviorPassed = true;
  } catch (error) { report.error = { message: error.message, stack: error.stack }; throw error; }
  finally {
    try {
      if (!shutdown) await host.shutdown(true).catch(error => { report.workerCleanupError = error.message; });
      job.close();
      // Test cleanup owns only generated fixture files and the uniquely observed
      // application started after the empty-desktop check. Never terminate by name.
      const current = listProcesses('POWERPNT.EXE');
      if (officeIdentity && sameProcess(officeIdentity) && current.length === 1 && current[0].pid === officeIdentity.pid) {
        application ||= new winax.Object('PowerPoint.Application', { activate: true });
        const presentations = application.Presentations;
        try {
          for (let i = Number(presentations.Count); i >= 1; i--) {
            const pres = presentations.Item(i);
            try { if (inside(root, String(pres.FullName))) { pres.Saved = -1; pres.Close(); } }
            finally { release(pres); }
          }
          const protectedViews = application.ProtectedViewWindows;
          try { if (Number(presentations.Count) === 0 && Number(protectedViews.Count) === 0) application.Quit(); }
          finally { release(protectedViews); }
        } finally { release(presentations); }
      }
      release(externalText, external, application);
      await new Promise(resolve => setImmediate(resolve));
      globalThis.gc();
      report.fixtureOfficeExited = officeIdentity ? await waitForExit(officeIdentity) : null;
      report.foreground = observer.stop();
      if (report.behaviorPassed) assert.equal(report.fixtureOfficeExited, true, 'Test fixture application must exit after cleanup');
      report.passed = report.behaviorPassed === true && report.fixtureOfficeExited === true;
    } catch (error) {
      report.passed = false;
      report.cleanupError = error.message;
      throw error;
    } finally {
      await writeJson(path.join(root, 'report.json'), report);
      console.log(JSON.stringify({ kind, root, passed: report.passed || false, behaviorPassed: report.behaviorPassed || false, shutdown: report.shutdown, externalPreserved: report.externalPreserved, fixtureOfficeExited: report.fixtureOfficeExited }));
    }
  }
}

await test('cold-start native editing closes its PowerPoint and Node processes', () => scenario('cold'));
await test('preexisting unsaved presentation survives native task shutdown', () => scenario('preexisting'));
await test('external presentation opened during a task prevents application Quit', () => scenario('adopted'));
