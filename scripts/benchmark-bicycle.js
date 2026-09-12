// Isolated creation benchmark. Neither adapter expands the public editing API.
import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { fork } from 'node:child_process';
import { once } from 'node:events';
import { createRequire } from 'node:module';
import { performance } from 'node:perf_hooks';
import { randomUUID } from 'node:crypto';
import { bicycle, bicycleDeck } from './fixtures/bicycle.js';
import { buildDeck } from '../src/build.js';
import { readPackage, writePackage, validatePackage, parseXml, xml, NS, descendants, child } from '../src/ooxml.js';
import { hash, writeJson } from '../src/storage.js';
import { initializeSta, pumpMessages, listProcesses, processIdentity, sameProcess, waitForExit, createOwnedJob, terminateOwnedProcess } from '../src/windows.js';

const WIDTH = 960, HEIGHT = 540, BACKGROUND = 'FFFFFF';
const rgb = s => parseInt(s.slice(0, 2), 16) + 256 * parseInt(s.slice(2, 4), 16) + 65536 * parseInt(s.slice(4), 16);

async function sanitizeAndValidate(bytes, items) {
  const parts = await readPackage(bytes);
  // Office can write the local profile name during SaveAs. Scrub both routes.
  for (const name of ['docProps/core.xml', 'docProps/app.xml']) {
    if (!parts.has(name)) continue;
    const doc = parseXml(parts.get(name));
    for (const el of Array.from(doc.getElementsByTagName('*'))) {
      if (['creator', 'lastModifiedBy', 'Company', 'Manager'].includes(el.localName)) el.parentNode.removeChild(el);
    }
    parts.set(name, Buffer.from(xml(doc)));
  }
  validatePackage(parts);
  const slide = parseXml(parts.get('ppt/slides/slide1.xml'));
  const names = descendants(slide, NS.p, 'cNvPr').map(n => n.getAttribute('name')).filter(n => n);
  for (const item of items) assert(names.includes(item.name), `Missing native object: ${item.name}`);
  assert.equal(descendants(slide, NS.p, 'pic').length, 0, 'Bicycle must not contain raster/vector pictures');
  assert.equal(descendants(slide, NS.p, 'sp').length + descendants(slide, NS.p, 'cxnSp').length, items.length);
  return writePackage(parts);
}

async function worker(config) {
  const { root, method, trial, items } = config;
  const directory = path.join(root, `${method}-${trial}`);
  await fs.mkdir(directory);
  const report = { method, trial, stages: {}, directory, node: process.version, shapeCount: items.length };
  let app, pres, winax, timer, uninitialize;
  const release = (...refs) => winax?.release(...refs.filter(Boolean));
  const stage = async (name, fn) => {
    process.send?.({ stage: name }); const t = performance.now();
    const result = await fn(); report.stages[name] = performance.now() - t; return result;
  };
  const close = () => { if (pres) { pres.Saved = -1; pres.Close(); release(pres); pres = null; } };
  const startOffice = () => stage('officeStartupMs', () => {
    assert.equal(listProcesses('POWERPNT.EXE').length, 0, 'Existing PowerPoint is never closed by this benchmark');
    uninitialize = initializeSta(); winax = createRequire(import.meta.url)('winax'); timer = setInterval(pumpMessages, 10);
    const activation = (BigInt(Date.now()) + 11644473600000n) * 10000n;
    app = new winax.Object('PowerPoint.Application', { activate: false });
    const processes = listProcesses('POWERPNT.EXE'); assert.equal(processes.length, 1);
    report.officeIdentity = processIdentity(processes[0].pid); process.send?.({ officeIdentity: report.officeIdentity });
    assert(BigInt(report.officeIdentity.created) >= activation);
    assert.equal(path.resolve(String(app.Path), 'POWERPNT.EXE').toLowerCase(), report.officeIdentity.image.toLowerCase());
    report.officeVersion = String(app.Version); report.officeBuild = String(app.Build);
    const docs = app.Presentations, pv = app.ProtectedViewWindows;
    try { assert.equal(Number(docs.Count), 0); assert.equal(Number(pv.Count), 0); assert.equal(Number(app.Visible), 0); }
    finally { release(pv, docs); }
  });
  try {
    let bytes;
    if (method === 'A') {
      bytes = await stage('creationMs', () => buildDeck(bicycleDeck(items)));
    } else {
      await startOffice();
      await stage('creationMs', () => {
        const docs = app.Presentations; try { pres = docs.Add(0); } finally { release(docs); }
        const setup = pres.PageSetup; setup.SlideWidth = WIDTH; setup.SlideHeight = HEIGHT; release(setup);
        const slides = pres.Slides, slide = slides.Add(1, 12), shapes = slide.Shapes;
        try {
          slide.FollowMasterBackground = 0;
          const bg = slide.Background, fill = bg.Fill, color = fill.ForeColor;
          fill.Solid(); color.RGB = rgb(BACKGROUND); release(color, fill, bg);
          for (const i of items) {
            const shape = i.type === 'line' ? shapes.AddLine(i.x1, i.y1, i.x2, i.y2) : shapes.AddShape(9, i.left, i.top, i.width, i.height);
            const ln = shape.Line, fc = ln.ForeColor, fill = shape.Fill, shadow = shape.Shadow;
            try {
              shape.Name = i.name; shadow.Visible = 0; ln.Visible = -1; fc.RGB = rgb(i.color); ln.Weight = i.weight;
              if (i.type === 'line') { ln.BeginArrowheadStyle = 1; ln.EndArrowheadStyle = 1; }
              if (i.fill) { fill.Solid(); const c = fill.ForeColor; c.RGB = rgb(i.fill); release(c); } else fill.Visible = 0;
            } finally { release(shadow, fill, fc, ln, shape); }
          }
        } finally { release(shapes, slide, slides); }
      });
      const raw = path.join(directory, 'private-raw.pptx');
      await stage('nativeSaveCloseMs', () => { pres.SaveAs(raw, 24); close(); });
      bytes = await fs.readFile(raw);
      // This intermediate can contain Office profile metadata; never publish it.
      await fs.unlink(raw);
    }
    report.output = path.join(directory, 'bicycle.pptx');
    await stage('finalizePackageMs', async () => {
      bytes = await sanitizeAndValidate(bytes, items);
      await fs.writeFile(report.output, bytes, { flag: 'wx' }); report.sha256 = hash(bytes); report.outputBytes = bytes.length;
    });
    report.authoringMs = report.stages.creationMs + (report.stages.nativeSaveCloseMs || 0) + report.stages.finalizePackageMs;
    if (!app) await startOffice();
    await stage('openFinalMs', () => { const docs = app.Presentations; try { pres = docs.Open(report.output, -1, 0, 0); } finally { release(docs); } });
    await stage('nativeReadbackMs', () => {
      const slides = pres.Slides, slide = slides.Item(1), shapes = slide.Shapes, setup = pres.PageSetup;
      try {
        assert.equal(Number(slides.Count), 1); assert.equal(Number(shapes.Count), items.length);
        assert.equal(Number(setup.SlideWidth), WIDTH); assert.equal(Number(setup.SlideHeight), HEIGHT);
        report.nativeObjects = [];
        for (const item of items) {
          const shape = shapes.Item(item.name), ln = shape.Line, color = ln.ForeColor, fill = shape.Fill;
          try {
            const geometry = { left: Number(shape.Left), top: Number(shape.Top), width: Number(shape.Width), height: Number(shape.Height) };
            for (const k of Object.keys(geometry)) assert(Math.abs(geometry[k] - item[k]) < 0.1, `${item.name}: ${k} mismatch (${geometry[k]} vs ${item[k]})`);
            assert.equal(Number(color.RGB), rgb(item.color), `${item.name}: line color`);
            assert(Math.abs(Number(ln.Weight) - item.weight) < 0.05, `${item.name}: line weight`);
            if (item.fill) { const c = fill.ForeColor; try { assert.equal(Number(c.RGB), rgb(item.fill)); } finally { release(c); } }
            else if (item.type !== 'line') assert.equal(Number(fill.Visible), 0);
            if (item.type === 'ellipse') assert.equal(Number(shape.AutoShapeType), 9);
            report.nativeObjects.push({ name: item.name, ...geometry, type: Number(shape.Type), rotation: Number(shape.Rotation), flipV: Number(shape.VerticalFlip), lineRGB: Number(color.RGB), weight: Number(ln.Weight) });
          } finally { release(fill, color, ln, shape); }
        }
      } finally { release(setup, shapes, slide, slides); }
    });
    report.preview = path.join(directory, 'bicycle.png');
    await stage('renderMs', () => { const slides = pres.Slides, slide = slides.Item(1); try { slide.Export(report.preview, 'PNG', 1600, 900); } finally { release(slide, slides); } });
    await stage('closeFinalMs', close);
    assert.equal(hash(await fs.readFile(report.output)), report.sha256, 'Read-only verification changed the output');
    report.verified = true;
  } catch (error) { report.error = { message: error.message, stack: error.stack }; }
  finally {
    try {
      await stage('shutdownMs', async () => {
        close();
        if (app) {
          const docs = app.Presentations, pv = app.ProtectedViewWindows;
          try { assert(sameProcess(report.officeIdentity)); assert.equal(Number(docs.Count), 0); assert.equal(Number(pv.Count), 0); assert.equal(Number(app.Visible), 0); app.Quit(); report.quitRequested = true; }
          finally { release(pv, docs); }
        }
        await new Promise(resolve => setImmediate(resolve)); globalThis.gc(); pumpMessages(); release(app); app = null; clearInterval(timer); uninitialize?.();
      });
    } catch (error) { report.cleanupError = error.message; }
    process.send?.({ finished: report });
  }
}

async function runChild(config, job) {
  const start = performance.now(); let report, officeIdentity, lastStage, diagnostics = '';
  const proc = fork(new URL(import.meta.url), ['--worker'], { windowsHide: true, execArgv: ['--expose-gc'], stdio: ['ignore', 'ignore', 'pipe', 'ipc'] });
  proc.stderr.on('data', data => { diagnostics = (diagnostics + data).slice(-4000); });
  proc.on('message', message => { if (message.finished) report = message.finished; if (message.officeIdentity) officeIdentity = message.officeIdentity; if (message.stage) lastStage = message.stage; });
  await once(proc, 'spawn'); const identity = processIdentity(proc.pid); job.add(identity);
  const timeout = setTimeout(() => terminateOwnedProcess(identity), 90000);
  proc.send(config); const [code] = await once(proc, 'exit'); clearTimeout(timeout);
  report ||= { method: config.method, trial: config.trial, officeIdentity, error: { message: `Worker failed at ${lastStage}`, diagnostics } };
  report.workerExitCode = code; report.workerExited = await waitForExit(identity, 0);
  report.officeExited = report.officeIdentity ? await waitForExit(report.officeIdentity, 15000) : null;
  report.endToEndMs = performance.now() - start;
  report.passed = report.verified === true && !report.error && !report.cleanupError && code === 0 && report.workerExited && report.officeExited;
  await writeJson(path.join(config.root, `${config.method}-${config.trial}`, 'report.json'), report);
  console.log(JSON.stringify({ method: report.method, trial: report.trial, passed: report.passed, authoringMs: report.authoringMs, endToEndMs: report.endToEndMs, error: report.error, cleanupError: report.cleanupError }));
  assert(report.passed, 'Benchmark stopped; inspect the trial report before retrying');
  return report;
}

if (process.argv[2] === '--worker') {
  process.once('message', async config => { try { await worker(config); } catch (error) { console.error(error); process.exitCode = 1; } finally { process.disconnect(); } });
} else {
  assert.equal(listProcesses('POWERPNT.EXE').length, 0, 'Benchmark requires no existing PowerPoint; user applications are preserved');
  const root = path.resolve('outputs', `bicycle-ab-${new Date().toISOString().replaceAll(':', '-')}-${randomUUID().slice(0, 6)}`);
  await fs.mkdir(root, { recursive: true });
  const items = bicycle();
  const summary = { root, started: new Date().toISOString(), node: process.version, shapeCount: items.length, trials: 3, workload: 'One new 960x540 slide; editable bicycle; shared geometry/colors/z-order; native readback and 1600x900 render on every trial', scope: 'Executor benchmark, not MCP/model latency; A uses formal buildDeck; B uses experimental COM shape creation; both sanitize metadata; fresh Office per trial; OS cache uncontrolled', reports: [] };
  await writeJson(path.join(root, 'workload.json'), items); await writeJson(path.join(root, 'summary.json'), summary);
  console.log(`Benchmark output: ${root}`);
  const job = createOwnedJob(), start = performance.now();
  try {
    for (let trial = 1; trial <= 3; trial++) for (const method of trial % 2 ? ['A', 'B'] : ['B', 'A']) {
      summary.reports.push(await runChild({ root, method, trial, items }, job));
      await writeJson(path.join(root, 'summary.json'), summary);
    }
    const median = a => [...a].sort((x, y) => x - y)[Math.floor(a.length / 2)];
    summary.statistics = ['A', 'B'].map(method => {
      const rows = summary.reports.filter(r => r.method === method);
      return { method, authoringMs: median(rows.map(r => r.authoringMs)), endToEndMs: median(rows.map(r => r.endToEndMs)), stages: Object.fromEntries([...new Set(rows.flatMap(r => Object.keys(r.stages)))].map(k => [k, median(rows.map(r => r.stages[k] || 0))])) };
    });
    summary.totalMs = performance.now() - start; summary.completed = new Date().toISOString();
    await writeJson(path.join(root, 'summary.json'), summary); console.log(JSON.stringify(summary.statistics, null, 2));
  } finally { job.close(); }
}
