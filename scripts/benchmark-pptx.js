import assert from 'node:assert/strict';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fork } from 'node:child_process';
import { once } from 'node:events';
import { performance } from 'node:perf_hooks';
import { randomUUID } from 'node:crypto';
import JSZip from 'jszip';
import { readPackage, writePackage, checkNativeSafe, parseXml, xml, NS, child, children, descendants, setChild, rels, resolvePart } from '../src/ooxml.js';
import { indexPackage, textBodyText } from '../src/file-engine.js';
import { hash, writeJson } from '../src/storage.js';
import { initializeSta, pumpMessages, listProcesses, processIdentity, sameProcess, waitForExit, createOwnedJob, terminateOwnedProcess, observeForeground } from '../src/windows.js';

const METHODS = ['com-runs', 'com-ranges', 'xml-repack', 'xml-reuse'];
const FONT = 'Microsoft YaHei', COLOR = '17365D', TITLE_COLOR = 'FFF2CC', BACKGROUND = 'EEF4FA', PREFIX = '※';
const norm = value => String(value).replaceAll('\r\n', '\n').replaceAll('\r', '\n').replaceAll('\v', '\n');
const rgb = hex => parseInt(hex.slice(0, 2), 16) + parseInt(hex.slice(2, 4), 16) * 256 + parseInt(hex.slice(4), 16) * 65536;
const fontMatches = value => ['microsoft yahei', '微软雅黑'].includes(String(value).toLowerCase());

function xmlShape(doc, target) {
  let container = descendants(doc, NS.p, 'spTree')[0];
  for (const id of target.groupPath) container = children(container, NS.p, 'grpSp').find(s => Number(descendants(s, NS.p, 'cNvPr')[0]?.getAttribute('id')) === id);
  const shape = children(container, NS.p).find(s => Number(descendants(s, NS.p, 'cNvPr')[0]?.getAttribute('id')) === target.shapeId);
  assert(shape, `Missing shape ${target.key}`); return shape;
}

function benchmarkIndex(parts) {
  const index = indexPackage(parts), seen = new Set();
  for (const slide of index.slides) {
    const layoutRel = rels(parts, slide.part).find(r => r.type.endsWith('/slideLayout'));
    const layout = resolvePart(slide.part, layoutRel.target);
    const masterRel = rels(parts, layout).find(r => r.type.endsWith('/slideMaster'));
    const part = resolvePart(layout, masterRel.target);
    if (seen.has(part)) continue; seen.add(part);
    const doc = parseXml(parts.get(part));
    const visit = (container, groupPath = []) => {
      for (const shape of children(container, NS.p)) {
        if (shape.localName === 'grpSp') { visit(shape, [...groupPath, Number(descendants(shape, NS.p, 'cNvPr')[0].getAttribute('id'))]); continue; }
        const body = child(shape, NS.p, 'txBody');
        if (!body || descendants(shape, NS.p, 'ph').length || !textBodyText(body).trim()) continue;
        const shapeId = Number(descendants(shape, NS.p, 'cNvPr')[0].getAttribute('id'));
        const transform = child(child(shape, NS.p, 'spPr'), NS.a, 'xfrm'), off = child(transform, NS.a, 'off'), ext = child(transform, NS.a, 'ext');
        const numeric = (node, name, scale = 12700) => node?.hasAttribute(name) ? Number(node.getAttribute(name)) / scale : null;
        const geometry = { left: numeric(off, 'x'), top: numeric(off, 'y'), width: numeric(ext, 'cx'), height: numeric(ext, 'cy'), rotation: numeric(transform, 'rot', 60000) ?? 0 };
        const runs = children(body, NS.a, 'p').flatMap(p => children(p, NS.a, 'r').map(r => {
          const prop = child(r, NS.a, 'rPr');
          return { text: child(r, NS.a, 't')?.textContent || '', style: { fontFace: child(prop, NS.a, 'latin')?.getAttribute('typeface'), fontSize: numeric(prop, 'sz', 100), color: child(child(prop, NS.a, 'solidFill'), NS.a, 'srgbClr')?.getAttribute('val') } };
        }));
        index.objects.push({ key: `master:${part}:${groupPath.join('.')}:${shapeId}`, kind: 'master', part, slide: slide.slide, slideId: slide.slideId, shapeId, groupPath, text: textBodyText(body), geometry, runs });
      }
    };
    visit(descendants(doc, NS.p, 'spTree')[0]);
  }
  return index;
}

function planFor(parts) {
  const index = benchmarkIndex(parts), docs = new Map();
  const targets = index.objects.filter(o => ['text', 'notes', 'master'].includes(o.kind) && o.text.trim()).map(o => {
    if (!docs.has(o.part)) docs.set(o.part, parseXml(parts.get(o.part)));
    const body = child(xmlShape(docs.get(o.part), o), NS.p, 'txBody');
    assert.equal(descendants(body, NS.a, 'fld').length, 0, 'Dynamic fields require a different workload');
    const paragraphs = children(body, NS.a, 'p').map(p => textBodyText({ childNodes: [p] }));
    // Generic test formatting; layout must be reviewed for each supplied deck.
    const title = o.kind === 'text' && o.groupPath.length === 0 && o.geometry.top < 50;
    const size = o.kind === 'notes' ? 10 : title ? 16 : 14;
    const color = title ? TITLE_COLOR : COLOR;
    return { ...o, paragraphs, expected: paragraphs.map(p => p.trim() ? PREFIX + p : p).join('\n'), size, color };
  });
  return { index, targets, font: FONT, color: COLOR, titleColor: TITLE_COLOR, background: BACKGROUND, prefix: PREFIX,
    counts: { slides: index.slides.length, textObjects: targets.length, masterTextObjects: targets.filter(o => o.kind === 'master').length, paragraphs: targets.reduce((n, t) => n + t.paragraphs.filter(p => p.trim()).length, 0), originalXmlRuns: targets.reduce((n, t) => n + t.runs.length, 0), images: index.objects.filter(o => o.kind === 'image').length, notesChanged: targets.filter(o => o.kind === 'notes').length } };
}

function styleXml(prop, size, color) {
  prop.setAttribute('sz', String(size * 100));
  for (const node of children(prop, NS.a)) if (['solidFill', 'noFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill'].includes(node.localName)) prop.removeChild(node);
  const fill = prop.ownerDocument.createElementNS(NS.a, 'a:solidFill');
  const next = children(prop, NS.a).find(n => ['effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'].includes(n.localName));
  prop.insertBefore(fill, next || null); setChild(fill, NS.a, 'a:srgbClr').setAttribute('val', color);
  for (const name of ['latin', 'ea']) {
    const old = child(prop, NS.a, name); if (old) prop.removeChild(old);
    const font = prop.ownerDocument.createElementNS(NS.a, `a:${name}`); font.setAttribute('typeface', FONT);
    const afterFont = children(prop, NS.a).find(n => ['cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'].includes(n.localName));
    prop.insertBefore(font, afterFont || null);
  }
}

function editXml(parts, plan) {
  const changed = new Map(), docs = new Map();
  const get = name => { if (!docs.has(name)) docs.set(name, parseXml(parts.get(name), name)); return docs.get(name); };
  for (const slide of plan.index.slides) {
    const doc = get(slide.part), common = child(doc.documentElement, NS.p, 'cSld');
    const old = child(common, NS.p, 'bg'); if (old) common.removeChild(old);
    const bg = doc.createElementNS(NS.p, 'p:bg'); common.insertBefore(bg, common.firstChild);
    const props = setChild(bg, NS.p, 'p:bgPr');
    setChild(setChild(props, NS.a, 'a:solidFill'), NS.a, 'a:srgbClr').setAttribute('val', BACKGROUND);
    setChild(props, NS.a, 'a:effectLst');
  }
  for (const target of plan.targets) {
    const body = child(xmlShape(get(target.part), target), NS.p, 'txBody'), bodyPr = child(body, NS.a, 'bodyPr');
    for (const node of children(bodyPr, NS.a)) if (['spAutoFit', 'normAutofit', 'noAutofit'].includes(node.localName)) bodyPr.removeChild(node);
    setChild(bodyPr, NS.a, 'a:noAutofit');
    children(body, NS.a, 'p').forEach((p, i) => {
      if (target.paragraphs[i].trim()) descendants(p, NS.a, 't')[0].textContent = PREFIX + descendants(p, NS.a, 't')[0].textContent;
      for (const run of children(p, NS.a, 'r')) {
        let props = child(run, NS.a, 'rPr');
        if (!props) { props = p.ownerDocument.createElementNS(NS.a, 'a:rPr'); run.insertBefore(props, run.firstChild); }
        styleXml(props, target.size, target.color);
      }
      styleXml(setChild(p, NS.a, 'a:endParaRPr'), target.size, target.color);
    });
  }
  for (const [name, doc] of docs) changed.set(name, Buffer.from(xml(doc)));
  return changed;
}

function verifyPackage(original, output, plan, method) {
  const index = benchmarkIndex(output), map = new Map(index.objects.map(o => [o.key, o]));
  assert.equal(index.slides.length, plan.counts.slides); assert.equal(index.objects.length, plan.index.objects.length);
  for (const target of plan.targets) {
    const actual = map.get(target.key); assert.equal(actual?.text, target.expected, `Text mismatch ${target.key}`);
    for (const run of actual.runs.filter(r => r.text.trim())) {
      assert(fontMatches(run.style.fontFace), `Font mismatch ${target.key}: ${run.style.fontFace}`);
      assert.equal(run.style.fontSize, target.size); assert.equal(run.style.color, target.color);
    }
  }
  for (const slide of index.slides) {
    const doc = parseXml(output.get(slide.part)), bg = child(child(doc.documentElement, NS.p, 'cSld'), NS.p, 'bg');
    assert.equal(descendants(bg, NS.a, 'srgbClr')[0]?.getAttribute('val'), BACKGROUND);
  }
  let maxGeometryDelta = 0;
  for (const old of plan.index.objects) {
    const actual = map.get(old.key); assert(actual);
    for (const key of ['left', 'top', 'width', 'height', 'rotation']) if (old.geometry[key] !== null && actual.geometry[key] !== null) maxGeometryDelta = Math.max(maxGeometryDelta, Math.abs(old.geometry[key] - actual.geometry[key]));
  }
  assert(maxGeometryDelta <= 0.1, `Geometry changed ${maxGeometryDelta} pt`);
  const mediaHashes = [...original].filter(([name]) => name.startsWith('ppt/media/')).map(([, data]) => hash(data));
  const outputMediaHashes = [...output].filter(([name]) => name.startsWith('ppt/media/')).map(([, data]) => hash(data));
  const backgroundHashes = new Set();
  for (const slide of plan.index.slides) {
    const doc = parseXml(original.get(slide.part)), bg = child(child(doc.documentElement, NS.p, 'cSld'), NS.p, 'bg');
    if (!bg) continue;
    const ids = Array.from(bg.getElementsByTagName('*')).map(n => n.getAttributeNS(NS.r, 'embed')).filter(Boolean);
    for (const rel of rels(original, slide.part)) if (ids.includes(rel.id)) backgroundHashes.add(hash(original.get(resolvePart(slide.part, rel.target))));
  }
  // Replacing a bitmap background legitimately removes that asset on Office
  // save. Office may also add PNG fallbacks for the source SVG pictures.
  const removedMedia = mediaHashes.filter(h => !outputMediaHashes.includes(h));
  assert(removedMedia.every(h => backgroundHashes.has(h)), 'A non-background media asset was lost or changed');
  let unchangedParts = 0;
  for (const [name, bytes] of original) if (output.get(name)?.equals(bytes)) unchangedParts++;
  if (method.startsWith('xml')) {
    const allowed = new Set([...plan.index.slides.map(s => s.part), ...plan.targets.map(t => t.part)]);
    for (const [name, bytes] of original) if (!allowed.has(name)) assert(output.get(name)?.equals(bytes), `Unexpected XML edit ${name}`);
  }
  return { passed: true, ...plan.counts, maxGeometryDelta, unchangedParts, originalParts: original.size, outputParts: output.size, mediaFilesPreserved: mediaHashes.length - removedMedia.length, replacedBackgroundAssetsRemoved: removedMedia.length, addedMediaAssets: outputMediaHashes.filter(h => !mediaHashes.includes(h)).length };
}

async function worker(config) {
  const { root, source, method, trial, visualFiles } = config;
  const directory = path.join(root, visualFiles ? 'visual' : `${method}-${trial}`); await fs.mkdir(directory, { recursive: true });
  const report = { method, trial, directory, node: process.version, started: new Date().toISOString(), stages: {}, counters: { fontAssignments: 0, textInsertions: 0, backgroundAssignments: 0 } };
  const start = performance.now(); let app, pres, winax, uninitialize, timer;
  const release = (...refs) => { if (winax) winax.release(...refs.filter(Boolean)); };
  const stage = async (name, fn) => { process.send?.({ stage: name }); const t = performance.now(); const result = await fn(); report.stages[name] = (report.stages[name] || 0) + performance.now() - t; return result; };
  const ensureApp = async () => {
    if (app) return;
    await stage('officeStartupMs', async () => {
      assert.equal(listProcesses('POWERPNT.EXE').length, 0, 'Benchmark requires no preexisting PowerPoint; it never closes user applications');
      uninitialize = initializeSta(); winax = createRequire(import.meta.url)('winax'); timer = setInterval(pumpMessages, 10);
      const activation = (BigInt(Date.now()) + 11644473600000n) * 10000n;
      app = new winax.Object('PowerPoint.Application', { activate: false });
      const processes = listProcesses('POWERPNT.EXE'); assert.equal(processes.length, 1);
      report.officeIdentity = processIdentity(processes[0].pid);
      process.send?.({ officeIdentity: report.officeIdentity });
      assert(BigInt(report.officeIdentity.created) >= activation);
      assert.equal(path.resolve(String(app.Path), 'POWERPNT.EXE').toLowerCase(), report.officeIdentity.image.toLowerCase());
      const docs = app.Presentations, pv = app.ProtectedViewWindows;
      try { assert.equal(Number(docs.Count), 0); assert.equal(Number(pv.Count), 0); assert.equal(Number(app.Visible), 0); } finally { release(pv, docs); }
    });
  };
  const open = file => { const docs = app.Presentations; try { pres = docs.Open(file, 0, 0, 0); assert.equal(path.resolve(String(pres.FullName)).toLowerCase(), path.resolve(file).toLowerCase()); } finally { release(docs); } };
  const close = () => { if (pres) { pres.Saved = -1; pres.Close(); release(pres); pres = null; } };
  const locate = target => {
    const refs = [], slides = pres.Slides; refs.push(slides);
    const slide = slides.FindBySlideID(target.slideId); refs.push(slide); let shapes;
    if (target.kind === 'notes') { const notes = slide.NotesPage; refs.push(notes); shapes = notes.Shapes; }
    else if (target.kind === 'master') { const master = slide.Master; refs.push(master); shapes = master.Shapes; }
    else shapes = slide.Shapes;
    refs.push(shapes);
    const byId = (collection, id) => { for (let i = 1; i <= Number(collection.Count); i++) { const s = collection.Item(i); if (Number(s.Id) === id) return s; release(s); } throw new Error(`Missing COM shape ${id}`); };
    for (const id of target.groupPath) { const group = byId(shapes, id); refs.push(group); shapes = group.GroupItems; refs.push(shapes); }
    const shape = byId(shapes, target.shapeId), frame = shape.TextFrame; let range = frame.TextRange; refs.push(shape, frame, range);
    return { shape, frame, get range() { return range; }, refresh() { release(range); range = frame.TextRange; refs[refs.length - 1] = range; }, dispose: () => release(...refs.reverse()) };
  };
  const setFont = (range, size, hex) => {
    const font = range.Font, color = font.Color;
    try { font.Name = FONT; font.NameFarEast = FONT; font.Size = size; color.RGB = rgb(hex); report.counters.fontAssignments += 4; }
    finally { release(color, font); }
  };
  try {
    if (visualFiles) {
      await ensureApp(); report.files = [];
      for (const [label, file] of Object.entries(visualFiles)) {
        open(file); const slides = pres.Slides, folder = path.join(directory, label); await fs.mkdir(folder, { recursive: true });
        await new Promise(resolve => setTimeout(resolve, 500)); pumpMessages();
        for (let i = 1; i <= Number(slides.Count); i++) { await new Promise(resolve => setTimeout(resolve, 50)); pumpMessages(); const slide = slides.Item(i); slide.Export(path.join(folder, `slide-${String(i).padStart(2, '0')}.png`), 'PNG', 960, 540); release(slide); }
        report.files.push({ label, file, slides: Number(slides.Count), folder }); release(slides); close();
      }
    } else {
      const bytes = await stage('inputReadMs', () => fs.readFile(source)); report.sourceHash = hash(bytes);
      const parts = await stage('preflightMs', () => readPackage(bytes)); checkNativeSafe(parts); const plan = planFor(parts); report.workload = plan.counts;
      const output = path.join(directory, 'edited-test.pptx'); report.output = output;
      if (method.startsWith('com')) {
        const working = path.join(directory, 'working.pptx'); await stage('copyMs', () => fs.copyFile(source, working));
        await ensureApp(); await stage('openForEditMs', () => open(working));
        await stage('editMs', async () => {
          const slides = pres.Slides;
          for (let i = 1; i <= Number(slides.Count); i++) {
            const slide = slides.Item(i); slide.FollowMasterBackground = 0; const bg = slide.Background, fill = bg.Fill, color = fill.ForeColor;
            fill.Solid(); color.RGB = rgb(BACKGROUND); fill.Transparency = 0; report.counters.backgroundAssignments += 4; release(color, fill, bg, slide);
          }
          release(slides);
          for (const target of plan.targets) {
            const t = locate(target);
            try {
              assert.equal(norm(t.range.Text), target.text, `Source text mismatch ${target.key}`); t.frame.AutoSize = 0;
              for (let p = target.paragraphs.length; p >= 1; p--) if (target.paragraphs[p - 1].trim()) { const paragraph = t.range.Paragraphs(p, 1), inserted = paragraph.InsertBefore(PREFIX); release(inserted, paragraph); report.counters.textInsertions++; }
              t.refresh();
              if (method === 'com-ranges') setFont(t.range, target.size, target.color);
              else {
                const allRuns = t.range.Runs(), count = Number(allRuns.Count), spans = []; release(allRuns);
                for (let r = 1; r <= count; r++) { const run = t.range.Runs(r, 1); spans.push({ start: Number(run.Start) - Number(t.range.Start) + 1, length: Number(run.Length) }); release(run); }
                for (const span of spans) { const run = t.range.Characters(span.start, span.length); setFont(run, target.size, target.color); release(run); }
              }
              assert.equal(norm(t.range.Text), target.expected);
            } finally { t.dispose(); }
          }
        });
        await stage('saveMs', () => pres.SaveCopyAs(output, 24)); await stage('closeForEditMs', close);
      } else {
        const changed = await stage('editMs', () => editXml(parts, plan)); report.changedXmlParts = changed.size;
        await stage('saveMs', async () => {
          let data;
          if (method === 'xml-repack') { const next = new Map(parts); for (const [name, part] of changed) next.set(name, part); data = await writePackage(next); }
          else {
            const zip = await JSZip.loadAsync(bytes, { createFolders: false });
            // Use STORE for media; reuse untouched compressed XML entries.
            // Recompressing PNG/JPEG dominates a naive whole-package write.
            for (const [name, part] of parts) if (name.startsWith('ppt/media/')) zip.file(name, part, { compression: 'STORE', date: zip.file(name).date });
            for (const [name, part] of changed) zip.file(name, part);
            data = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
          }
          await fs.writeFile(output, data);
        });
      }
      report.editPipelineMs = performance.now() - start;
      await stage('packageReadbackMs', async () => { const outParts = await readPackage(await fs.readFile(output)); report.packageValidation = verifyPackage(parts, outParts, plan, method); });
      await ensureApp(); await stage('openReadbackMs', () => open(output));
      await stage('nativeReadbackMs', () => {
        report.layoutFlags = []; const slides = pres.Slides;
        assert.equal(Number(slides.Count), plan.counts.slides);
        for (let i = 1; i <= Number(slides.Count); i++) { const slide = slides.Item(i), bg = slide.Background, fill = bg.Fill, color = fill.ForeColor; assert.equal(Number(color.RGB), rgb(BACKGROUND)); assert.equal(Number(fill.Type), 1); release(color, fill, bg, slide); }
        release(slides);
        for (const target of plan.targets) {
          const t = locate(target);
          try {
            assert.equal(norm(t.range.Text), target.expected, `Native text mismatch ${target.key}`);
            const font = t.range.Font, color = font.Color;
            assert(fontMatches(font.Name), `Native Latin font mismatch ${target.key}: ${String(font.Name)}`); assert(fontMatches(font.NameFarEast), `Native CJK font mismatch ${target.key}: ${String(font.NameFarEast)}`); assert.equal(Number(font.Size), target.size); assert.equal(Number(color.RGB), rgb(target.color)); release(color, font);
            const excess = Number(t.range.BoundHeight) - Number(t.shape.Height);
            if (excess > 2) report.layoutFlags.push({ slide: target.slide, shapeId: target.shapeId, excessHeightPt: excess });
          } finally { t.dispose(); }
        }
        report.nativeReadback = { passed: true, backgrounds: plan.counts.slides, textObjects: plan.counts.textObjects };
      });
      await stage('closeReadbackMs', close);
      assert.equal(hash(await fs.readFile(source)), report.sourceHash, 'Original source changed'); report.sourceUnchanged = true;
    }
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
    report.workerTotalMs = performance.now() - start;
    await writeJson(path.join(directory, 'report.json'), report); process.send?.({ finished: report });
  }
}

async function runChild(config, job) {
  const start = performance.now();
  // WinEvent callbacks belong in the parent: COM Release can dispatch Windows
  // messages during V8 GC, when JavaScript callbacks in the worker are forbidden.
  const observer = observeForeground();
  const proc = fork(new URL(import.meta.url), ['--worker'], { execArgv: ['--expose-gc'], windowsHide: true, stdio: ['ignore', 'ignore', 'pipe', 'ipc'] });
  let report, lastStage, officeIdentity, diagnostics = '';
  proc.stderr.on('data', chunk => { diagnostics += chunk; });
  proc.on('message', message => { if (message.stage) lastStage = message.stage; if (message.officeIdentity) officeIdentity = message.officeIdentity; if (message.finished) report = message.finished; });
  await once(proc, 'spawn'); const identity = processIdentity(proc.pid); job.add(identity);
  const timeout = setTimeout(() => { terminateOwnedProcess(identity); }, 180000);
  proc.send(config); const [code] = await once(proc, 'exit'); clearTimeout(timeout);
  if (!report) {
    const directory = path.join(config.root, config.visualFiles ? 'visual' : `${config.method}-${config.trial}`);
    report = { directory, method: config.method, trial: config.trial, officeIdentity, error: { message: `Worker failed at ${lastStage}`, diagnostics } };
  }
  report.workerExitCode = code; report.workerExited = await waitForExit(identity, 0);
  report.officeExited = report.officeIdentity ? await waitForExit(report.officeIdentity, 15000) : null;
  report.foreground = observer.stop();
  report.endToEndMs = performance.now() - start;
  report.passed = !report.error && !report.cleanupError && code === 0 && report.workerExited && report.officeExited === true;
  await writeJson(path.join(report.directory, 'report.json'), report);
  console.log(JSON.stringify({ method: config.method, trial: config.trial, passed: report.passed, editPipelineMs: report.editPipelineMs, endToEndMs: report.endToEndMs, error: report.error?.message, root: report.directory }));
  assert(report.passed, `Benchmark failed; see ${report.directory}/report.json`); return report;
}

if (process.argv[2] === '--worker') {
  process.once('message', async config => { try { await worker(config); } catch (error) { console.error(error); process.exitCode = 1; } finally { process.disconnect(); } });
} else {
  const args = process.argv.slice(2), value = (name, fallback) => args.includes(name) ? args[args.indexOf(name) + 1] : fallback;
  const sourceArgument = value('--source');
  assert(sourceArgument && !sourceArgument.startsWith('--'), 'An explicit --source path to an authorized sample is required');
  const source = path.resolve(sourceArgument);
  const root = path.resolve(value('--out', `outputs/pptx-benchmark-${new Date().toISOString().replaceAll(':', '-')}-${randomUUID().slice(0, 6)}`));
  const trials = Number(value('--trials', 3)); assert(Number.isInteger(trials) && trials >= 1 && trials <= 5);
  const methods = value('--methods', METHODS.join(',')).split(','); assert(methods.length && methods.every(m => METHODS.includes(m)) && new Set(methods).size === methods.length);
  assert.equal(listProcesses('POWERPNT.EXE').length, 0, 'Close PowerPoint before benchmarking; existing applications are never closed by setup');
  await fs.mkdir(root, { recursive: true }); const bytes = await fs.readFile(source), plan = planFor(await readPackage(bytes));
  const summary = { source, sourceBytes: bytes.length, sourceHash: hash(bytes), root, started: new Date().toISOString(), node: process.version, trials, workload: { ...plan.counts, font: FONT, color: COLOR, titleColor: TITLE_COLOR, background: BACKGROUND, prefix: PREFIX }, reports: [] };
  await writeJson(path.join(root, 'workload.json'), plan); await writeJson(path.join(root, 'summary.json'), summary); console.log(`Benchmark output: ${root}`);
  const job = createOwnedJob(), start = performance.now();
  try {
    for (let trial = 1; trial <= trials; trial++) {
      const offset = (trial - 1) % methods.length, order = [...methods.slice(offset), ...methods.slice(0, offset)];
      for (const method of order) { summary.reports.push(await runChild({ source, root, method, trial }, job)); await writeJson(path.join(root, 'summary.json'), summary); }
    }
    summary.benchmarkMs = performance.now() - start;
    const median = values => { const a = [...values].sort((a, b) => a - b); return a.length % 2 ? a[(a.length - 1) / 2] : (a[a.length / 2 - 1] + a[a.length / 2]) / 2; };
    summary.statistics = methods.map(method => {
      const rows = summary.reports.filter(r => r.method === method), metrics = {};
      for (const metric of ['editPipelineMs', 'endToEndMs']) metrics[metric] = { median: median(rows.map(r => r[metric])), min: Math.min(...rows.map(r => r[metric])), max: Math.max(...rows.map(r => r[metric])) };
      metrics.stagesMedian = Object.fromEntries([...new Set(rows.flatMap(r => Object.keys(r.stages)))].map(key => [key, median(rows.map(r => r.stages[key] || 0))]));
      return { method, ...metrics, fontAssignments: rows[0].counters.fontAssignments, layoutFlags: rows[0].layoutFlags };
    });
    const visualFiles = { original: source };
    for (const method of methods) { const selected = summary.reports.find(r => r.method === method); const destination = path.join(root, `${method}-测试编辑版.pptx`); await fs.copyFile(selected.output, destination); visualFiles[method] = destination; }
    await writeJson(path.join(root, 'summary.json'), summary);
    if (!args.includes('--no-render')) summary.visual = await runChild({ source, root, visualFiles }, job);
    summary.sourceUnchanged = hash(await fs.readFile(source)) === summary.sourceHash; assert(summary.sourceUnchanged);
    summary.totalMs = performance.now() - start; summary.finished = new Date().toISOString();
    await writeJson(path.join(root, 'summary.json'), summary); console.log(JSON.stringify(summary.statistics, null, 2));
  } finally { job.close(); }
}
