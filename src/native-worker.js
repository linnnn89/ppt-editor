import { createRequire } from 'node:module';
import { promises as fs } from 'node:fs';
import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { setTimeout as delay } from 'node:timers/promises';
import { check, PptError, errorResult } from './errors.js';
import { hash, stableJson, publishNew, prepareNewOutput } from './storage.js';
import { readPackage, checkNativeSafe, validatePackage } from './ooxml.js';
import { initializeSta, acquireNativeMutex, listProcesses, processIdentity, sameProcess, windowProcessId, pumpMessages } from './windows.js';
import { backgroundPatterns, readBackgroundImage } from './background.js';

const require = createRequire(import.meta.url);
let winax, app, appLease, releaseSta, releaseMutex, pump, started = false, stopping = false;
const documents = new Map();
let workerOwnerKey;
const normalize = text => String(text || '').replaceAll('\r\n', '\n').replaceAll('\r', '\n').replaceAll('\v', '\n');
const release = (...refs) => { try { winax?.release(...refs.filter(Boolean)); } catch {} };
const stateHash = snapshot => hash(stableJson(snapshot.objects.map(o => [o.key, o.fingerprint])));
const rgbNumber = value => { const [r, g, b] = value.match(/../g).map(s => parseInt(s, 16)); return r | g << 8 | b << 16; };
const rgbHex = value => (Number(value) >>> 0).toString(16).padStart(6, '0').match(/../g).reverse().join('').toUpperCase();

function backgroundSnapshot(slide) {
  const bg = slide.Background, fill = bg.Fill;
  const color = name => { const c = fill[name]; try { return rgbHex(c.RGB); } finally { release(c); } };
  try {
    const type = ({ 1: 'solid', 2: 'pattern', 3: 'gradient', 4: 'texture', 5: 'inherit', 6: 'image' })[Number(fill.Type)] || 'other';
    if (Number(slide.FollowMasterBackground) === -1) return { type: 'inherit', effectiveType: type, color: type === 'solid' ? color('ForeColor') : null };
    const result = { type, transparency: Number(fill.Transparency) * 100 };
    if (type === 'solid') result.color = color('ForeColor');
    if (type === 'pattern') Object.assign(result, { pattern: Object.entries(backgroundPatterns).find(([, pair]) => pair[1] === Number(fill.Pattern))?.[0] || Number(fill.Pattern), foreground: color('ForeColor'), background: color('BackColor') });
    if (type === 'gradient') {
      // Office does not expose a linear angle for every existing gradient style.
      try { result.angle = Number(fill.GradientAngle); } catch { result.angle = null; }
      result.stops = [];
      const stops = fill.GradientStops;
      try {
        for (let i = 1; i <= Number(stops.Count); i++) {
          const stop = stops.Item(i), c = stop.Color;
          try { result.stops.push({ position: Number(stop.Position), color: rgbHex(c.RGB), transparency: Number(stop.Transparency) * 100 }); }
          finally { release(c, stop); }
        }
      } finally { release(stops); }
    }
    return result;
  } finally { release(fill, bg); }
}

async function invoke(callback) {
  const deadline = Date.now() + 2000;
  let attempt = 0;
  while (true) {
    try { return callback(); }
    catch (error) {
      const diagnostic = `${error.message || ''} ${((error.number || 0) >>> 0).toString(16)}`;
      if (!/80010001|8001010a/i.test(diagnostic) || Date.now() >= deadline) throw error;
      await delay(Math.min(30 * 2 ** attempt++, 200)); pumpMessages();
    }
  }
}

function ensureApplication() {
  if (app) return;
  releaseMutex = acquireNativeMutex();
  const before = new Set(listProcesses('POWERPNT.EXE').map(p => p.pid));
  try {
    const activationStarted = (BigInt(Date.now()) + 11644473600000n) * 10000n;
    app = new winax.Object('PowerPoint.Application', { activate: false });
    let pid = 0;
    try { const hwnd = Number(app.HWND); if (Number.isFinite(hwnd) && hwnd) pid = windowProcessId(hwnd); } catch {}
    const after = listProcesses('POWERPNT.EXE');
    const observedNewProcesses = after.filter(p => !before.has(p.pid)).map(p => processIdentity(p.pid)).filter(Boolean);
    const presentations = app.Presentations, protectedViews = app.ProtectedViewWindows;
    let preexistingDocumentCount, protectedViewCount;
    try { preexistingDocumentCount = Number(presentations.Count); protectedViewCount = Number(protectedViews.Count); }
    finally { release(protectedViews, presentations); }
    let identity = pid ? processIdentity(pid) : null;
    let identityMethod = identity ? 'application-window' : null;
    // Hidden Office builds may not expose HWND through IDispatch. A cold,
    // exclusive activation also needs time, executable and empty-app evidence.
    // This grants only guarded COM Quit, never process termination authority.
    const candidate = observedNewProcesses[0];
    if (!identity && before.size === 0 && after.length === 1 && observedNewProcesses.length === 1 &&
        BigInt(candidate.created) >= activationStarted && Number(app.Visible) === 0 &&
        preexistingDocumentCount === 0 && protectedViewCount === 0 &&
        path.resolve(candidate.image).toLowerCase() === path.resolve(String(app.Path), 'POWERPNT.EXE').toLowerCase()) {
      identity = candidate; identityMethod = 'exclusive-hidden-activation';
    }
    appLease = { identity, identityMethod, observedNewProcesses, authority: identity ? 'identified-application' : 'document-only', createdByTask: Boolean(identity && !before.has(identity.pid)), preexistingDocumentCount, protectedViewCount, version: String(app.Version), build: String(app.Build), quitRequested: false };
    process.send({ stage: 'application-ready', lease: appLease });
  } catch (error) { release(app); app = null; releaseMutex?.(); releaseMutex = null; throw error; }
}

function fontSnapshot(font) {
  const tri = value => Number(value) === -2 ? null : Number(value) === -1;
  let color = null;
  try { color = (Number(font.Color.RGB) >>> 0).toString(16).padStart(6, '0').match(/../g).reverse().join('').toUpperCase(); } catch {}
  return { underline: tri(font.Underline), fontFaceFarEast: String(font.NameFarEast || '') || null, fontSize: Number(font.Size) || null, fontFace: String(font.Name || '') || null, bold: tri(font.Bold), italic: tri(font.Italic), color, source: 'native-effective-properties' };
}

function smartArtSnapshot(shape) {
  const art = shape.SmartArt, nodes = art.AllNodes, layout = art.Layout, result = { layoutId: String(layout.Id), nodes: [] };
  try {
    for (let i = 1; i <= Number(nodes.Count); i++) {
      const node = nodes.Item(i), frame = node.TextFrame2, range = frame.TextRange;
      try { result.nodes.push({ nodeIndex: i, level: Number(node.Level), text: normalize(range.Text) }); }
      finally { release(range, frame, node); }
    }
    return result;
  } finally { release(layout, nodes, art); }
}

function chartSnapshot(shape) {
  const chart = shape.Chart, series = chart.SeriesCollection(), result = { series: [] };
  try {
    for (let i = 1; i <= Number(series.Count); i++) {
      const item = series.Item(i);
      try { result.series.push({ name: String(item.Name), chartType: Number(item.ChartType), axisGroup: Number(item.AxisGroup) }); }
      finally { release(item); }
    }
    return result;
  } finally { release(series, chart); }
}

function inspectDocument(doc) {
  const objects = [], slides = [];
  const size = doc.pres.PageSetup;
  const width = Number(size.SlideWidth), height = Number(size.SlideHeight); release(size);
  const slideCollection = doc.pres.Slides;
  for (let slideIndex = 1; slideIndex <= Number(slideCollection.Count); slideIndex++) {
    const slide = slideCollection.Item(slideIndex), slideId = Number(slide.SlideID);
    const page = { slide: slideIndex, slideId, objectCount: 0 }; slides.push(page);
    // 索引幻灯片背景对象，读取有效背景颜色
    let bgColor = null;
    try {
      const bg = slide.Background, fill = bg.Fill;
      const rawRgb = Number(fill.ForeColor.RGB);
      bgColor = (rawRgb >>> 0).toString(16).padStart(6, '0').match(/../g).reverse().join('').toUpperCase();
      release(fill, bg);
    } catch {}
    const bgItem = {
      key: `background:${slideId}::0`,
      slide: slideIndex,
      slideId,
      shapeId: 0,
      groupPath: [],
      name: 'Background',
      kind: 'background',
      color: bgColor,
      fill: backgroundSnapshot(slide),
      text: '',
      runs: [],
      capabilities: ['set_slide_background']
    };
    bgItem.fingerprint = hash(stableJson(bgItem));
    objects.push(bgItem);
    page.objectCount++;
    function visit(shapes, groupPath = [], notes = false) {
      for (let i = 1; i <= Number(shapes.Count); i++) {
        const shape = shapes.Item(i);
        let textFrame, range, table;
        try {
          const shapeId = Number(shape.Id), type = Number(shape.Type);
          let placeholder = null;
          if (type === 14) { try { placeholder = Number(shape.PlaceholderFormat.Type); } catch {} }
          if (notes && placeholder !== 2) continue;
          const hasText = Number(shape.HasTextFrame) === -1, hasTable = Number(shape.HasTable) === -1;
          const hasSmartArt = Number(shape.HasSmartArt) === -1, hasChart = Number(shape.HasChart) === -1;
          const kind = notes ? 'notes' : hasSmartArt ? 'smartart' : hasChart ? 'chart' : hasTable ? 'table' : type === 6 ? 'group' : hasText ? 'text' : [11, 13].includes(type) ? 'image' : 'shape';
          const key = `${notes ? 'notes' : 'slide'}:${slideId}:${groupPath.join('.')}:${shapeId}`;
          let text = '', runs = [], rows;
          if (hasText && !hasTable) {
            textFrame = shape.TextFrame; range = textFrame.TextRange; text = normalize(range.Text);
            if (text) {
              const allRuns = range.Runs(), count = Number(allRuns.Count); release(allRuns);
              let offset = 0;
              for (let r = 1; r <= count; r++) {
                const run = range.Runs(r, 1), font = run.Font;
                const runText = normalize(run.Text);
                runs.push({ text: runText, start: offset, length: runText.length, style: fontSnapshot(font) }); offset += runText.length;
                release(font, run);
              }
            }
          }
          if (hasTable) {
            table = shape.Table; rows = [];
            const rowCount = Number(table.Rows.Count), columnCount = Number(table.Columns.Count);
            for (let row = 1; row <= rowCount; row++) {
              const cells = [];
              for (let column = 1; column <= columnCount; column++) {
                const cell = table.Cell(row, column), cellShape = cell.Shape, frame = cellShape.TextFrame, cellRange = frame.TextRange;
                cells.push(normalize(cellRange.Text)); release(cellRange, frame, cellShape, cell);
              }
              rows.push(cells);
            }
            text = rows.map(row => row.join('\t')).join('\n');
          }
          const item = { key, slide: slideIndex, slideId, shapeId, groupPath, name: String(shape.Name), kind, text, runs, placeholder,
            ...(hasSmartArt ? { smartart: smartArtSnapshot(shape) } : {}), ...(hasChart ? { chart: chartSnapshot(shape) } : {}),
            geometry: { left: Number(shape.Left), top: Number(shape.Top), width: Number(shape.Width), height: Number(shape.Height), rotation: Number(shape.Rotation), unit: 'pt', source: 'native-effective' },
            ...(rows ? { rows, rowCount: rows.length, columnCount: rows[0]?.length || 0 } : {}) };
          item.capabilities = [ ...(hasSmartArt ? ['set_smartart_text'] : []), ...(kind === 'text' && !groupPath.length ? ['convert_to_smartart'] : []), ...(hasText && !hasTable && !hasSmartArt ? ['replace_text', 'set_style'] : []), ...(hasTable ? ['set_table_cell'] : []), ...(!notes && groupPath.length === 0 && kind !== 'group' ? ['set_geometry'] : []) ];
          item.fingerprint = hash(stableJson(item));
          // Diagnostic text bounds are kept out of the mutation fingerprint:
          // Office may round layout measurements between read-only queries.
          if (range && text && kind === 'text' && !groupPath.length && Math.abs(item.geometry.rotation) < 0.01) {
            try {
              const measured = { left: Number(range.BoundLeft), top: Number(range.BoundTop), width: Number(range.BoundWidth), height: Number(range.BoundHeight) };
              if (Object.values(measured).every(Number.isFinite) && measured.width >= 0 && measured.height >= 0) item.textBounds = measured;
            } catch { /* Keep unavailable text measurements explicitly uncovered. */ }
          }
          objects.push(item); page.objectCount++;
          if (type === 6) { const group = shape.GroupItems; visit(group, [...groupPath, shapeId], notes); release(group); }
        } finally { release(table, range, textFrame, shape); }
      }
    }
    const shapes = slide.Shapes; visit(shapes); release(shapes);
    const notesPage = slide.NotesPage, noteShapes = notesPage.Shapes; visit(noteShapes, [], true); release(noteShapes, notesPage, slide);
  }
  release(slideCollection);
  return { width, height, slides, objects };
}

function getDocument(id) { const doc = documents.get(id); check(doc && !doc.unusable, 'DOCUMENT_UNAVAILABLE', 'Native document is unavailable.'); return doc; }
function guard(doc) {
  check(path.resolve(String(doc.pres.FullName)).toLowerCase() === doc.openedPath.toLowerCase(), 'DOCUMENT_IDENTITY_MISMATCH', 'Native document path no longer matches the task binding.');
  if (appLease.identity) check(sameProcess(appLease.identity), 'OFFICE_PROCESS_MISMATCH', 'The bound PowerPoint process identity changed.');
  const snapshot = inspectDocument(doc);
  check(!doc.stateHash || stateHash(snapshot) === doc.stateHash, 'EXTERNAL_CHANGE', 'PowerPoint content changed outside the current revision; preserve a recovery copy before continuing.');
  return snapshot;
}

async function checkpoint(doc, destination) {
  if (path.resolve(destination).toLowerCase() !== path.resolve(String(doc.pres.FullName)).toLowerCase()) {
    await prepareNewOutput(destination);
    await invoke(() => doc.pres.SaveCopyAs(destination, 24));
  }
  const bytes = await fs.readFile(destination); await readPackage(bytes);
  return { path: destination, sha256: hash(bytes) };
}

function locate(doc, target) {
  let slideId = target.slideId, shapeId = target.shapeId, groupPath = target.groupPath || [], kind = target.kind;
  if ((!slideId || !shapeId) && target.key) {
    const parts = target.key.split(':');
    kind = parts[0];
    slideId = Number(parts[1]);
    groupPath = parts[2] ? parts[2].split('.').map(Number) : [];
    shapeId = Number(parts[3]);
  }
  const slide = doc.pres.Slides.FindBySlideID(slideId);
  check(slide, 'TARGET_NOT_FOUND', 'Slide not found.');
  if (kind === 'background') {
    return { slide, shape: null, dispose: () => release(slide) };
  }
  const references = [slide];
  let shapes;
  if (kind === 'notes') { const notes = slide.NotesPage; references.push(notes); shapes = notes.Shapes; }
  else shapes = slide.Shapes;
  references.push(shapes);
  function find(collection, id) {
    for (let i = 1; i <= Number(collection.Count); i++) { const candidate = collection.Item(i); if (Number(candidate.Id) === id) return candidate; release(candidate); }
    throw new PptError('TARGET_NOT_FOUND', 'Native shape no longer exists.');
  }
  for (const id of groupPath) { const group = find(shapes, id); references.push(group); shapes = group.GroupItems; references.push(shapes); }
  const shape = find(shapes, shapeId); references.push(shape);
  return { shape, dispose: () => release(...references.reverse()) };
}

function matchPositions(text, search) {
  const positions = []; let offset = 0;
  while ((offset = text.indexOf(search, offset)) !== -1) { positions.push(offset); offset += search.length; }
  return positions;
}

function preflight(snapshot, operations) {
  const baseline = new Map(snapshot.objects.map(o => [o.key, o])), projected = new Map(snapshot.objects.map(o => [o.key, o.text]));
  for (const op of operations) {
    const target = baseline.get(op.target.key);
    check(target && target.fingerprint === op.target.fingerprint, 'TARGET_CHANGED', 'Native target changed since inspection.');
    check(target.capabilities.includes(op.type), 'OPERATION_UNSUPPORTED', 'Operation is not supported on this native target.');
    if (op.type === 'convert_to_smartart') {
      check(target.text.trim() && target.text.split('\n').length <= 30, 'SMARTART_CONTENT_INVALID', 'SmartArt needs 1–30 text paragraphs.');
      check(operations.filter(o => o.target.key === op.target.key).length === 1, 'SMARTART_CONVERSION_BATCH', 'Convert a shape in its own batch before inspecting new node references.');
    }
    if (op.type === 'set_smartart_text') check(op.nodeIndex <= target.smartart.nodes.length, 'SMARTART_NODE_NOT_FOUND', 'SmartArt node index is out of range.');
    if (op.type === 'replace_text') {
      const text = projected.get(target.key), positions = matchPositions(text, op.search);
      check(positions.length === op.expectedMatches, 'MATCH_COUNT_MISMATCH', 'Native text match count mismatch.', { expected: op.expectedMatches, actual: positions.length });
      if (op.crossRunPolicy === 'reject') {
        check(text === target.text || target.runs.length <= 1, 'CROSS_RUN_REPLACEMENT', 'Repeated edits on a multi-run object require first_run policy.');
        for (const start of positions) check(target.runs.some(r => start >= r.start && start + op.search.length <= r.start + r.length), 'CROSS_RUN_REPLACEMENT', 'Native text spans multiple runs; explicitly select first_run policy.');
      }
      projected.set(target.key, text.split(op.search).join(op.replacement));
    }
    if (op.type === 'set_table_cell') check(op.row <= target.rowCount && op.column <= target.columnCount, 'CELL_NOT_FOUND', 'Native table cell is out of range.');
  }
}

function setFont(font, style) {
  if (style.fontSize !== undefined && style.fontSize !== null) font.Size = style.fontSize;
  if (style.fontFace) { font.Name = style.fontFace; font.NameFarEast = style.fontFaceFarEast || style.fontFace; }
  if (style.bold !== undefined && style.bold !== null) font.Bold = style.bold ? -1 : 0;
  if (style.italic !== undefined && style.italic !== null) font.Italic = style.italic ? -1 : 0;
  if (style.underline !== undefined && style.underline !== null) font.Underline = style.underline ? -1 : 0;
  if (style.color) { const rgb = style.color.match(/../g).map(s => parseInt(s, 16)); font.Color.RGB = rgb[0] | rgb[1] << 8 | rgb[2] << 16; }
}

async function applyOperation(doc, op) {
  const target = locate(doc, op.target), refs = [];
  try {
    if (op.type === 'convert_to_smartart') {
      const layouts = app.SmartArtLayouts; refs.push(layouts);
      const pattern = { process: /\/process1(?:#|$)/i, cycle: /\/cycle2(?:#|$)/i, hierarchy: /\/orgChart1(?:#|$)/i }[op.layout];
      let selected;
      for (let i = 1; layouts && i <= Number(layouts.Count); i++) { const layout = layouts.Item(i); if (pattern.test(String(layout.Id))) { selected = layout; break; } release(layout); }
      // Some Office installations expose SmartArt but return no global catalog.
      // Reuse a layout from an explicitly opened document in this same task.
      if (!selected) for (const existing of documents.values()) {
        guard(existing);
        const slides = existing.pres.Slides;
        try { for (let s = 1; !selected && s <= Number(slides.Count); s++) {
          const slide = slides.Item(s), shapes = slide.Shapes;
          try { for (let i = 1; i <= Number(shapes.Count); i++) {
            const shape = shapes.Item(i); let art, layout;
            try { if (Number(shape.HasSmartArt) === -1) { art = shape.SmartArt; layout = art.Layout; if (pattern.test(String(layout.Id))) { selected = layout; layout = null; break; } } }
            finally { release(layout, art, shape); }
          } } finally { release(shapes, slide); }
        } } finally { release(slides); }
        if (selected) break;
      }
      check(selected, 'SMARTART_LAYOUT_UNAVAILABLE', 'Requested built-in SmartArt layout is unavailable in this Office installation.');
      refs.push(selected); await invoke(() => target.shape.ConvertTextToSmartArt(selected)); return;
    }
    if (op.type === 'set_smartart_text') {
      const art = target.shape.SmartArt, nodes = art.AllNodes, node = nodes.Item(op.nodeIndex), frame = node.TextFrame2, range = frame.TextRange;
      refs.push(art, nodes, node, frame, range);
      const before = normalize(range.Text), replacement = op.text.replaceAll('\n', '\r');
      try { await invoke(() => { range.Text = replacement; }); }
      catch (error) {
        // Hidden PowerPoint can expose a readable node range whose setter returns
        // E_FAIL. Its associated shape remains editable; never guess among shapes.
        if (!/E_FAIL/.test(error.message) || normalize(range.Text) !== before) throw error;
        const shapes = node.Shapes, candidates = []; refs.push(shapes);
        for (let i = 1; i <= Number(shapes.Count); i++) {
          const shape = shapes.Item(i); refs.push(shape);
          const shapeFrame = shape.TextFrame2, shapeRange = shapeFrame.TextRange;
          refs.push(shapeFrame, shapeRange);
          if (normalize(shapeRange.Text) === before) candidates.push(shapeRange);
        }
        check(candidates.length === 1, 'SMARTART_TEXT_TARGET_AMBIGUOUS', 'Cannot uniquely identify the text shape belonging to this SmartArt node.');
        await invoke(() => { candidates[0].Text = replacement; });
      }
      check(normalize(range.Text) === normalize(replacement), 'SMARTART_TEXT_READBACK_FAILED', 'SmartArt node text did not match after writing.');
      return;
    }
    if (op.type === 'set_slide_background') {
      const slide = target.slide;
      const background = op.fill || { type: 'solid', color: op.color };
      if (background.type === 'inherit') { await invoke(() => { slide.FollowMasterBackground = -1; }); return; }
      if (['image', 'texture'].includes(background.type)) {
        const asset = await readBackgroundImage(background.path);
        check(!op.imageHash || asset.sha256 === op.imageHash, 'IMAGE_CHANGED', 'Prepared background image changed.');
      }
      await invoke(() => {
        // 显式脱离母版背景继承，以独立纯色填充幻灯片背景
        slide.FollowMasterBackground = 0; // msoFalse
        const bg = slide.Background, fill = bg.Fill;
        refs.push(fill, bg);
        if (background.type === 'solid') { fill.Solid(); fill.ForeColor.RGB = rgbNumber(background.color); }
        else if (background.type === 'gradient') {
          fill.TwoColorGradient(1, 1); fill.ForeColor.RGB = rgbNumber(background.startColor); fill.BackColor.RGB = rgbNumber(background.endColor); fill.GradientAngle = background.angle;
        } else if (background.type === 'pattern') {
          fill.Patterned(backgroundPatterns[background.pattern][1]); fill.ForeColor.RGB = rgbNumber(background.foreground); fill.BackColor.RGB = rgbNumber(background.background);
        } else if (background.type === 'image') fill.UserPicture(background.path);
        else if (background.type === 'texture') { fill.UserTextured(background.path); fill.TextureAlignment = 0; fill.TextureHorizontalScale = 1; fill.TextureVerticalScale = 1; fill.TextureOffsetX = 0; fill.TextureOffsetY = 0; }
        if (background.transparency !== undefined) fill.Transparency = background.transparency / 100;
      });
      return;
    }
    if (op.type === 'set_geometry') for (const [property, value] of Object.entries(op.geometry)) await invoke(() => { target.shape[property[0].toUpperCase() + property.slice(1)] = value; });
    if (op.type === 'replace_text' || op.type === 'set_style') {
      const frame = target.shape.TextFrame, range = frame.TextRange; refs.push(frame, range);
      if (op.type === 'set_style') { const font = range.Font; refs.push(font); await invoke(() => setFont(font, op.style)); }
      else {
        const positions = matchPositions(normalize(range.Text), op.search);
        check(positions.length === op.expectedMatches, 'MATCH_COUNT_MISMATCH', 'Text changed before native write.');
        for (const start of positions.reverse()) {
          const replacementRange = range.Characters(start + 1, op.search.length);
          const first = range.Characters(start + 1, 1), firstFont = first.Font;
          const style = fontSnapshot(firstFont); release(firstFont, first);
          await invoke(() => { replacementRange.Text = op.replacement; }); release(replacementRange);
          if (op.crossRunPolicy === 'first_run' && op.replacement.length) {
            const inserted = range.Characters(start + 1, op.replacement.length), font = inserted.Font;
            await invoke(() => setFont(font, style)); release(font, inserted);
          }
        }
      }
    }
    if (op.type === 'set_table_cell') {
      const table = target.shape.Table, cell = table.Cell(op.row, op.column), cellShape = cell.Shape, frame = cellShape.TextFrame, range = frame.TextRange;
      refs.push(table, cell, cellShape, frame, range);
      const original = String(range.Text) ? range.Characters(1, 1) : range;
      const font = original.Font, style = fontSnapshot(font); release(font); if (original !== range) release(original);
      await invoke(() => { range.Text = op.text.replaceAll('\n', '\r'); });
      const newFont = range.Font; await invoke(() => setFont(newFont, style)); release(newFont);
    }
  } finally { release(...refs.reverse()); target.dispose(); }
}

async function closeDocument(doc, { interrupted = false } = {}) {
  if (!doc.pres) return;
  if (!doc.readOnly) {
    if (interrupted) doc.checkpoint = await checkpoint(doc, path.join(doc.directory, `interrupted-${Date.now()}.pptx`));
    else guard(doc);
  }
  const bytes = await fs.readFile(doc.checkpoint.path);
  check(hash(bytes) === doc.checkpoint.sha256, 'CHECKPOINT_CORRUPT', 'Cannot discard task memory without a verified checkpoint.');
  await readPackage(bytes);
  // This is exclusively a task copy whose content is in the verified checkpoint.
  // Saved=true is used only to discard that duplicate in-memory state on close.
  if (!doc.readOnly) doc.pres.Saved = -1;
  doc.pres.Close(); release(doc.pres); doc.pres = null; documents.delete(doc.id);
}

async function validateFile(args) {
  const bytes = await fs.readFile(args.path), sha256 = hash(bytes);
  check(sha256 === args.sha256, 'CHECKPOINT_CORRUPT', 'Readback candidate bytes mismatch.');
  const parts = await readPackage(bytes); await checkNativeSafe(parts);
  const structural = validatePackage(parts);
  ensureApplication();
  const { pres } = await invoke(() => ({ pres: app.Presentations.Open(args.path, -1, 0, 0) }));
  const doc = { id: randomUUID(), pres, readOnly: true, checkpoint: { path: args.path, sha256 } };
  documents.set(doc.id, doc);
  try {
    check(path.resolve(String(pres.FullName)).toLowerCase() === path.resolve(args.path).toLowerCase(),
      'DOCUMENT_IDENTITY_MISMATCH', 'PowerPoint opened a different readback path.');
    check(Number(pres.ReadOnly) === -1, 'NATIVE_READBACK_FAILED', 'Readback presentation must be read-only.');
    const snapshot = inspectDocument(doc);
    const images = [];
    if (args.render) {
      await fs.mkdir(args.render.directory, { recursive: true });
      const slides = pres.Slides;
      try {
        for (const number of [...new Set(args.render.slides)]) {
          check(number <= snapshot.slides.length, 'SLIDE_NOT_FOUND', 'Render slide is out of range.');
          const destination = path.join(args.render.directory, `slide-${number}.png`), slide = slides.Item(number);
          try { await invoke(() => slide.Export(destination, 'PNG', args.render.width, Math.round(args.render.width * snapshot.height / snapshot.width))); }
          finally { release(slide); }
          images.push({ slide: number, path: destination, sha256: hash(await fs.readFile(destination)) });
        }
      } finally { release(slides); }
    }
    return { nativeReadback: 'passed', readOnly: true, sha256, structural, slidesCount: snapshot.slides.length, snapshot, ...(args.render ? { images } : {}) };
  } finally { await closeDocument(doc); }
}

async function openDocument(args) {
  check(!documents.has(args.documentId), 'DOCUMENT_ALREADY_OPEN', 'Document ID is already bound in this worker.');
  const bytes = await fs.readFile(args.path); await checkNativeSafe(await readPackage(bytes));
  ensureApplication();
  process.send({ stage: 'opening-task-copy' });
  const { pres } = await invoke(() => ({ pres: app.Presentations.Open(args.path, 0, 0, args.visible ? -1 : 0) }));
  check(path.resolve(String(pres.FullName)).toLowerCase() === path.resolve(args.path).toLowerCase(), 'DOCUMENT_IDENTITY_MISMATCH', 'PowerPoint opened a different path.');
  const doc = { id: args.documentId, pres, openedPath: path.resolve(args.path), directory: args.directory, revision: args.revision || 0, generation: args.generation || 1, visible: args.visible, unusable: false };
  documents.set(doc.id, doc);
  try {
    process.send({ stage: 'saving-initial-checkpoint' });
    doc.checkpoint = await checkpoint(doc, path.join(doc.directory, `native-${doc.generation}-${doc.revision}.pptx`));
    process.send({ stage: 'reading-native-object-snapshot' });
    const snapshot = inspectDocument(doc); doc.stateHash = stateHash(snapshot);
    return { snapshot, checkpoint: doc.checkpoint, revision: doc.revision, generation: doc.generation, lease: appLease,
      binding: { documentId: doc.id, openedPath: doc.openedPath, worker: processIdentity(process.pid), office: appLease.identity, officeIdentityMethod: appLease.identityMethod, boundAt: new Date().toISOString() } };
  } catch (error) {
    doc.checkpoint = { path: args.path, sha256: hash(bytes) };
    await closeDocument(doc, { interrupted: true }).catch(() => {}); throw error;
  }
}

async function apply(args) {
  const doc = getDocument(args.documentId), snapshot = guard(doc);
  check(args.generation === doc.generation, 'GENERATION_CHANGED', 'Native generation changed.');
  preflight(snapshot, args.operations);
  if (args.dryRun) return { outcome: 'preflight_passed', revision: doc.revision, generation: doc.generation, operations: args.operations.length };
  // Fail before changing COM objects if the next checkpoint cannot be created.
  const nextCheckpointPath = path.join(doc.directory, `native-${doc.generation}-${doc.revision + 1}.pptx`);
  await prepareNewOutput(nextCheckpointPath);
  const completed = [];
  try {
    for (const [index, op] of args.operations.entries()) {
      await applyOperation(doc, op); completed.push({ index, type: op.type, key: op.target.key });
      pumpMessages(); await delay(0);
    }
    const nextSnapshot = inspectDocument(doc);
    const nextCheckpoint = await checkpoint(doc, nextCheckpointPath);
    doc.revision++; doc.checkpoint = nextCheckpoint; doc.stateHash = stateHash(nextSnapshot);
    return { outcome: 'completed', durable: true, revision: doc.revision, generation: doc.generation, snapshot: nextSnapshot, checkpoint: doc.checkpoint, changes: completed };
  } catch (error) {
    try {
      const saved = await fs.readFile(doc.checkpoint.path);
      check(hash(saved) === doc.checkpoint.sha256, 'CHECKPOINT_CORRUPT', 'Recovery checkpoint hash changed.'); await readPackage(saved);
      doc.pres.Saved = -1; doc.pres.Close(); release(doc.pres);
      doc.pres = (await invoke(() => ({ pres: app.Presentations.Open(doc.checkpoint.path, 0, 0, doc.visible ? -1 : 0) }))).pres;
      doc.openedPath = path.resolve(doc.checkpoint.path);
      doc.generation++; const restored = inspectDocument(doc); doc.stateHash = stateHash(restored);
      return { outcome: 'restored_from_checkpoint', durable: true, revision: doc.revision, generation: doc.generation, checkpoint: doc.checkpoint, snapshot: restored, completedBeforeFailure: completed, error: errorResult(error) };
    } catch (restoreError) {
      doc.unusable = true;
      return { outcome: 'outcome_unknown', durable: false, revision: doc.revision, generation: doc.generation, checkpoint: doc.checkpoint, completedBeforeFailure: completed, error: errorResult(error), recoveryError: errorResult(restoreError) };
    }
  }
}

async function validateDocument(doc, args) {
  if (args.expectedRevision !== undefined) check(args.expectedRevision === doc.revision, 'REVISION_MISMATCH', 'Refresh layout references before declaring design overlaps.');
  const snapshot = guard(doc);
  if (args.checks === 'layout') return { structural: 'not_run', nativeReadback: 'passed', slidesCount: snapshot.slides.length,
    revision: doc.revision, generation: doc.generation, snapshot };
  const bytes = await fs.readFile(doc.checkpoint.path);
  check(hash(bytes) === doc.checkpoint.sha256, 'CHECKPOINT_CORRUPT', 'Latest checkpoint bytes mismatch.');
  const parts = await readPackage(bytes);
  await checkNativeSafe(parts);
  const structural = validatePackage(parts);
  const slidesCount = Number(doc.pres.Slides.Count);
  return {
    structural,
    nativeReadback: 'passed',
    slidesCount,
    revision: doc.revision,
    generation: doc.generation,
    checkpoint: doc.checkpoint,
    ...(args.layoutCheck ? { snapshot } : {})
  };
}

async function commitDocument(doc, args) {
  guard(doc);
  check(args.generation === doc.generation, 'GENERATION_CHANGED', 'Native generation changed before commit.');
  const targetPath = path.resolve(args.outputPath);
  check(path.extname(targetPath).toLowerCase() === '.pptx', 'UNSUPPORTED_FORMAT', 'Output must be a .pptx file.');
  await prepareNewOutput(targetPath);
  const candidate = path.join(doc.directory, `commit-${randomUUID()}.pptx`);
  try {
    await prepareNewOutput(candidate);
    await invoke(() => doc.pres.SaveCopyAs(candidate, 24));
    const bytes = await fs.readFile(candidate);
    const parts = await readPackage(bytes);
    await checkNativeSafe(parts);
    const validation = { ...validatePackage(parts), basis: 'generated-output-bytes' };
    const published = await publishNew(targetPath, bytes);
    return { outputPath: published.path, sha256: published.sha256, bytes: published.bytes,
      revision: doc.revision, generation: doc.generation, validation };
  } finally { await fs.unlink(candidate).catch(() => {}); }
}

async function shutdown(interrupted = false) {
  const closed = [], errors = [];
  for (const doc of Array.from(documents.values())) {
    try { await closeDocument(doc, { interrupted }); closed.push(doc.id); }
    catch (error) { errors.push({ documentId: doc.id, ...errorResult(error) }); }
  }
  if (app && errors.length === 0 && appLease.createdByTask) {
    const presentations = app.Presentations, protectedViews = app.ProtectedViewWindows;
    let remaining;
    try { remaining = Number(presentations.Count) + Number(protectedViews.Count); }
    finally { release(protectedViews, presentations); }
    if (remaining === 0 && Number(app.Visible) === 0 && sameProcess(appLease.identity)) { await invoke(() => app.Quit()); appLease.quitRequested = true; }
    else appLease.quitSkipped = 'Application contains external presentations, is visible, or its process identity changed.';
  }
  // winax temporary property/method wrappers retain COM references. Collect
  // them before releasing Application and while this STA is still alive.
  await new Promise(resolve => setImmediate(resolve));
  globalThis.gc();
  pumpMessages();
  release(app); app = null;
  clearInterval(pump); releaseMutex?.(); releaseMutex = null; releaseSta?.();
  return { closed, errors, lease: appLease };
}

let queue = Promise.resolve();
process.on('message', message => {
  queue = queue.then(async () => {
    const { id, method, args = {}, ownerKey } = message;
    try {
      if (method === 'initialize') {
        check(!started, 'ALREADY_STARTED', 'Worker already initialized.');
        check(typeof ownerKey === 'string' && ownerKey.length >= 20, 'TASK_BINDING_REQUIRED', 'Worker requires a private task binding.');
        workerOwnerKey = ownerKey;
        check(typeof globalThis.gc === 'function', 'NATIVE_GC_UNAVAILABLE', 'Native executor must start with --expose-gc to release temporary COM references before shutdown.');
        releaseSta = initializeSta(); winax = require('winax'); started = true; pump = setInterval(pumpMessages, 10);
        process.send({ id, result: { ready: true, identity: processIdentity(process.pid) } }); return;
      }
      check(started, 'NOT_INITIALIZED', 'Worker has not been assigned and initialized.');
      check(ownerKey === workerOwnerKey, 'TASK_BINDING_MISMATCH', 'Request belongs to another native task.');
      let result;
      if (method === 'open') result = await openDocument(args);
      else if (method === 'inspect') { const doc = getDocument(args.documentId); result = { snapshot: guard(doc), revision: doc.revision, generation: doc.generation, checkpoint: doc.checkpoint }; }
      else if (method === 'apply') result = await apply(args);
      else if (method === 'validate') result = await validateDocument(getDocument(args.documentId), args);
      else if (method === 'validateFile') result = await validateFile(args);
      else if (method === 'commit') result = await commitDocument(getDocument(args.documentId), args);
      else if (method === 'render') {
        const doc = getDocument(args.documentId), snapshot = guard(doc), images = [];
        await fs.mkdir(args.directory, { recursive: true });
        for (const number of args.slides) {
          check(number <= snapshot.slides.length, 'SLIDE_NOT_FOUND', 'Render slide is out of range.');
          const destination = path.join(args.directory, `slide-${number}.png`), slide = doc.pres.Slides.Item(number);
          await invoke(() => slide.Export(destination, 'PNG', args.width, Math.round(args.width * snapshot.height / snapshot.width))); release(slide);
          images.push({ slide: number, path: destination, sha256: hash(await fs.readFile(destination)) });
        }
        result = { images, revision: doc.revision, generation: doc.generation, visualReview: 'not_run' };
      } else if (method === 'close') { const doc = getDocument(args.documentId); await closeDocument(doc, { interrupted: args.interrupted }); result = { closed: true }; }
      else if (method === 'shutdown') {
        stopping = true;
        result = await shutdown(args.interrupted);
        process.send({ id, result }, () => { process.disconnect(); process.exit(0); }); return;
      } else throw new PptError('UNKNOWN_NATIVE_METHOD', 'Unknown native method.');
      process.send({ id, result });
    } catch (error) { process.send?.({ id, error: errorResult(error) }); }
  }).catch(error => { process.stderr.write(JSON.stringify(errorResult(error)) + '\n'); process.exit(1); });
});
process.once('disconnect', () => { if (!stopping) queue.then(() => shutdown(true)).finally(() => process.exit(0)); });
