import { promises as fs } from 'node:fs';
import path from 'node:path';
import { check } from './errors.js';
import { hash, stableJson, writeJson, readJson, atomicWrite } from './storage.js';
import { NS, child, children, descendants, parseXml, xml, rels, resolvePart, setChild, readPackage, writePackage, validatePackage } from './ooxml.js';

const PT = 12700;
const numeric = (node, attr, factor = 1) => node?.hasAttribute(attr) ? Number(node.getAttribute(attr)) / factor : null;
const colorNode = node => child(child(node, NS.a, 'solidFill'), NS.a, 'srgbClr');

export function textBodyText(body) {
  return children(body, NS.a, 'p').map(paragraph => children(paragraph, NS.a).map(run => {
    if (run.localName === 'br') return '\n';
    if (['r', 'fld'].includes(run.localName)) return child(run, NS.a, 't')?.textContent || '';
    return '';
  }).join('')).join('\n');
}

function runStyle(run) {
  const prop = child(run, NS.a, 'rPr');
  const color = colorNode(prop)?.getAttribute('val') || null;
  const scheme = child(child(prop, NS.a, 'solidFill'), NS.a, 'schemeClr')?.getAttribute('val') || null;
  return { fontSize: numeric(prop, 'sz', 100), fontFace: child(prop, NS.a, 'latin')?.getAttribute('typeface') || null, bold: prop?.hasAttribute('b') ? prop.getAttribute('b') === '1' : null, italic: prop?.hasAttribute('i') ? prop.getAttribute('i') === '1' : null, color, themeColor: scheme, source: 'explicit-run-properties; null means inherited or unspecified' };
}

function geometryOf(shape) {
  const transform = child(child(shape, NS.p, 'spPr'), NS.a, 'xfrm') || child(shape, NS.p, 'xfrm') || child(child(shape, NS.p, 'grpSpPr'), NS.a, 'xfrm');
  const offset = child(transform, NS.a, 'off'), extent = child(transform, NS.a, 'ext');
  return { left: numeric(offset, 'x', PT), top: numeric(offset, 'y', PT), width: numeric(extent, 'cx', PT), height: numeric(extent, 'cy', PT), rotation: numeric(transform, 'rot', 60000) ?? 0, unit: 'pt', source: transform ? 'explicit' : 'inherited-unresolved' };
}

export function indexPackage(parts) {
  const presentation = parseXml(parts.get('ppt/presentation.xml'));
  const relationMap = new Map(rels(parts, 'ppt/presentation.xml').map(r => [r.id, r]));
  const size = child(presentation.documentElement, NS.p, 'sldSz');
  const slides = [], objects = [];
  for (const entry of descendants(presentation, NS.p, 'sldId')) {
    const slideId = Number(entry.getAttribute('id')), slideNumber = slides.length + 1;
    const part = resolvePart('ppt/presentation.xml', relationMap.get(entry.getAttributeNS(NS.r, 'id')).target);
    const doc = parseXml(parts.get(part), part);
    const tree = descendants(doc, NS.p, 'spTree')[0];
    const slide = { slide: slideNumber, slideId, part, objectCount: 0 };
    function visit(container, groups = [], notes = false, sourcePart = part) {
      for (const shape of children(container, NS.p)) {
        if (!['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp'].includes(shape.localName)) continue;
        const identity = descendants(shape, NS.p, 'cNvPr')[0];
        if (!identity) continue;
        const id = Number(identity.getAttribute('id'));
        const body = child(shape, NS.p, 'txBody');
        const table = descendants(shape, NS.a, 'tbl')[0];
        const placeholder = descendants(shape, NS.p, 'ph')[0];
        if (notes && placeholder?.getAttribute('type') !== 'body') continue;
        const kind = notes ? 'notes' : table ? 'table' : body ? 'text' : shape.localName === 'pic' ? 'image' : shape.localName === 'grpSp' ? 'group' : 'shape';
        const tableRows = table ? children(table, NS.a, 'tr').map(row => children(row, NS.a, 'tc').map(cell => textBodyText(child(cell, NS.a, 'txBody')))) : undefined;
        const key = `${notes ? 'notes' : 'slide'}:${slideId}:${groups.join('.')}:${id}`;
        const item = {
          key, part: sourcePart, slide: slideNumber, slideId, shapeId: id, groupPath: groups, name: identity.getAttribute('name') || '', kind,
          text: body ? textBodyText(body) : tableRows ? tableRows.map(row => row.join('\t')).join('\n') : '',
          geometry: geometryOf(shape), placeholder: placeholder?.getAttribute('type') || null,
          runs: body ? children(body, NS.a, 'p').flatMap((p, paragraph) => children(p, NS.a, 'r').map(r => ({ paragraph, text: child(r, NS.a, 't')?.textContent || '', style: runStyle(r) }))) : [],
          ...(tableRows ? { rows: tableRows, rowCount: tableRows.length, columnCount: tableRows[0]?.length || 0 } : {})
        };
        item.fingerprint = hash(xml(shape));
        item.capabilities = [
          ...(body ? ['replace_text', 'set_style'] : []), ...(table ? ['set_table_cell'] : []),
          ...(!notes && groups.length === 0 && kind !== 'group' && item.geometry.source === 'explicit' ? ['set_geometry'] : [])
        ];
        objects.push(item); slide.objectCount++;
        if (shape.localName === 'grpSp') visit(shape, [...groups, id], notes, sourcePart);
      }
    }
    if (tree) visit(tree);
    const notesRel = rels(parts, part).find(r => r.type.endsWith('/notesSlide') && !r.external);
    if (notesRel) {
      const notePart = resolvePart(part, notesRel.target), notesDoc = parseXml(parts.get(notePart), notePart);
      const noteTree = descendants(notesDoc, NS.p, 'spTree')[0];
      if (noteTree) visit(noteTree, [], true, notePart);
    }
    slides.push(slide);
  }
  return { width: numeric(size, 'cx', PT), height: numeric(size, 'cy', PT), slides, objects };
}

function locateShape(doc, locator) {
  let container = descendants(doc, NS.p, 'spTree')[0];
  for (const id of locator.groupPath) {
    container = children(container, NS.p, 'grpSp').find(shape => Number(descendants(shape, NS.p, 'cNvPr')[0]?.getAttribute('id')) === id);
    check(container, 'TARGET_NOT_FOUND', 'Group no longer exists.');
  }
  const shape = children(container, NS.p).find(shape => Number(descendants(shape, NS.p, 'cNvPr')[0]?.getAttribute('id')) === locator.shapeId);
  check(shape, 'TARGET_NOT_FOUND', 'Shape no longer exists.'); return shape;
}

function replaceBody(body, op) {
  const matches = [];
  for (const paragraph of children(body, NS.a, 'p')) {
    const segments = [];
    let content = '';
    for (const node of children(paragraph, NS.a)) {
      if (node.localName === 'br') { content += '\n'; continue; }
      const value = ['r', 'fld'].includes(node.localName) ? child(node, NS.a, 't') : null;
      if (!value) continue;
      segments.push({ node: value, start: content.length, end: content.length + value.textContent.length, field: node.localName === 'fld' }); content += value.textContent;
    }
    let cursor = 0;
    while ((cursor = content.indexOf(op.search, cursor)) !== -1) {
      const selected = segments.filter(s => s.end > cursor && s.start < cursor + op.search.length);
      check(selected.length > 0 && selected.every(s => !s.field), 'UNSUPPORTED_TEXT_RANGE', 'Field text cannot be replaced.');
      check(selected.length === 1 || op.crossRunPolicy === 'first_run', 'CROSS_RUN_REPLACEMENT', 'Text spans multiple runs; explicitly select first_run policy.');
      matches.push({ segments: selected, start: cursor, end: cursor + op.search.length }); cursor += op.search.length;
    }
  }
  check(matches.length === op.expectedMatches, 'MATCH_COUNT_MISMATCH', 'Text match count differs from the requested count.', { expected: op.expectedMatches, actual: matches.length });
  for (const match of matches.reverse()) {
    const first = match.segments[0], last = match.segments.at(-1);
    if (first === last) {
      const text = first.node.textContent;
      first.node.textContent = text.slice(0, match.start - first.start) + op.replacement + text.slice(match.end - first.start);
    } else {
      first.node.textContent = first.node.textContent.slice(0, match.start - first.start) + op.replacement;
      for (const middle of match.segments.slice(1, -1)) middle.node.textContent = '';
      last.node.textContent = last.node.textContent.slice(match.end - last.start);
    }
  }
  return { matches: matches.length };
}

function styleBody(body, style) {
  for (const paragraph of children(body, NS.a, 'p')) {
    const properties = children(paragraph, NS.a, 'r').map(r => {
      let prop = child(r, NS.a, 'rPr');
      if (!prop) { prop = r.ownerDocument.createElementNS(NS.a, 'a:rPr'); r.insertBefore(prop, r.firstChild); }
      return prop;
    });
    properties.push(setChild(paragraph, NS.a, 'a:endParaRPr'));
    for (const prop of properties) {
      if (style.fontSize !== undefined) prop.setAttribute('sz', String(Math.round(style.fontSize * 100)));
      if (style.bold !== undefined) prop.setAttribute('b', style.bold ? '1' : '0');
      if (style.italic !== undefined) prop.setAttribute('i', style.italic ? '1' : '0');
      if (style.fontFace !== undefined) for (const name of ['a:latin', 'a:ea', 'a:cs']) setChild(prop, NS.a, name).setAttribute('typeface', style.fontFace);
      if (style.color !== undefined) {
        for (const item of children(prop, NS.a)) if (['solidFill', 'noFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill'].includes(item.localName)) prop.removeChild(item);
        const fill = prop.ownerDocument.createElementNS(NS.a, 'a:solidFill');
        // Fill precedes font-family and hyperlink nodes in CT_TextCharacterProperties.
        const next = children(prop, NS.a).find(n => ['effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'].includes(n.localName));
        prop.insertBefore(fill, next || null); setChild(fill, NS.a, 'a:srgbClr').setAttribute('val', style.color.toUpperCase());
      }
    }
  }
}

function setCell(table, op) {
  const row = children(table, NS.a, 'tr')[op.row - 1], cell = children(row, NS.a, 'tc')[op.column - 1];
  check(cell, 'CELL_NOT_FOUND', 'Table cell is out of range.');
  check(!['rowSpan', 'gridSpan', 'hMerge', 'vMerge'].some(a => cell.hasAttribute(a) && !['0', '1'].includes(cell.getAttribute(a))) && cell.getAttribute('hMerge') !== '1' && cell.getAttribute('vMerge') !== '1', 'MERGED_CELL_UNSUPPORTED', 'Merged cells require a separate capability.');
  const body = child(cell, NS.a, 'txBody'); check(body, 'CELL_NOT_FOUND', 'Table cell has no text body.');
  const firstParagraph = child(body, NS.a, 'p'), firstRun = firstParagraph && child(firstParagraph, NS.a, 'r');
  const originalProperties = firstRun && child(firstRun, NS.a, 'rPr');
  const paragraphProperties = firstParagraph && child(firstParagraph, NS.a, 'pPr');
  for (const paragraph of children(body, NS.a, 'p')) body.removeChild(paragraph);
  for (const line of op.text.replaceAll('\r\n', '\n').replaceAll('\r', '\n').split('\n')) {
    const p = body.ownerDocument.createElementNS(NS.a, 'a:p'); body.appendChild(p);
    if (paragraphProperties) p.appendChild(paragraphProperties.cloneNode(true));
    const r = body.ownerDocument.createElementNS(NS.a, 'a:r'); p.appendChild(r);
    if (originalProperties) r.appendChild(originalProperties.cloneNode(true));
    setChild(r, NS.a, 'a:t').textContent = line;
  }
}

export class FileEngine {
  constructor(directory, sourceBytes, parts, revision = 0) { this.directory = directory; this.sourceBytes = sourceBytes; this.sourceParts = parts; this.parts = new Map(parts); this.revision = revision; }

  static async create(directory, bytes) {
    const parts = await readPackage(bytes);
    const engine = new FileEngine(directory, bytes, parts);
    await atomicWrite(path.join(directory, 'source.pptx'), bytes);
    await engine.persist(); return engine;
  }

  static async restore(directory) {
    const bytes = await fs.readFile(path.join(directory, 'source.pptx'));
    const meta = await readJson(path.join(directory, 'revision.json'));
    check(hash(bytes) === meta.sourceHash, 'CHECKPOINT_CORRUPT', 'Source snapshot hash mismatch.');
    const engine = new FileEngine(directory, bytes, await readPackage(bytes), meta.revision);
    for (const [name, entry] of Object.entries(meta.overlay)) {
      check(/^[a-f0-9]{64}\.part$/.test(entry.file), 'CHECKPOINT_CORRUPT', 'Invalid checkpoint part name.');
      const data = await fs.readFile(path.join(directory, 'parts', entry.file));
      check(hash(data) === entry.sha256, 'CHECKPOINT_CORRUPT', 'Checkpoint part hash mismatch.'); engine.parts.set(name, data);
    }
    validatePackage(engine.parts); return engine;
  }

  inspect() { return indexPackage(this.parts); }
  validate() {
    const structural = validatePackage(this.parts);
    const changedParts = [], preservedParts = [];
    for (const [name, bytes] of this.parts) (this.sourceParts.has(name) && bytes.equals(this.sourceParts.get(name)) ? preservedParts : changedParts).push(name);
    return { structural, preservation: { basis: 'uncompressed part bytes', changedParts, unchangedPartCount: preservedParts.length, unknownPartsRetained: true }, nativeReadback: 'not_run', visualReview: 'not_run' };
  }

  prepare(operations) {
    const baseline = new Map(this.inspect().objects.map(o => [o.key, o])), candidate = new Map(this.parts), documents = new Map(), changes = [];
    for (const op of operations) {
      const item = baseline.get(op.target.key);
      check(item && item.fingerprint === op.target.fingerprint, 'TARGET_CHANGED', 'Target changed since inspection.');
      check(item.capabilities.includes(op.type), 'OPERATION_UNSUPPORTED', 'Operation is not supported on this object.', { type: op.type, kind: item.kind });
    }
    for (const op of operations) {
      const item = baseline.get(op.target.key);
      if (!documents.has(item.part)) documents.set(item.part, parseXml(candidate.get(item.part), item.part));
      const doc = documents.get(item.part), shape = locateShape(doc, item), body = child(shape, NS.p, 'txBody');
      let detail = {};
      if (op.type === 'replace_text') detail = replaceBody(body, op);
      if (op.type === 'set_style') styleBody(body, op.style);
      if (op.type === 'set_geometry') {
        const transform = child(child(shape, NS.p, 'spPr'), NS.a, 'xfrm') || child(shape, NS.p, 'xfrm');
        check(transform, 'GEOMETRY_UNSUPPORTED', 'Explicit geometry is required.');
        const offset = child(transform, NS.a, 'off'), extent = child(transform, NS.a, 'ext');
        for (const [key, attr] of [['left', 'x'], ['top', 'y']]) if (op.geometry[key] !== undefined) offset.setAttribute(attr, String(Math.round(op.geometry[key] * PT)));
        for (const [key, attr] of [['width', 'cx'], ['height', 'cy']]) if (op.geometry[key] !== undefined) extent.setAttribute(attr, String(Math.round(op.geometry[key] * PT)));
        if (op.geometry.rotation !== undefined) transform.setAttribute('rot', String(Math.round(((op.geometry.rotation % 360 + 360) % 360) * 60000)));
      }
      if (op.type === 'set_table_cell') setCell(descendants(shape, NS.a, 'tbl')[0], op);
      changes.push({ type: op.type, key: item.key, ...detail });
    }
    for (const [name, doc] of documents) candidate.set(name, Buffer.from(xml(doc)));
    validatePackage(candidate);
    return { candidate, changes };
  }

  async apply(operations, { dryRun = false } = {}) {
    const prepared = this.prepare(operations);
    if (dryRun) return { outcome: 'preflight_passed', revision: this.revision, changes: prepared.changes };
    const previous = this.parts, revision = this.revision;
    this.parts = prepared.candidate; this.revision++;
    try { await this.persist(); } catch (error) { this.parts = previous; this.revision = revision; throw error; }
    return { outcome: 'completed', durable: true, revision: this.revision, changes: prepared.changes, stateHash: hash(stableJson(this.inspect().objects.map(o => [o.key, o.fingerprint]))) };
  }

  async persist() {
    const overlay = {};
    for (const [name, bytes] of this.parts) {
      if (this.sourceParts.has(name) && bytes.equals(this.sourceParts.get(name))) continue;
      const sha256 = hash(bytes), file = `${sha256}.part`;
      await atomicWrite(path.join(this.directory, 'parts', file), bytes); overlay[name] = { file, sha256 };
    }
    await writeJson(path.join(this.directory, 'revision.json'), { revision: this.revision, sourceHash: hash(this.sourceBytes), overlay });
  }

  async bytes() { return writePackage(this.parts); }
}
