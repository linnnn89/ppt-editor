import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import path from 'node:path';
import { check, PptError } from './errors.js';

export const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types'
};
export const children = (node, namespace, name) => Array.from(node?.childNodes || []).filter(n => n.nodeType === 1 && n.namespaceURI === namespace && (!name || n.localName === name));
export const child = (node, namespace, name) => children(node, namespace, name)[0] || null;
export const descendants = (node, namespace, name) => Array.from(node.getElementsByTagNameNS(namespace, name));
export const xml = node => new XMLSerializer().serializeToString(node);
export function parseXml(bytes, label = 'XML') {
  let text;
  if (bytes[0] === 0xff && bytes[1] === 0xfe) text = bytes.subarray(2).toString('utf16le');
  else {
    check(!(bytes[0] === 0xfe && bytes[1] === 0xff), 'UNSUPPORTED_XML_ENCODING', 'Big-endian XML is not supported.');
    try { text = new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
    catch { throw new PptError('INVALID_XML', `${label} is not valid UTF-8.`); }
  }
  check(!/<!DOCTYPE|<!ENTITY/i.test(text), 'UNSAFE_XML', 'DTD and entity declarations are not allowed.');
  try {
    const doc = new DOMParser({ onError: (level, message) => { throw new Error(`${level}: ${message}`); } }).parseFromString(text, 'application/xml');
    check(doc.documentElement, 'INVALID_XML', `Missing root element in ${label}.`);
    return doc;
  } catch (error) { throw new PptError('INVALID_XML', `Invalid XML in ${label}.`, { reason: error.message }); }
}

export function resolvePart(source, target) {
  check(!/[\\\x00]/.test(target), 'INVALID_RELATIONSHIP', 'Invalid relationship target.');
  let decoded;
  try { decoded = decodeURIComponent(target.split('#')[0]); } catch { throw new PptError('INVALID_RELATIONSHIP', 'Invalid relationship URI.'); }
  const joined = path.posix.normalize(decoded.startsWith('/') ? decoded.slice(1) : path.posix.join(path.posix.dirname(source), decoded));
  check(joined !== '..' && !joined.startsWith('../') && !/^[a-z]+:/i.test(joined), 'INVALID_RELATIONSHIP', 'Relationship escapes the package.');
  return joined;
}
export const relationshipsPath = source => source ? path.posix.join(path.posix.dirname(source), '_rels', path.posix.basename(source) + '.rels') : '_rels/.rels';
export function relationshipsSource(name) {
  if (name === '_rels/.rels') return '';
  return path.posix.join(path.posix.dirname(path.posix.dirname(name)), path.posix.basename(name).slice(0, -5));
}

// Read central-directory sizes before decompression. ZIP64, encrypted archives,
// duplicates and ambiguous names are rejected rather than interpreted differently.
export function zipManifest(bytes) {
  check(bytes.length <= 100 * 1024 * 1024, 'PACKAGE_TOO_LARGE', 'Compressed package exceeds 100 MiB.');
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (bytes.readUInt32LE(i) === 0x06054b50 && i + 22 + bytes.readUInt16LE(i + 20) === bytes.length) { end = i; break; }
  }
  check(end >= 0, 'INVALID_PACKAGE', 'Expected a normal, non-encrypted PPTX ZIP package.');
  const count = bytes.readUInt16LE(end + 10), size = bytes.readUInt32LE(end + 12), start = bytes.readUInt32LE(end + 16);
  check(bytes.readUInt16LE(end + 4) === 0 && bytes.readUInt16LE(end + 6) === 0 && bytes.readUInt16LE(end + 8) === count, 'UNSUPPORTED_ZIP', 'Split archives are not supported.');
  check(count < 65535 && count <= 10000 && start !== 0xffffffff && start + size === end, 'UNSUPPORTED_ZIP', 'ZIP64 or excessive package entries are not supported.');
  const entries = [], names = new Set();
  let offset = start, total = 0;
  for (let i = 0; i < count; i++) {
    check(offset + 46 <= end && bytes.readUInt32LE(offset) === 0x02014b50, 'INVALID_PACKAGE', 'Invalid ZIP directory.');
    const flags = bytes.readUInt16LE(offset + 8), method = bytes.readUInt16LE(offset + 10);
    const compressed = bytes.readUInt32LE(offset + 20), uncompressed = bytes.readUInt32LE(offset + 24);
    const nameLength = bytes.readUInt16LE(offset + 28), extraLength = bytes.readUInt16LE(offset + 30), commentLength = bytes.readUInt16LE(offset + 32);
    const next = offset + 46 + nameLength + extraLength + commentLength;
    check(next <= end, 'INVALID_PACKAGE', 'Truncated ZIP entry.');
    const nameBytes = bytes.subarray(offset + 46, offset + 46 + nameLength);
    check((flags & 1) === 0 && [0, 8].includes(method), 'UNSUPPORTED_ZIP', 'Encrypted or unsupported compression in package.');
    check((flags & 0x800) || nameBytes.every(b => b < 128), 'UNSUPPORTED_ZIP', 'Non-UTF-8 ZIP names are not supported.');
    const name = new TextDecoder('utf-8', { fatal: true }).decode(nameBytes);
    check(name && !name.startsWith('/') && !name.includes('\\') && !name.includes('\0') && !name.split('/').includes('..') && !name.includes(':') && path.posix.normalize(name) === name, 'INVALID_PACKAGE', 'Unsafe or ambiguous package name.');
    check(!names.has(name.toLowerCase()), 'INVALID_PACKAGE', 'Duplicate package part.'); names.add(name.toLowerCase());
    total += uncompressed;
    check(uncompressed <= 100 * 1024 * 1024 && total <= 512 * 1024 * 1024, 'PACKAGE_TOO_LARGE', 'Expanded package exceeds resource limits.');
    entries.push({ name, compressed, uncompressed }); offset = next;
  }
  check(offset === end, 'INVALID_PACKAGE', 'ZIP directory length mismatch.');
  return entries;
}

export async function readPackage(bytes) {
  const entries = zipManifest(bytes);
  const zip = await JSZip.loadAsync(bytes, { checkCRC32: true, createFolders: false });
  const parts = new Map();
  for (const entry of entries) {
    if (entry.name.endsWith('/')) continue;
    const file = zip.file(entry.name); check(file, 'INVALID_PACKAGE', 'ZIP entry name changed while parsing.');
    const data = await file.async('nodebuffer');
    check(data.length === entry.uncompressed, 'INVALID_PACKAGE', 'Expanded ZIP size mismatch.');
    parts.set(entry.name, data);
  }
  validatePackage(parts);
  return parts;
}

export function rels(parts, source) {
  const name = relationshipsPath(source);
  if (!parts.has(name)) return [];
  const doc = parseXml(parts.get(name), name);
  return children(doc.documentElement, NS.rel, 'Relationship').map(node => ({ id: node.getAttribute('Id'), type: node.getAttribute('Type'), target: node.getAttribute('Target'), external: node.getAttribute('TargetMode') === 'External' }));
}

export function validatePackage(parts) {
  check(parts.has('[Content_Types].xml') && parts.has('_rels/.rels') && parts.has('ppt/presentation.xml'), 'INVALID_PRESENTATION', 'Missing required PowerPoint parts.');
  const unsupported = Array.from(parts.keys()).filter(name => /(^_xmlsignatures\/|vbaProject|encryption|drm|activeX\/)/i.test(name));
  check(unsupported.length === 0, 'UNSUPPORTED_PRESENTATION', 'Macros, signatures, DRM and ActiveX are outside the supported package scope.', { parts: unsupported });
  const main = parseXml(parts.get('ppt/presentation.xml'), 'ppt/presentation.xml');
  check(main.documentElement.namespaceURI === NS.p, 'UNSUPPORTED_PRESENTATION', 'Only transitional PresentationML is supported.');
  const contentTypes = parseXml(parts.get('[Content_Types].xml'));
  check(!xml(contentTypes).includes('macroEnabled'), 'UNSUPPORTED_PRESENTATION', 'Macro-enabled content is not supported.');
  const overrides = new Map(children(contentTypes.documentElement, NS.ct, 'Override').map(n => [n.getAttribute('PartName').replace(/^\//, ''), n.getAttribute('ContentType')]));
  const defaults = new Map(children(contentTypes.documentElement, NS.ct, 'Default').map(n => [n.getAttribute('Extension').toLowerCase(), n.getAttribute('ContentType')]));
  let relationshipCount = 0;
  for (const [name, data] of parts) {
    if (name !== '[Content_Types].xml') check(overrides.has(name) || defaults.has(name.split('.').at(-1).toLowerCase()), 'INVALID_CONTENT_TYPE', `No content type for ${name}.`);
    if (name.endsWith('.xml') || name.endsWith('.rels')) {
      const parsed = parseXml(data, name);
      if (/^ppt\/(slides|notesSlides)\/[^/]+\.xml$/.test(name)) {
        const ids = new Set();
        for (const node of descendants(parsed, NS.p, 'cNvPr')) {
          const id = node.getAttribute('id');
          check(/^\d+$/.test(id) && !ids.has(id), 'DUPLICATE_SHAPE_ID', `Invalid or duplicate shape ID in ${name}.`); ids.add(id);
        }
      }
    }
    if (!name.endsWith('.rels')) continue;
    const source = relationshipsSource(name), ids = new Set();
    if (source) check(parts.has(source), 'INVALID_RELATIONSHIP', `Missing relationship source ${source}.`);
    for (const rel of rels(parts, source)) {
      check(rel.id && !ids.has(rel.id), 'INVALID_RELATIONSHIP', `Duplicate relationship ID in ${name}.`); ids.add(rel.id); relationshipCount++;
      if (!rel.external) check(parts.has(resolvePart(source, rel.target)), 'MISSING_PART', `Relationship target is missing in ${name}.`);
    }
  }
  const presentationRelations = new Map(rels(parts, 'ppt/presentation.xml').map(r => [r.id, r]));
  const slideIds = new Set();
  for (const slide of descendants(main, NS.p, 'sldId')) {
    const id = slide.getAttribute('id'), rel = presentationRelations.get(slide.getAttributeNS(NS.r, 'id'));
    check(!slideIds.has(id) && rel && rel.type.endsWith('/slide') && !rel.external, 'INVALID_SLIDE_LIST', 'Invalid slide identity or relationship.'); slideIds.add(id);
  }
  return { level: 'package', passed: true, parts: parts.size, relationships: relationshipCount, slides: slideIds.size };
}

export async function checkNativeSafe(parts) {
  const workbooks = new Set();
  for (const name of parts.keys()) {
    check(!/(oleObject|activeX\/)/i.test(name), 'NATIVE_CONTENT_UNSUPPORTED', 'OLE and ActiveX objects are not enabled for native opening.');
    if (name.endsWith('.rels')) for (const rel of rels(parts, relationshipsSource(name))) {
      check(!rel.external || rel.type.endsWith('/hyperlink'), 'NATIVE_EXTERNAL_CONTENT', 'Native opening of external data/media relationships is not enabled.');
      check(!/\/(oleObject|activeXControl)$/.test(rel.type), 'NATIVE_CONTENT_UNSUPPORTED', 'OLE/ActiveX relationships are not supported.');
      if (!rel.external) {
        const source = relationshipsSource(name), target = resolvePart(source, rel.target);
        if (/^ppt\/embeddings\//i.test(target)) {
          check(/^ppt\/charts\/[^/]+\.xml$/.test(source) && rel.type.endsWith('/package') && /\.xlsx$/i.test(target),
            'NATIVE_CONTENT_UNSUPPORTED', 'Only chart-owned XLSX data workbooks are supported.');
          workbooks.add(target);
        }
      }
    }
  }
  for (const name of parts.keys()) if (/^ppt\/embeddings\//i.test(name)) {
    check(workbooks.has(name), 'NATIVE_CONTENT_UNSUPPORTED', 'Unreferenced or non-chart embedded content is unsupported.');
    const bytes = parts.get(name), entries = zipManifest(bytes);
    check(bytes.length <= 20 * 1024 * 1024 && entries.reduce((sum,e)=>sum+e.uncompressed,0) <= 40 * 1024 * 1024,
      'NATIVE_CONTENT_UNSUPPORTED', 'Chart workbook exceeds supported size.');
    const zip = await JSZip.loadAsync(bytes, { checkCRC32: true }), workbook = new Map();
    for (const entry of entries.filter(e=>!e.name.endsWith('/'))) {
      check(/^(\[Content_Types\]\.xml|_rels\/\.rels|docProps\/(app|core)\.xml|xl\/(workbook\.xml|styles\.xml|sharedStrings\.xml|calcChain\.xml|theme\/theme\d+\.xml|worksheets\/sheet\d+\.xml|tables\/table\d+\.xml|worksheets\/_rels\/sheet\d+\.xml\.rels|_rels\/workbook\.xml\.rels))$/.test(entry.name),
        'NATIVE_CONTENT_UNSUPPORTED', 'Chart workbook contains unsupported active or linked parts.');
      const data = await zip.file(entry.name).async('nodebuffer');
      check(data.length === entry.uncompressed, 'INVALID_PACKAGE', 'Workbook expanded size mismatch.');
      workbook.set(entry.name,data);
      const doc = parseXml(data);
      check(!/macroEnabled|vbaProject|oleObject/i.test(xml(doc)), 'NATIVE_CONTENT_UNSUPPORTED', 'Chart workbook has unsupported content types.');
      check(!Array.from(doc.getElementsByTagName('*')).some(n=>['f','calculatedColumnFormula','totalsRowFormula','externalReference','definedName'].includes(n.localName)),
        'NATIVE_CONTENT_UNSUPPORTED', 'Chart workbooks must contain static data, without formulas or external references.');
    }
    check(workbook.has('xl/workbook.xml') && workbook.has('[Content_Types].xml'), 'NATIVE_CONTENT_UNSUPPORTED', 'Invalid embedded chart workbook.');
    for (const p of workbook.keys()) if (p.endsWith('.rels')) for (const r of rels(workbook, relationshipsSource(p))) {
      check(!r.external && workbook.has(resolvePart(relationshipsSource(p),r.target)), 'NATIVE_EXTERNAL_CONTENT', 'Workbook relationships must resolve to internal static parts.');
    }
  }
}

export async function writePackage(parts) {
  const zip = new JSZip();
  for (const [name, data] of parts) zip.file(name, data, { date: new Date('2000-01-01T00:00:00Z'), createFolders: false });
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

/**
 * 优化写包实现：复用源包未变动部件的压缩字节流，避免全量重新 DEFLATE；
 * 对媒体文件（ppt/media/*）采用 STORE 策略，规避对二进制媒体重复压缩的 CPU 损耗。
 * 若无 sourceBytes 或解析失败，优雅降级为标准 writePackage。
 */
export async function writePackageOptimized(sourceBytes, currentParts, sourceParts) {
  if (!sourceBytes || !sourceParts) return writePackage(currentParts);
  try {
    const zip = await JSZip.loadAsync(sourceBytes, { createFolders: false });
    const fixedDate = new Date('2000-01-01T00:00:00Z');

    // 1. 清理在 currentParts 中已被移除的文件
    for (const relativePath of Object.keys(zip.files)) {
      // JSZip.remove(directory) recursively removes retained children too.
      if (zip.files[relativePath]?.dir) continue;
      if (!currentParts.has(relativePath)) {
        zip.remove(relativePath);
      }
    }

    // 2. 仅更新发生变更或新增的部件，媒体文件指定 STORE
    for (const [name, data] of currentParts) {
      const isMedia = name.startsWith('ppt/media/');
      const origData = sourceParts.get(name);
      const isChanged = !origData || !data.equals(origData);

      if (isChanged) {
        const options = { date: fixedDate, createFolders: false };
        if (isMedia) options.compression = 'STORE';
        zip.file(name, data, options);
      } else if (isMedia) {
        const existing = zip.file(name);
        const date = existing?.date || fixedDate;
        zip.file(name, data, { compression: 'STORE', date, createFolders: false });
      }
      // 未变动的 XML 和关系部件不调用 zip.file()，JSZip 自动复用原压缩数据流
    }

    return await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  } catch {
    // 容灾降级：若增量复用失败，回退至确定性的标准全量打包
    return writePackage(currentParts);
  }
}

export function setChild(parent, namespace, qualifiedName) {
  let node = child(parent, namespace, qualifiedName.split(':').at(-1));
  if (!node) { node = parent.ownerDocument.createElementNS(namespace, qualifiedName); parent.appendChild(node); }
  return node;
}
