import { promises as fs } from 'node:fs';
import path from 'node:path';
import { check } from './errors.js';
import { hash } from './storage.js';
import { NS, child, children, descendants, parseXml, xml, relationshipsPath, rels, resolvePart } from './ooxml.js';

// DrawingML preset name and Office MsoPatternType value.
export const backgroundPatterns = {
  horizontal: ['horz', 49], vertical: ['vert', 50], cross: ['cross', 51],
  downDiagonal: ['dnDiag', 52], upDiagonal: ['upDiag', 53], diagonalCross: ['diagCross', 54],
  dots5: ['pct5', 1], dots10: ['pct10', 2], dots20: ['pct20', 3], dots25: ['pct25', 4], dots50: ['pct50', 7]
};

export async function readBackgroundImage(source) {
  check(typeof source === 'string' && !/^\\\\|^[a-z]+:\/\//i.test(source), 'INVALID_IMAGE_PATH', 'Use a local PNG or JPEG file path.');
  const handle = await fs.open(path.resolve(source), 'r');
  let bytes;
  try {
    const stat = await handle.stat();
    check(stat.isFile() && stat.size > 0 && stat.size <= 20 * 1024 * 1024, 'INVALID_IMAGE', 'Background image must be a file of at most 20 MiB.');
    bytes = await handle.readFile();
    check(bytes.length <= 20 * 1024 * 1024, 'INVALID_IMAGE', 'Background image exceeds 20 MiB.');
  } finally { await handle.close(); }
  const png = bytes.length >= 24 && bytes.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
  const jpeg = bytes.length >= 4 && bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255;
  check(png || jpeg, 'INVALID_IMAGE', 'Background images currently support PNG and JPEG only.');
  if (png) check(bytes.readUInt32BE(16) > 0 && bytes.readUInt32BE(20) > 0 && bytes.readUInt32BE(16) * bytes.readUInt32BE(20) <= 100000000,
    'INVALID_IMAGE', 'PNG background dimensions exceed the supported limit.');
  return { bytes, extension: png ? 'png' : 'jpeg', sha256: hash(bytes) };
}

const add = (parent, name, attrs = {}) => {
  const node = parent.ownerDocument.createElementNS(NS.a, `a:${name}`);
  for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, String(value));
  parent.appendChild(node); return node;
};
const addColor = (parent, color, transparency = 0) => {
  const node = add(parent, 'srgbClr', { val: color.toUpperCase() });
  if (transparency) add(node, 'alpha', { val: Math.round((1 - transparency / 100) * 100000) });
};

export function setBackgroundXml(parts, slidePart, fill, image, existingDoc) {
  const doc = existingDoc || parseXml(parts.get(slidePart)), common = child(doc.documentElement, NS.p, 'cSld');
  check(common, 'CORRUPT_SLIDE', 'Slide root does not contain cSld.');
  const previous = child(common, NS.p, 'bg'); if (previous) common.removeChild(previous);
  if (fill.type === 'inherit') { parts.set(slidePart, Buffer.from(xml(doc))); return; }
  const bg = doc.createElementNS(NS.p, 'p:bg'), props = doc.createElementNS(NS.p, 'p:bgPr');
  common.insertBefore(bg, child(common, NS.p, 'spTree')); bg.appendChild(props);
  if (fill.type === 'solid') addColor(add(props, 'solidFill'), fill.color, fill.transparency);
  else if (fill.type === 'gradient') {
    const gradient = add(props, 'gradFill', { rotWithShape: '1' }), stops = add(gradient, 'gsLst');
    addColor(add(stops, 'gs', { pos: 0 }), fill.startColor, fill.transparency);
    addColor(add(stops, 'gs', { pos: 100000 }), fill.endColor, fill.transparency);
    add(gradient, 'lin', { ang: Math.round((fill.angle % 360) * 60000), scaled: '0' });
  } else if (fill.type === 'pattern') {
    const pattern = add(props, 'pattFill', { prst: backgroundPatterns[fill.pattern][0] });
    addColor(add(pattern, 'fgClr'), fill.foreground, fill.transparency);
    addColor(add(pattern, 'bgClr'), fill.background, fill.transparency);
  } else {
    check(image?.bytes, 'INVALID_IMAGE', 'Image bytes must be prepared before modifying the package.');
    const imagePart = `ppt/media/bg-${image.sha256}.${image.extension}`;
    parts.set(imagePart, image.bytes);
    const relPart = relationshipsPath(slidePart);
    const relationships = parseXml(parts.get(relPart) || Buffer.from(`<Relationships xmlns="${NS.rel}"/>`));
    const target = path.posix.relative(path.posix.dirname(slidePart), imagePart);
    const all = descendants(relationships, NS.rel, 'Relationship');
    let relation = all.find(r => r.getAttribute('Target') === target && r.getAttribute('Type').endsWith('/image'));
    if (!relation) {
      const used = new Set(all.map(r => r.getAttribute('Id'))); let n = 1; while (used.has(`rIdBackground${n}`)) n++;
      relation = relationships.createElementNS(NS.rel, 'Relationship');
      relation.setAttribute('Id', `rIdBackground${n}`); relation.setAttribute('Target', target); relation.setAttribute('Type', `${NS.r}/image`);
      relationships.documentElement.appendChild(relation);
    }
    parts.set(relPart, Buffer.from(xml(relationships)));
    const types = parseXml(parts.get('[Content_Types].xml'));
    if (!descendants(types, NS.ct, 'Default').some(n => n.getAttribute('Extension').toLowerCase() === image.extension)) {
      const entry = types.createElementNS(NS.ct, 'Default'); entry.setAttribute('Extension', image.extension);
      entry.setAttribute('ContentType', `image/${image.extension}`); types.documentElement.appendChild(entry);
      parts.set('[Content_Types].xml', Buffer.from(xml(types)));
    }
    const picture = add(props, 'blipFill', { dpi: 96, rotWithShape: '1' }), blip = add(picture, 'blip');
    blip.setAttributeNS(NS.r, 'r:embed', relation.getAttribute('Id'));
    if (fill.transparency) add(blip, 'alphaModFix', { amt: Math.round((1 - fill.transparency / 100) * 100000) });
    if (fill.type === 'texture') add(picture, 'tile', { tx: 0, ty: 0, sx: 100000, sy: 100000, flip: 'none', algn: 'tl' });
    else add(add(picture, 'stretch'), 'fillRect');
  }
  add(props, 'effectLst'); parts.set(slidePart, Buffer.from(xml(doc)));
}

export function describeBackground(parts, slidePart, bg) {
  if (!bg) return { type: 'inherit', source: 'layout-or-master' };
  const props = child(bg, NS.p, 'bgPr');
  if (!props) return { type: 'theme', source: 'background-reference' };
  const color = node => child(node, NS.a, 'srgbClr')?.getAttribute('val') || null;
  const transparency = node => {
    const alpha = node && descendants(node, NS.a, 'alpha')[0];
    return alpha ? 100 - Number(alpha.getAttribute('val')) / 1000 : 0;
  };
  const solid = child(props, NS.a, 'solidFill');
  if (solid) return { type: 'solid', color: color(solid), transparency: transparency(solid) };
  const gradient = child(props, NS.a, 'gradFill');
  if (gradient) return { type: 'gradient', angle: Number(child(gradient, NS.a, 'lin')?.getAttribute('ang') || 0) / 60000,
    stops: descendants(gradient, NS.a, 'gs').map(n => ({ position: Number(n.getAttribute('pos')) / 100000, color: color(n), transparency: transparency(n) })) };
  const pattern = child(props, NS.a, 'pattFill');
  if (pattern) return { type: 'pattern', pattern: Object.entries(backgroundPatterns).find(([, value]) => value[0] === pattern.getAttribute('prst'))?.[0] || pattern.getAttribute('prst'),
    foreground: color(child(pattern, NS.a, 'fgClr')), background: color(child(pattern, NS.a, 'bgClr')), transparency: transparency(pattern) };
  const picture = child(props, NS.a, 'blipFill');
  if (picture) {
    const blip = child(picture, NS.a, 'blip'), id = blip?.getAttributeNS(NS.r, 'embed');
    const rel = rels(parts, slidePart).find(r => r.id === id && !r.external);
    const imagePart = rel ? resolvePart(slidePart, rel.target) : null;
    const alpha = child(blip, NS.a, 'alphaModFix');
    return { type: child(picture, NS.a, 'tile') ? 'texture' : 'image', imagePart,
      sha256: imagePart && parts.has(imagePart) ? hash(parts.get(imagePart)) : null,
      transparency: alpha ? 100 - Number(alpha.getAttribute('amt')) / 1000 : 0 };
  }
  return { type: 'other', source: 'unsupported-fill' };
}
