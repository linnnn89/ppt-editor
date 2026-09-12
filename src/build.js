import pptxgen from 'pptxgenjs';
import JSZip from 'jszip';
import { deckSchema } from './contracts.js';
import { NS, child, descendants, validatePackage, writePackage, parseXml, xml } from './ooxml.js';
import { check } from './errors.js';

export async function buildDeck(specification) {
  const deck = deckSchema.parse(specification), presentation = new pptxgen();
  presentation.defineLayout({ name: 'CUSTOM', width: deck.width / 72, height: deck.height / 72 });
  presentation.layout = 'CUSTOM'; presentation.author = ''; presentation.company = ''; presentation.subject = ''; presentation.title = deck.title;
  presentation.lang = 'en-GB';
  for (const page of deck.slides) {
    const slide = presentation.addSlide(); slide.background = { color: page.background };
    for (const item of page.items) {
      const geometry = { x: item.left / 72, y: item.top / 72, w: item.width / 72, h: item.height / 72, objectName: item.name };
      if (item.type === 'text') slide.addText(item.text, { ...geometry, fontSize: item.fontSize, fontFace: item.fontFace, color: item.color, bold: item.bold, margin: 0, breakLine: false, valign: 'mid' });
      if (item.type === 'shape') slide.addShape(presentation.ShapeType[item.shape], { ...geometry, ...(item.fill ? { fill: { color: item.fill } } : { fill: { transparency: 100 } }), line: { color: item.line, width: 1 } });
      if (item.type === 'table') {
        check(item.rows.every(row => row.length === item.rows[0].length), 'INVALID_TABLE', 'Generated table rows must have equal column counts.');
        slide.addTable(item.rows, { ...geometry, fontSize: item.fontSize, margin: 0.04, border: { type: 'solid', color: '000000', pt: 1 }, color: '000000', fill: 'FFFFFF', autoPage: false, rowH: item.height / 72 / item.rows.length, colW: item.width / 72 / item.rows[0].length });
      }
    }
    if (page.notes) slide.addNotes(page.notes);
  }
  const generated = await JSZip.loadAsync(await presentation.write({ outputType: 'nodebuffer' }));
  const parts = new Map();
  for (const entry of Object.values(generated.files)) if (!entry.dir) parts.set(entry.name, await entry.async('nodebuffer'));
  // This generator version can emit duplicate shape IDs when text and tables
  // share a slide. Only normalize our own new, animation-free generated slides.
  // Existing user files with ambiguous identities are rejected by validation.
  for (const [index, page] of deck.slides.entries()) {
    const name = `ppt/slides/slide${index + 1}.xml`, doc = parseXml(parts.get(name));
    descendants(doc, NS.p, 'cNvPr').forEach((node, i) => node.setAttribute('id', String(i + 1)));
    const frames = descendants(doc, NS.p, 'graphicFrame');
    let tableIndex = 0;
    for (const item of page.items) if (item.type === 'table') {
      const frame = frames[tableIndex++], transform = child(frame, NS.p, 'xfrm');
      child(transform, NS.a, 'ext').setAttribute('cx', String(Math.round(item.width * 12700)));
      child(transform, NS.a, 'ext').setAttribute('cy', String(Math.round(item.height * 12700)));
      for (const col of descendants(frame, NS.a, 'gridCol')) col.setAttribute('w', String(Math.round(item.width * 12700 / item.rows[0].length)));
    }
    parts.set(name, Buffer.from(xml(doc)));
  }
  // Clear incidental generator metadata while retaining the user's document title.
  for (const name of ['docProps/core.xml', 'docProps/app.xml']) {
    if (!parts.has(name)) continue;
    const doc = parseXml(parts.get(name));
    for (const element of Array.from(doc.getElementsByTagName('*'))) if (['creator', 'lastModifiedBy', 'Application', 'AppVersion', 'Company'].includes(element.localName)) element.parentNode.removeChild(element);
    parts.set(name, Buffer.from(xml(doc)));
  }
  validatePackage(parts); return writePackage(parts);
}
