import pptxgen from 'pptxgenjs';
import JSZip from 'jszip';
import { deckSchema } from './contracts.js';
import { NS, child, descendants, validatePackage, writePackage, parseXml, xml } from './ooxml.js';
import { check } from './errors.js';
import { setBackgroundXml, readBackgroundImage } from './background.js';
import { applyDesignTheme } from './design.js';
import { CHART_NS } from './advanced-objects.js';

export async function buildDeck(specification) {
  const deck = deckSchema.parse(applyDesignTheme(specification)), presentation = new pptxgen();
  presentation.defineLayout({ name: 'CUSTOM', width: deck.width / 72, height: deck.height / 72 });
  presentation.layout = 'CUSTOM'; presentation.author = ''; presentation.company = ''; presentation.subject = ''; presentation.title = deck.title;
  presentation.lang = 'en-GB';
  for (const page of deck.slides) {
    const slide = presentation.addSlide(); slide.background = { color: typeof page.background === 'string' ? page.background : 'FFFFFF' };
    for (const rawItem of page.items) {
      // Apply legacy defaults after theme selection; MCP schema parsing must not
      // turn an omitted style into an explicit override of the selected theme.
      const item = { fontFace: rawItem.type === 'chart' ? 'Microsoft YaHei' : 'Arial', color: '000000', line: '000000',
        colors: ['355DA8','368E87','B88540','8C6BB1'], ...rawItem };
      const geometry = { x: item.left / 72, y: item.top / 72, w: item.width / 72, h: item.height / 72, objectName: item.name };
      if (item.type === 'chart') {
        const options = { ...geometry, chartColors: item.colors, showLegend: item.showLegend, legendPos: 'b',
          showValue: item.showValue, showTitle: !!item.title, title: item.title, titleFontFace: item.fontFace, titleFontSize: 18,
          catAxisLabelFontFace: item.fontFace, catAxisLabelFontSize: 12, valAxisLabelFontFace: item.fontFace, valAxisLabelFontSize: 12,
          legendFontFace: item.fontFace, legendFontSize: 12, showBorder: false,
          barDir: item.horizontal ? 'bar' : 'col', barGrouping: item.stacked ? 'stacked' : 'clustered', grouping: item.stacked ? 'stacked' : 'standard' };
        const data = s => ({ name: s.name, labels: s.labels, values: s.values });
        if (item.chartType === 'combo') {
          if (item.series.some(s => s.secondaryAxis)) { options.catAxes = [{}, { catAxisHidden: true }]; options.valAxes = [{}, {}]; }
          slide.addChart(item.series.map(s => ({ type: presentation.ChartType[s.type], data: [data(s)], options: { secondaryValAxis: !!s.secondaryAxis, secondaryCatAxis: !!s.secondaryAxis, chartColors: [item.colors[item.series.indexOf(s) % item.colors.length]] } })), options);
        }
        else slide.addChart(presentation.ChartType[item.chartType], item.series.map(data), options);
      }
      if (item.type === 'text') slide.addText(item.text, { ...geometry, fontSize: item.fontSize, fontFace: item.fontFace, color: item.color, bold: item.bold, italic: item.italic, underline: item.underline, margin: 0, breakLine: false, valign: 'mid' });
      if (item.type === 'shape') {
        Object.assign(geometry, { flipV: item.flipV, flipH: item.flipH, rotate: item.rotation });
        const shapeType = presentation.ShapeType[item.shape];
        const lineOpts = {
          color: item.line,
          width: item.lineWidth ?? 1,
          ...(item.arrow === 'end' || item.arrow === 'both' ? { endArrowType: 'triangle' } : {}),
          ...(item.arrow === 'begin' || item.arrow === 'both' ? { beginArrowType: 'triangle' } : {})
        };
        const fillOpts = item.fill ? { fill: { color: item.fill } } : { fill: { type: 'none' } };
        if (item.text) {
          slide.addText(item.text, {
            ...geometry,
            shape: shapeType,
            ...fillOpts,
            line: lineOpts,
            fontSize: item.fontSize,
            fontFace: item.fontFace,
            color: item.color,
            bold: item.bold,
            italic: item.italic,
            underline: item.underline,
            align: item.align,
            valign: item.valign,
            margin: 0.05
          });
        } else {
          slide.addShape(shapeType, {
            ...geometry,
            ...fillOpts,
            line: lineOpts
          });
        }
      }
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
  // PptxGenJS 4.0.1 emits an unbacked series axis for 2D charts (#1534).
  // Normalize only generated charts; never silently repair imported user charts.
  for (const [part, bytes] of parts) if (/^ppt\/charts\/chart\d+\.xml$/.test(part)) {
    const doc = parseXml(bytes), ids = descendants(doc, CHART_NS, 'axId');
    const defined = new Set(ids.filter(n => ['catAx','valAx','serAx','dateAx'].includes(n.parentNode.localName)).map(n => n.getAttribute('val')));
    for (const id of ids.filter(n => n.parentNode.localName.endsWith('Chart'))) {
      if (!defined.has(id.getAttribute('val')) && id.getAttribute('val') === '2094734556') id.parentNode.removeChild(id);
      else check(defined.has(id.getAttribute('val')), 'INVALID_CHART_AXES', 'Generated chart refers to an undefined axis.');
    }
    parts.set(part, Buffer.from(xml(doc)));
  }
  // This generator version can emit duplicate shape IDs when text and tables
  // share a slide. Only normalize our own new, animation-free generated slides.
  // Existing user files with ambiguous identities are rejected by validation.
  for (const [index, page] of deck.slides.entries()) {
    const name = `ppt/slides/slide${index + 1}.xml`, doc = parseXml(parts.get(name));
    descendants(doc, NS.p, 'cNvPr').forEach((node, i) => node.setAttribute('id', String(i + 1)));
    const frames = descendants(doc, NS.p, 'graphicFrame').filter(frame => descendants(frame, NS.a, 'tbl').length);
    let tableIndex = 0;
    for (const item of page.items) if (item.type === 'table') {
      const frame = frames[tableIndex++], transform = child(frame, NS.p, 'xfrm');
      child(transform, NS.a, 'ext').setAttribute('cx', String(Math.round(item.width * 12700)));
      child(transform, NS.a, 'ext').setAttribute('cy', String(Math.round(item.height * 12700)));
      for (const col of descendants(frame, NS.a, 'gridCol')) col.setAttribute('w', String(Math.round(item.width * 12700 / item.rows[0].length)));
    }
    parts.set(name, Buffer.from(xml(doc)));
  }
  for (const [index, page] of deck.slides.entries()) if (page.background && typeof page.background !== 'string') {
    const fill = page.background, image = ['image', 'texture'].includes(fill.type) ? await readBackgroundImage(fill.path) : undefined;
    setBackgroundXml(parts, `ppt/slides/slide${index + 1}.xml`, fill, image);
  }
  // Clear incidental generator metadata while retaining the user's document title.
  for (const name of ['docProps/core.xml', 'docProps/app.xml']) {
    if (!parts.has(name)) continue;
    const doc = parseXml(parts.get(name));
    for (const element of Array.from(doc.getElementsByTagName('*'))) if (['creator', 'lastModifiedBy', 'Application', 'AppVersion', 'Company', 'Manager'].includes(element.localName)) element.parentNode.removeChild(element);
    parts.set(name, Buffer.from(xml(doc)));
  }
  validatePackage(parts); return writePackage(parts);
}
