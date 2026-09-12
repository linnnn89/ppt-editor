import { NS, descendants, child, parseXml, rels, resolvePart } from './ooxml.js';
import { hash } from './storage.js';
export const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
export const DIAGRAM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

export function advancedObject(parts, part, shape) {
  const chartRef = descendants(shape, CHART_NS, 'chart')[0], diagramRef = descendants(shape, DIAGRAM_NS, 'relIds')[0];
  const relationship = id => rels(parts, part).find(r => r.id === id && !r.external);
  if (chartRef) {
    const relation = relationship(chartRef.getAttributeNS(NS.r, 'id'));
    if (!relation) return { kind: 'chart', chart: { unavailable: true } };
    const chartPart = resolvePart(part, relation.target), doc = parseXml(parts.get(chartPart));
    const values = (node, name) => node ? descendants(node, CHART_NS, name).map(n => n.textContent) : [];
    return { kind: 'chart', chart: { part: chartPart, sha256: hash(parts.get(chartPart)),
      types: [...new Set(Array.from(doc.getElementsByTagName('*')).filter(n => n.namespaceURI === CHART_NS && n.localName.endsWith('Chart')).map(n => n.localName))],
      series: descendants(doc, CHART_NS, 'ser').map(s => ({ name: values(child(s, CHART_NS, 'tx'), 'v')[0] || '',
        labels: values(child(s, CHART_NS, 'cat') || child(s, CHART_NS, 'xVal'), 'v'),
        values: values(child(s, CHART_NS, 'val') || child(s, CHART_NS, 'yVal'), 'v').map(Number) })),
      workbooks: rels(parts, chartPart).filter(r => !r.external && r.type.endsWith('/package')).map(r => { const p = resolvePart(chartPart, r.target); return { part: p, sha256: parts.has(p) ? hash(parts.get(p)) : null }; }) } };
  }
  if (diagramRef) {
    const relation = relationship(diagramRef.getAttributeNS(NS.r, 'dm'));
    if (!relation) return { kind: 'smartart', smartart: { unavailable: true } };
    const dataPart = resolvePart(part, relation.target), doc = parseXml(parts.get(dataPart));
    return { kind: 'smartart', smartart: { part: dataPart, sha256: hash(parts.get(dataPart)),
      nodes: descendants(doc, DIAGRAM_NS, 'pt').filter(n => !n.getAttribute('type') || n.getAttribute('type') === 'node').map(n => ({ id: n.getAttribute('modelId'), text: descendants(n, NS.a, 't').map(t => t.textContent).join('\n') })) } };
  }
  return {};
}
