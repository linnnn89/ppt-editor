import { readFile } from 'node:fs/promises';
import { buildDeck } from '../../src/build.js';
import { readPackage, writePackage, parseXml, descendants, NS, xml } from '../../src/ooxml.js';

// MIT-licensed Open XML SDK test diagram, with no source document metadata,
// thumbnail, notes, or animations. See smartart/LICENSE.txt and README.md.
export async function smartArtSeed() {
  const parts = await readPackage(await buildDeck({ slides: [{ items: [] }] }));
  const source = name => readFile(new URL(`./smartart/${name}`, import.meta.url));
  const slide = parseXml(parts.get('ppt/slides/slide1.xml')), frame = parseXml(await source('frame.xml'));
  descendants(slide, NS.p, 'spTree')[0].appendChild(slide.importNode(frame.documentElement, true));
  const relations = parseXml(parts.get('ppt/slides/_rels/slide1.xml.rels')), types = parseXml(parts.get('[Content_Types].xml'));
  const diagramNs = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
  const diagram = descendants(slide, diagramNs, 'relIds')[0];
  for (const [attr, kind, file, contentType] of [['dm','diagramData','data1.xml','data'],['lo','diagramLayout','layout1.xml','layout'],['qs','diagramQuickStyle','quickStyle1.xml','style'],['cs','diagramColors','colors1.xml','colors']]) {
    const id = 'rIdSeed_' + attr, part = 'ppt/diagrams/' + file;
    diagram.setAttributeNS(NS.r,'r:'+attr,id); parts.set(part,await source(file));
    const rel = relations.createElementNS(NS.rel,'Relationship');rel.setAttribute('Id',id);rel.setAttribute('Type',NS.r+'/'+kind);rel.setAttribute('Target','../diagrams/'+file);relations.documentElement.appendChild(rel);
    const type = types.createElementNS(NS.ct,'Override');type.setAttribute('PartName','/'+part);type.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.drawingml.diagram'+contentType.charAt(0).toUpperCase()+contentType.slice(1)+'+xml');types.documentElement.appendChild(type);
  }
  parts.set('ppt/slides/slide1.xml',Buffer.from(xml(slide)));parts.set('ppt/slides/_rels/slide1.xml.rels',Buffer.from(xml(relations)));parts.set('[Content_Types].xml',Buffer.from(xml(types)));
  return writePackage(parts);
}
