import path from 'node:path';
import { randomUUID } from 'node:crypto';
import { check } from './errors.js';
import { NS, readPackage, writePackage, validatePackage, checkNativeSafe, parseXml, xml, child, children, descendants, rels, relationshipsPath, resolvePart, setChild } from './ooxml.js';
import { styleBody } from './file-engine.js';

// Copy complete OPC dependencies, preserving their relationship IDs. Only new
// parts are rewritten; existing template slides and their dependencies stay intact.
export async function composeDeck(templateBytes, contentBytes, options) {
  const target = await readPackage(templateBytes), source = await readPackage(contentBytes);
  await checkNativeSafe(target); await checkNativeSafe(source); validatePackage(target); validatePackage(source);
  const presentation = parseXml(target.get('ppt/presentation.xml'));
  const sourcePresentation = parseXml(source.get('ppt/presentation.xml'));
  const slideParts = (parts, doc) => {
    const relationships = rels(parts, 'ppt/presentation.xml');
    return descendants(doc, NS.p, 'sldId').map(n => resolvePart('ppt/presentation.xml', relationships.find(r => r.id === n.getAttributeNS(NS.r, 'id')).target));
  };
  const templateSlides = slideParts(target, presentation), sourceSlides = slideParts(source, sourcePresentation);
  check(options.templateSlide <= templateSlides.length, 'SLIDE_NOT_FOUND', 'Template reference page is out of range.');
  const selected = options.sourceSlides || sourceSlides.map((_, i) => i + 1);
  check(selected.length && selected.every(i => i <= sourceSlides.length) && new Set(selected).size === selected.length, 'INVALID_SLIDE_SELECTION', 'Select unique existing source pages.');
  check(templateSlides.length + selected.length <= 100, 'TOO_MANY_SLIDES', 'Composition is limited to 100 output pages.');
  const templatePart = templateSlides[options.templateSlide - 1], templateDoc = parseXml(target.get(templatePart));
  const templateLayoutRel = rels(target, templatePart).find(r => r.type.endsWith('/slideLayout'));
  check(templateLayoutRel, 'INVALID_TEMPLATE', 'Template reference page requires a slide layout.');
  const templateLayout = resolvePart(templatePart, templateLayoutRel.target);
  const types = parseXml(target.get('[Content_Types].xml')), sourceTypes = parseXml(source.get('[Content_Types].xml'));
  const typeOf = part => descendants(sourceTypes, NS.ct, 'Override').find(n => n.getAttribute('PartName') === '/' + part)?.getAttribute('ContentType') || descendants(sourceTypes, NS.ct, 'Default').find(n => n.getAttribute('Extension') === part.split('.').at(-1))?.getAttribute('ContentType');
  const destinationRels = parseXml(target.get('ppt/_rels/presentation.xml.rels'));
  const prefix = `import-${randomUUID().slice(0, 8)}-`, mapped = new Map(), warnings = [];
  function copyPart(part) {
    if (mapped.has(part)) return mapped.get(part);
    check(source.has(part), 'INVALID_RELATIONSHIP', 'Source dependency is missing.');
    const result = path.posix.join(path.posix.dirname(part), prefix + path.posix.basename(part)); mapped.set(part, result);
    target.set(result, source.get(part));
    const override = types.createElementNS(NS.ct, 'Override'); override.setAttribute('PartName', '/' + result);
    override.setAttribute('ContentType', typeOf(part)); types.documentElement.appendChild(override);
    const relationPart = relationshipsPath(part);
    if (source.has(relationPart)) {
      const doc = parseXml(source.get(relationPart));
      for (const relation of descendants(doc, NS.rel, 'Relationship')) {
        if (relation.getAttribute('TargetMode') === 'External') continue;
        const dependency = resolvePart(part, relation.getAttribute('Target'));
        const isSourcePage = sourceSlides.includes(part);
        if (isSourcePage && relation.getAttribute('Type').endsWith('/slideLayout')) {
          relation.setAttribute('Target', path.posix.relative(path.posix.dirname(result), templateLayout)); continue;
        }
        check(!relation.getAttribute('Type').endsWith('/slide') || mapped.has(dependency), 'CROSS_SLIDE_LINK_UNSUPPORTED', 'Cross-page source links require manual review before template composition.');
        relation.setAttribute('Target', path.posix.relative(path.posix.dirname(result), copyPart(dependency)));
      }
      target.set(relationshipsPath(result), Buffer.from(xml(doc)));
    }
    return result;
  }
  const sizeOf = doc => { const size = descendants(doc, NS.p, 'sldSz')[0]; return { width: Number(size.getAttribute('cx')), height: Number(size.getAttribute('cy')) }; };
  const toSize = sizeOf(presentation), fromSize = sizeOf(sourcePresentation), margin = options.margin * 12700;
  check(toSize.width > 2 * margin && toSize.height > 2 * margin, 'INVALID_MARGIN', 'Margins leave no space for imported content.');
  const scale = Math.min((toSize.width - 2 * margin) / fromSize.width, (toSize.height - 2 * margin) / fromSize.height);
  const dx = (toSize.width - fromSize.width * scale) / 2, dy = (toSize.height - fromSize.height * scale) / 2;
  const templateTree = descendants(templateDoc, NS.p, 'spTree')[0];
  const decorative = children(templateTree, NS.p).filter(n => ['sp', 'pic', 'grpSp'].includes(n.localName) && !descendants(n, NS.p, 'ph').length && !descendants(n, NS.a, 't').some(t => t.textContent.trim()));
  const explicitColor = descendants(templateDoc, NS.a, 'rPr').map(p => child(child(p, NS.a, 'solidFill'), NS.a, 'srgbClr')?.getAttribute('val')).find(Boolean);
  const explicitFont = descendants(templateDoc, NS.a, 'latin').map(n => n.getAttribute('typeface')).find(n => n && !n.startsWith('+'));
  const titleStyle = { fontFace: options.fontFace || explicitFont || 'Microsoft YaHei', fontSize: options.titleSize, color: options.textColor || explicitColor || '17365D', bold: true };
  const bodyStyle = { ...titleStyle, fontSize: options.bodySize, bold: false };
  if (options.beautify && !options.textColor && !explicitColor) warnings.push({ code: 'TEXT_COLOR_FALLBACK', message: 'Template has no explicit text color; verify contrast or supply textColor.' });
  let nextSlideId = Math.max(255, ...descendants(presentation, NS.p, 'sldId').map(n => Number(n.getAttribute('id')))) + 1;
  const usedRels = new Set(descendants(destinationRels, NS.rel, 'Relationship').map(n => n.getAttribute('Id')));
  const appended = [];
  for (const number of selected) {
    const sourcePart = sourceSlides[number - 1], part = copyPart(sourcePart), doc = parseXml(target.get(part));
    check(!descendants(doc, NS.p, 'timing').length, 'ANIMATION_COMPOSITION_UNSUPPORTED', 'Animated source slides require a separate workflow; this importer preserves static content.');
    const tree = descendants(doc, NS.p, 'spTree')[0], common = child(doc.documentElement, NS.p, 'cSld');
    doc.documentElement.setAttribute('showMasterSp', '1');
    const sourceLayout = resolvePart(sourcePart, rels(source, sourcePart).find(r => r.type.endsWith('/slideLayout')).target);
    const layoutDoc = parseXml(source.get(sourceLayout));
    const masterRelation = rels(source, sourceLayout).find(r => r.type.endsWith('/slideMaster'));
    const masterDoc = masterRelation ? parseXml(source.get(resolvePart(sourceLayout, masterRelation.target))) : null;
    for (const shape of children(tree, NS.p).filter(n => ['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp'].includes(n.localName))) {
      const ph = descendants(shape, NS.p, 'ph')[0];
      let isTitle = ph && ['title', 'ctrTitle'].includes(ph.getAttribute('type'));
      let transform = descendants(shape, NS.a, 'xfrm')[0] || child(shape, NS.p, 'xfrm');
      if (!transform && ph) {
        const index = ph.getAttribute('idx') || '0', type = ph.getAttribute('type') || 'body';
        for (const fallback of [layoutDoc, masterDoc].filter(Boolean)) {
          const inherited = descendants(fallback, NS.p, 'sp').find(n => { const p = descendants(n, NS.p, 'ph')[0]; return p && ((p.getAttribute('idx') || '0') === index || (fallback === masterDoc && p.getAttribute('type') === type)); });
          if (inherited) {
            const inheritedType = descendants(inherited, NS.p, 'ph')[0]?.getAttribute('type');
            isTitle ||= ['title', 'ctrTitle'].includes(inheritedType);
            const inheritedTransform = descendants(inherited, NS.a, 'xfrm')[0];
            if (inheritedTransform) { transform = doc.importNode(inheritedTransform, true); const props = child(shape, NS.p, 'spPr'); check(props, 'IMPLICIT_GEOMETRY_UNSUPPORTED', 'Cannot materialize source placeholder geometry.'); props.insertBefore(transform, props.firstChild); break; }
          }
        }
      }
      check(transform, 'IMPLICIT_GEOMETRY_UNSUPPORTED', 'Source object has no resolvable geometry. Use explicit placement before importing.');
      const off = child(transform, NS.a, 'off'), ext = child(transform, NS.a, 'ext');
      check(off && ext, 'IMPLICIT_GEOMETRY_UNSUPPORTED', 'Source object geometry is incomplete.');
      const originalWidth = Number(ext.getAttribute('cx')), originalHeight = Number(ext.getAttribute('cy'));
      if (originalWidth > fromSize.width * 0.9 && originalHeight > fromSize.height * 0.9) warnings.push({ code: 'LARGE_SOURCE_OBJECT', sourceSlide: number, message: 'A large source object may cover the template background; it was preserved.' });
      off.setAttribute('x', String(Math.round(dx + Number(off.getAttribute('x')) * scale))); off.setAttribute('y', String(Math.round(dy + Number(off.getAttribute('y')) * scale)));
      ext.setAttribute('cx', String(Math.round(originalWidth * scale))); ext.setAttribute('cy', String(Math.round(originalHeight * scale)));
      // Detach placeholders only after materializing their inherited placement.
      if (ph) ph.parentNode.removeChild(ph);
      if (options.beautify) for (const body of [...descendants(shape, NS.p, 'txBody'), ...descendants(shape, NS.a, 'txBody')]) {
        const style = isTitle ? titleStyle : bodyStyle; styleBody(body, style);
        const characters = descendants(body, NS.a, 't').reduce((n, t) => n + t.textContent.length, 0);
        const estimatedLines = characters * style.fontSize * 0.65 / Math.max(1, originalWidth * scale / 12700);
        if (estimatedLines * style.fontSize * 1.3 > originalHeight * scale / 12700) warnings.push({ code: 'TEXT_OVERFLOW_RISK', sourceSlide: number, message: 'Text may overflow after uniform styling; content was retained. Inspect the preview.' });
      }
    }
    const oldBg = child(common, NS.p, 'bg'); if (oldBg) common.removeChild(oldBg);
    const pageRelsPart = relationshipsPath(part), pageRels = parseXml(target.get(pageRelsPart));
    const ids = new Set(descendants(pageRels, NS.rel, 'Relationship').map(n => n.getAttribute('Id')));
    const importTemplateNode = node => {
      const copy = doc.importNode(node, true);
      for (const element of [copy, ...Array.from(copy.getElementsByTagName('*'))]) for (const attribute of Array.from(element.attributes || [])) if (attribute.namespaceURI === NS.r) {
        const relationship = rels(target, templatePart).find(r => r.id === attribute.value);
        check(relationship, 'INVALID_TEMPLATE', 'Template decoration has a missing relationship.');
        let n = 1; while (ids.has(`rIdTemplate${n}`)) n++; const id = `rIdTemplate${n}`; ids.add(id);
        const rel = pageRels.createElementNS(NS.rel, 'Relationship'); rel.setAttribute('Id', id); rel.setAttribute('Type', relationship.type);
        rel.setAttribute('Target', relationship.external ? relationship.target : path.posix.relative(path.posix.dirname(part), resolvePart(templatePart, relationship.target)));
        if (relationship.external) rel.setAttribute('TargetMode', 'External'); pageRels.documentElement.appendChild(rel);
        element.setAttributeNS(NS.r, attribute.name, id);
      }
      return copy;
    };
    const bg = child(child(templateDoc.documentElement, NS.p, 'cSld'), NS.p, 'bg'); if (bg) common.insertBefore(importTemplateNode(bg), tree);
    if (options.copyDecorations) {
      let shapeId = Math.max(1, ...descendants(doc, NS.p, 'cNvPr').map(n => Number(n.getAttribute('id')))) + 1;
      const firstContent = children(tree, NS.p).find(n => !['nvGrpSpPr', 'grpSpPr'].includes(n.localName));
      const decorationIds = new Map(), decorationCopies = [];
      for (const original of decorative) {
        const copy = importTemplateNode(original);
        for (const id of descendants(copy, NS.p, 'cNvPr')) { const next = String(shapeId++); decorationIds.set(id.getAttribute('id'), next); id.setAttribute('id', next); id.setAttribute('name', 'Template_' + id.getAttribute('name')); }
        decorationCopies.push(copy);
        tree.insertBefore(copy, firstContent || null);
      }
      for (const copy of decorationCopies) for (const connection of [...descendants(copy, NS.a, 'stCxn'), ...descendants(copy, NS.a, 'endCxn')]) {
        const id = decorationIds.get(connection.getAttribute('id'));
        if (id) connection.setAttribute('id', id); else connection.parentNode.removeChild(connection);
      }
    }
    target.set(part, Buffer.from(xml(doc))); target.set(pageRelsPart, Buffer.from(xml(pageRels)));
    let n = 1; while (usedRels.has(`rIdImported${n}`)) n++; const id = `rIdImported${n}`; usedRels.add(id);
    const relation = destinationRels.createElementNS(NS.rel, 'Relationship'); relation.setAttribute('Id', id); relation.setAttribute('Type', `${NS.r}/slide`); relation.setAttribute('Target', path.posix.relative('ppt', part)); destinationRels.documentElement.appendChild(relation);
    const slideId = presentation.createElementNS(NS.p, 'p:sldId'); slideId.setAttribute('id', String(nextSlideId++)); slideId.setAttributeNS(NS.r, 'r:id', id); child(presentation.documentElement, NS.p, 'sldIdLst').appendChild(slideId);
    appended.push({ sourceSlide: number, outputSlide: templateSlides.length + appended.length + 1 });
  }
  target.set('ppt/presentation.xml', Buffer.from(xml(presentation))); target.set('ppt/_rels/presentation.xml.rels', Buffer.from(xml(destinationRels))); target.set('[Content_Types].xml', Buffer.from(xml(types)));
  for (const name of ['docProps/core.xml', 'docProps/app.xml']) if (target.has(name)) {
    const doc = parseXml(target.get(name));
    for (const n of Array.from(doc.getElementsByTagName('*'))) if (['creator', 'lastModifiedBy', 'Company', 'Manager'].includes(n.localName)) n.parentNode.removeChild(n);
    target.set(name, Buffer.from(xml(doc)));
  }
  validatePackage(target); await checkNativeSafe(target);
  return { bytes: await writePackage(target), report: { templateSlidesPreserved: templateSlides.length, appended, templateSlide: options.templateSlide,
    backgroundPolicy: 'reference-page-background-and-layout', decorationPolicy: options.copyDecorations ? 'non-placeholder-non-text-shapes' : 'master-and-layout-only',
    beautify: options.beautify, style: options.beautify ? { title: titleStyle, body: bodyStyle } : null, scale, warnings, visualReview: 'required' } };
}
