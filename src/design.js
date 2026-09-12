import { check } from './errors.js';

// Original palettes following restrained typography, whitespace and accent hierarchy.
export const designThemes = {
  executive: { background: 'F5F7FB', foreground: '152D4A', accent: '197C80', colors: ['197C80', '5478A3', 'E0A54B', 'AFC8DA'], fontFace: 'Microsoft YaHei' },
  editorial: { background: 'FAF6EF', foreground: '332C29', accent: 'AD493C', colors: ['AD493C', 'DCA65F', '657B68', '9E8E82'], fontFace: 'Microsoft YaHei' },
  academic: { background: 'FFFFFF', foreground: '172B4D', accent: '355DA8', colors: ['355DA8', '368E87', 'B88540', '8C6BB1'], fontFace: 'Microsoft YaHei' }
};

export function applyDesignTheme(spec) {
  if (!spec.theme || !designThemes[spec.theme]) return spec;
  const theme = designThemes[spec.theme];
  return { ...spec, slides: spec.slides?.map(page => ({ ...page, background: page.background ?? theme.background,
    items: page.items?.map(item => ({ ...(item.type === 'chart' ? { colors: theme.colors, fontFace: theme.fontFace } :
      item.type === 'text' ? { color: theme.foreground, fontFace: theme.fontFace } :
      item.type === 'shape' ? { color: theme.foreground, line: theme.accent, fontFace: theme.fontFace } : {}), ...item })) })) };
}

export function planLayout(snapshot, options) {
  const refs = [...(options.titleRef ? [options.titleRef] : []), ...options.targetRefs];
  check(new Set(refs).size === refs.length, 'DUPLICATE_TARGET', 'Layout targets must be unique.');
  const chosen = refs.map(ref => {
    const object = snapshot.objects.find(o => o.targetRef === ref);
    check(object, 'TARGET_CHANGED', 'Inspect fresh layout references before planning.');
    check(object.slide === options.slide && object.capabilities.includes('set_geometry') && !object.groupPath.length,
      'LAYOUT_TARGET_UNSUPPORTED', 'Layout requires top-level movable objects on one slide.');
    return object;
  });
  const { width, height } = snapshot, { margin, gap } = options;
  const top = margin + (options.titleRef ? 60 + gap : 0);
  const columns = options.layout === 'stack' ? 1 : 2;
  const rows = Math.ceil(options.targetRefs.length / columns);
  const cellWidth = (width - margin * 2 - gap * (columns - 1)) / columns;
  const cellHeight = (height - top - margin - gap * (rows - 1)) / rows;
  check(cellWidth >= 80 && cellHeight >= 50, 'LAYOUT_TOO_DENSE', 'Not enough room for readable layout. Select fewer objects or another layout.');
  const operations = [], placements = [], warnings = [];
  for (const [i, object] of chosen.entries()) {
    const title = !!options.titleRef && i === 0, index = i - (options.titleRef ? 1 : 0);
    const cell = title ? { left: margin, top: margin, width: width - margin * 2, height: 60 } :
      { left: margin + (index % columns) * (cellWidth + gap), top: top + Math.floor(index / columns) * (cellHeight + gap), width: cellWidth, height: cellHeight };
    // Office chart/table/SmartArt frames do not expose ordinary shape rotation.
    let geometry = { ...cell, ...(!['chart', 'table', 'smartart'].includes(object.kind) ? { rotation: 0 } : {}) };
    if (['image', 'smartart'].includes(object.kind)) {
      check(object.geometry.width > 0 && object.geometry.height > 0, 'LAYOUT_TARGET_UNSUPPORTED', 'Object needs explicit dimensions.');
      const scale = Math.min(cell.width / object.geometry.width, cell.height / object.geometry.height);
      geometry.width = object.geometry.width * scale; geometry.height = object.geometry.height * scale;
      geometry.left += (cell.width - geometry.width) / 2; geometry.top += (cell.height - geometry.height) / 2;
    }
    operations.push({ type: 'set_geometry', targetRef: object.targetRef, geometry });
    if (object.kind === 'text' && options.theme) {
      const theme = designThemes[options.theme], fontSize = title ? 32 : options.fontSize;
      operations.push({ type: 'set_style', targetRef: object.targetRef, style: { fontFace: theme.fontFace, fontSize, color: theme.foreground, bold: title } });
    }
    const fontSize = title ? 32 : options.fontSize;
    if (object.kind === 'text') {
      const lines = object.text.split('\n').reduce((sum, paragraph) => sum + Math.max(1, Math.ceil(paragraph.length * fontSize / Math.max(1, geometry.width))), 0);
      if (lines * fontSize * 1.35 > geometry.height) warnings.push({ code: 'TEXT_OVERFLOW_RISK', name: object.name, message: 'Text may not fit; content was not shortened.' });
    }
    placements.push({ name: object.name, targetRef: object.targetRef, before: object.geometry, after: geometry });
  }
  // Unselected diagrams and decorations can overlap intentionally; disclose them.
  for (const object of snapshot.objects.filter(o => o.slide === options.slide && !refs.includes(o.targetRef) && o.geometry && !o.groupPath.length)) {
    if (placements.some(p => p.after.left < object.geometry.left + object.geometry.width && p.after.left + p.after.width > object.geometry.left && p.after.top < object.geometry.top + object.geometry.height && p.after.top + p.after.height > object.geometry.top))
      warnings.push({ code: 'UNSELECTED_OBJECT_OVERLAP', name: object.name });
  }
  return { operations, placements, warnings, visualReview: 'required' };
}
