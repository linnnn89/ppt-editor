import { check } from './errors.js';

const coverageFor = textFitDetails => ({
  objectBounds: true, objectOverlap: 'possible-overlap-only',
  textFit: textFitDetails.measured ? 'native-text-bounds-partial' : 'not_checked', textFitDetails,
  chartAndSmartArtInternals: 'not_checked', groupChildCoordinates: 'unresolved', strokeAndEffects: 'not_checked'
});

export function summarizeLayoutCoverage(pages) {
  const details = { measured: 0, unmeasured: 0, reasons: {} };
  for (const page of pages) {
    const counts = page.coverage?.textFitDetails;
    // Stale pages and pre-count worklists cannot establish current measurements.
    if (page.stale || !counts) { details.unknownPageCount = (details.unknownPageCount || 0) + 1; continue; }
    details.measured += counts.measured; details.unmeasured += counts.unmeasured;
    for (const [reason, count] of Object.entries(counts.reasons)) details.reasons[reason] = (details.reasons[reason] || 0) + count;
  }
  return coverageFor(details);
}

// Object-frame checks only. Text layout, strokes, transparency and intentional
// layering need final review; a rectangle intersection is not proof of occlusion.
export function auditLayout(snapshot, { slides, tolerancePt = 0.5, allowedOverlapPairs = [] } = {}) {
  check(Number.isFinite(snapshot.width) && snapshot.width > 0 && Number.isFinite(snapshot.height) && snapshot.height > 0,
    'LAYOUT_SIZE_UNAVAILABLE', 'Page dimensions are unavailable.');
  const selected = slides || snapshot.slides.map(s => s.slide);
  check(new Set(selected).size === selected.length, 'DUPLICATE_SLIDE', 'Layout check slides must be unique.');
  const available = new Set(snapshot.slides.map(s => s.slide));
  const objectsByKey = new Map(snapshot.objects.filter(o => selected.includes(o.slide) && !['background','notes'].includes(o.kind)).map(o => [o.key,o]));
  const pairKey = pair => JSON.stringify([...pair].sort());
  const allowed = new Set();
  for (const pair of allowedOverlapPairs) {
    const [a,b] = pair.map(key => objectsByKey.get(key));
    check(a && b && a.key !== b.key && a.slide === b.slide, 'INVALID_OVERLAP_EXCEPTION', 'Design overlap exceptions must identify two distinct inspected objects on the same checked slide.');
    allowed.add(pairKey(pair));
  }
  const round = value => Math.round(value * 1000) / 1000;
  const pages = selected.map(slide => {
    check(available.has(slide), 'SLIDE_NOT_FOUND', 'Layout check slide is out of range.', { slide });
    const objects = snapshot.objects.filter(o => o.slide === slide && !['background', 'notes'].includes(o.kind));
    const boxes = [], issues = [], counts = { outOfBounds: 0, overlaps: 0, unresolved: 0, textOverflow: 0 };
    const textFitDetails = { measured: 0, unmeasured: 0, reasons: {} };
    let allowedOverlapCount = 0;
    const add = (type, issue) => { counts[type]++; if (issues.length < 100) issues.push(issue); };
    for (const object of objects) {
      const ref = { key: object.key, name: object.name, kind: object.kind };
      const g = object.geometry, t = object.textBounds;
      const grouped = !!object.groupPath?.length;
      const geometryResolved = g && ['left', 'top', 'width', 'height'].every(k => Number.isFinite(g[k])) &&
        g.width >= 0 && g.height >= 0 && Number.isFinite(g.rotation ?? 0);
      let measuredText = false;
      if (object.kind === 'text' && (object.text || t)) {
        const reason = grouped ? 'group_child' : !geometryResolved ? 'geometry_unresolved' :
          Math.abs(g.rotation || 0) >= 0.01 ? 'rotated_text' :
          !t ? g.source === 'native-effective' ? 'native_measurement_unavailable' : 'native_readback_not_run' :
          ![t.left, t.top, t.width, t.height].every(Number.isFinite) || t.width < 0 || t.height < 0 ? 'invalid_text_bounds' : null;
        if (reason) {
          textFitDetails.unmeasured++;
          textFitDetails.reasons[reason] = (textFitDetails.reasons[reason] || 0) + 1;
        } else { textFitDetails.measured++; measuredText = true; }
      }
      if (grouped || !geometryResolved) {
        add('unresolved', { code: 'GEOMETRY_UNRESOLVED', object: ref }); continue;
      }
      const angle = (g.rotation || 0) * Math.PI / 180;
      const width = Math.abs(g.width * Math.cos(angle)) + Math.abs(g.height * Math.sin(angle));
      const height = Math.abs(g.width * Math.sin(angle)) + Math.abs(g.height * Math.cos(angle));
      const left = g.left + (g.width - width) / 2, top = g.top + (g.height - height) / 2;
      const box = { object: ref, left, top, right: left + width, bottom: top + height, rotated: Math.abs(g.rotation || 0) % 180 !== 0 };
      const outside = { left: Math.max(0, -box.left), top: Math.max(0, -box.top),
        right: Math.max(0, box.right - snapshot.width), bottom: Math.max(0, box.bottom - snapshot.height) };
      if (Object.values(outside).some(v => v > tolerancePt)) add('outOfBounds', {
        code: 'OUT_OF_SLIDE', object: ref, overflowPt: Object.fromEntries(Object.entries(outside).map(([k,v]) => [k,round(v)]))
      });
      boxes.push(box);
      if (measuredText) {
        const overflow = { left: Math.max(0,g.left-t.left), top: Math.max(0,g.top-t.top),
          right: Math.max(0,t.left+t.width-g.left-g.width), bottom: Math.max(0,t.top+t.height-g.top-g.height) };
        if (Object.values(overflow).some(v => v > tolerancePt)) add('textOverflow', { code: 'TEXT_OUTSIDE_FRAME', object: ref,
          overflowPt: Object.fromEntries(Object.entries(overflow).map(([k,v]) => [k,round(v)])), basis: 'native-text-bounds' });
      }
    }
    for (let i = 0; i < boxes.length; i++) for (let j = i + 1; j < boxes.length; j++) {
      const a = boxes[i], b = boxes[j];
      const width = Math.min(a.right,b.right) - Math.max(a.left,b.left), height = Math.min(a.bottom,b.bottom) - Math.max(a.top,b.top);
      if (width > tolerancePt && height > tolerancePt && allowed.has(pairKey([a.object.key,b.object.key]))) { allowedOverlapCount++; continue; }
      if (width > tolerancePt && height > tolerancePt) add('overlaps', { code: 'POSSIBLE_OVERLAP',
        objects: [a.object,b.object], intersectionPt: { width: round(width), height: round(height) },
        basis: a.rotated || b.rotated ? 'conservative-rotated-bounds' : 'object-bounds', review: 'May be intentional layering; do not automatically move or delete content.' });
    }
    const issueCount = Object.values(counts).reduce((a,b) => a+b,0);
    return { slide, status: counts.outOfBounds || counts.overlaps || counts.textOverflow ? 'issues' : counts.unresolved ? 'incomplete' : 'clear',
      objectsChecked: boxes.length, allowedOverlapCount, allowedOverlapPairs: allowedOverlapPairs.filter(pair => objectsByKey.get(pair[0]).slide === slide),
      counts, issueCount, issuesTruncated: issueCount > issues.length, issues, coverage: coverageFor(textFitDetails) };
  });
  return { scope: slides ? 'selected-slides' : 'all-slides', tolerancePt, pages,
    summary: { slidesChecked: pages.length, clearSlides: pages.filter(p => p.status === 'clear').map(p => p.slide),
      attentionSlides: pages.filter(p => p.status !== 'clear').map(p => p.slide), issueCount: pages.reduce((sum,p) => sum+p.issueCount,0) },
    coverage: summarizeLayoutCoverage(pages),
    allowedOverlapPairs,
    nextAction: pages.some(p => p.counts.overlaps) ? 'review_overlap_before_continuing' : 'continue_drafting_then_review_all_slides',
    screenshotRequiredNow: pages.some(p => p.counts.overlaps) };
}
