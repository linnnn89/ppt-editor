import path from 'node:path';
import { hash, stableJson, readJson, writeJson } from './storage.js';
import { layoutPageHashes } from './file-engine.js';
import { auditLayout, summarizeLayoutCoverage } from './layout-audit.js';

export const readOptional = file => readJson(file).catch(error => { if (error.code === 'ENOENT') return null; throw error; });

export function auditSummary(record, revision = record?.revision, generation = record?.generation) {
  if (!record) return { status: 'not_checked', nextAction: 'check_layout', pages: [], finalWholeDeckCheck: 'required' };
  const sameRevision = record.revision === revision && record.generation === generation;
  const pages = record.pages.map(p => ({ slide: p.slide, status: p.status,
    stale: !sameRevision || p.stale, counts: p.counts, checkedRevision: p.checkedRevision, coverage: p.coverage }));
  const pendingSlides = pages.filter(p => p.stale || p.status !== 'clear');
  const reviewNowSlides = pages.filter(p => p.counts?.overlaps > 0).map(p => p.slide);
  const finalCurrent = sameRevision && record.wholeDeckCheck?.revision === revision && record.wholeDeckCheck?.generation === generation;
  return { status: pendingSlides.length ? 'attention' : 'clear', pages, pendingSlides, reviewNowSlides,
    finalWholeDeckCheck: finalCurrent ? 'current' : 'required',
    nextAction: reviewNowSlides.length ? 'review_overlap_before_next_batch' : pendingSlides.length ? 'repair_or_check_listed_slides' : 'continue_batch_or_review_final_deck' };
}

export function compactAudit(audit) {
  const { pages: workflowPages, pendingSlides: workflowPending, ...workflow } = audit.workflow || {};
  const { allowedOverlapPairs: declared, ...result } = audit;
  return { ...result, workflow, pages: audit.pages.map(({ allowedOverlapPairs, ...page }) => {
    if (!page.issueCount) { delete page.issues; delete page.issuesTruncated; }
    return page;
  }) };
}

export async function recordAudit(taskDir, session, snapshot, options = {}) {
  const destination = path.join(taskDir, 'layout-audits', `${session.documentId}.json`);
  const previous = await readOptional(destination);
  const backend = snapshot.objects.some(o => o.geometry?.source === 'native-effective') ? 'native' : 'file';
  const hashes = session.mode === 'file' ? layoutPageHashes(session.engine.parts) :
    new Map(snapshot.slides.map(s => [s.slide, hash(stableJson([snapshot.width, snapshot.height,
      snapshot.objects.filter(o => o.slide === s.slide).map(o => [o.key, o.fingerprint])]))]));
  const selected = options.slides || snapshot.slides.map(s => s.slide);
  const reusable = (previous?.pages || []).filter(p => selected.includes(p.slide) && p.pageHash === hashes.get(p.slide) &&
    p.checkedGeneration === session.generation && p.backend === backend).flatMap(p => p.allowedOverlapPairs || []);
  const previousById = new Map((previous?.pages || []).map(page => [page.slideId, page]));
  const previousTextPages = new Map(), overflowBySlide = new Map();
  for (const slide of snapshot.slides) {
    const page = previousById.get(slide.slideId);
    if (page && page.checkedGeneration === session.generation && page.pageHash === hashes.get(slide.slide) &&
        (page.coverage?.textFitDetails?.measured || page.counts?.textOverflow)) previousTextPages.set(slide.slide, page);
  }
  const audit = auditLayout(snapshot, { ...options, allowedOverlapPairs: options.allowedOverlapPairs ?? reusable },
    { previousTextPages, fileCheck: backend === 'file', overflowBySlide });
  const checkedAt = new Date().toISOString(), checked = new Map(audit.pages.map(p => [p.slide, p]));
  const pages = snapshot.slides.map(s => {
    const current = checked.get(s.slide);
    if (current) return { ...current, slideId: s.slideId, pageHash: hashes.get(s.slide), backend,
      checkedRevision: session.revision, checkedGeneration: session.generation, checkedAt, stale: false,
      ...(overflowBySlide.has(s.slide) ? { textOverflowIssues: overflowBySlide.get(s.slide) } : {}) };
    const old = previousById.get(s.slideId);
    if (!old) return { slide: s.slide, slideId: s.slideId, status: 'not_checked', stale: true };
    // File sessions use the same package dependency hashes for both audit sources.
    const compatibleBasis = session.mode === 'file' || old.backend === backend;
    return { ...old, slide: s.slide, stale: old.pageHash !== hashes.get(s.slide) || old.checkedGeneration !== session.generation || !compatibleBasis };
  });
  const basis = { taskId: session.taskId, documentId: session.documentId, revision: session.revision,
    generation: session.generation, checkedAt, basis: 'page-content-and-layout-dependencies' };
  const wholeDeckCheck = options.slides ? previous?.wholeDeckCheck : { revision: session.revision, generation: session.generation, checkedAt };
  const record = { ...basis, coverage: summarizeLayoutCoverage(pages), wholeDeckCheck, pages };
  await writeJson(destination, record);
  const workflow = auditSummary(record);
  const result = { ...audit, ...basis, worklistPath: destination, pendingSlides: workflow.pendingSlides,
    workflow, nextAction: workflow.nextAction, screenshotRequiredNow: workflow.reviewNowSlides.length > 0 };
  return options.detail === 'summary' ? compactAudit(result) : result;
}

export function summarizeNative(native) {
  const { snapshot, ...summary } = native;
  return { ...summary, ...(snapshot ? { objectCount: snapshot.objects.length } : {}) };
}

export async function assessReview(taskDir, review) {
  let record = await readOptional(path.join(taskDir, 'layout-audits', `${review.documentId}.json`));
  // An older server could already have erased measured overflow from the
  // worklist. The committed preview still records it for these exact bytes.
  const saved = review.preview?.layoutAudit;
  if (record && saved && review.preview.outputHash === review.outputHash &&
      saved.revision === record.revision && saved.generation === record.generation) {
    const overflows = new Map(saved.pages.filter(p => p.counts.textOverflow).map(p => [p.slide, p]));
    const pages = record.pages.map(page => {
      const measured = overflows.get(page.slide);
      if (!measured || page.backend !== 'file' || page.coverage?.textFitDetails?.measured || page.counts?.textOverflow) return page;
      return { ...page, status: 'issues', counts: { ...page.counts, textOverflow: measured.counts.textOverflow }, coverage: measured.coverage };
    });
    record = { ...record, pages, coverage: summarizeLayoutCoverage(pages) };
  }
  const layout = auditSummary(record, review.revision, review.generation);
  const slideCount = review.preview?.nativeValidation?.slidesCount;
  const previewed = new Set((review.preview?.images || []).map(i => i.slide));
  const missingPreviewSlides = slideCount ? Array.from({ length: slideCount }, (_, i) => i + 1).filter(i => !previewed.has(i)) : [];
  const visualEvidenceComplete = !!slideCount && missingPreviewSlides.length === 0 && review.preview.outputHash === review.outputHash;
  const reasons = [];
  if (layout.finalWholeDeckCheck !== 'current') reasons.push('FINAL_LAYOUT_CHECK_REQUIRED');
  const layoutBlocked = !record || layout.pages.some(p => p.stale || p.status === 'not_checked' || p.status === 'issues');
  if (layoutBlocked) reasons.push('LAYOUT_ATTENTION_REQUIRED');
  if (!visualEvidenceComplete) reasons.push('VISUAL_EVIDENCE_INCOMPLETE');
  return { status: reasons.length ? 'incomplete' : 'accepted', reasons, visualEvidenceComplete, missingPreviewSlides,
    visualFallbackSlides: layout.pages.filter(p => p.status === 'incomplete').map(p => p.slide),
    coverage: record?.coverage || null,
    layout: { status: layout.status, finalWholeDeckCheck: layout.finalWholeDeckCheck, pendingSlides: layout.pendingSlides || [] },
    outputHash: review.outputHash, basis: 'checked-output-and-caller-confirmation' };
}
