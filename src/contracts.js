import { z } from 'zod';
import { backgroundPatterns } from './background.js';
import { designThemes } from './design.js';

const text = z.string().max(20000).refine(s => s.isWellFormed() && !/[\x00-\x08\x0B\x0C\x0E-\x1F]/.test(s), 'Text contains invalid XML characters.');
const shortText = z.string().min(1).max(300);
const operationId = z.string().min(1).max(100).regex(/^[a-zA-Z0-9_-]+$/);
const documentId = z.string().uuid();
const color = z.string().regex(/^[0-9a-fA-F]{6}$/);
const transparency = z.number().min(0).max(100).default(0);
export const backgroundFillSchema = z.discriminatedUnion('type', [
  z.strictObject({ type: z.literal('solid'), color, transparency }),
  z.strictObject({ type: z.literal('gradient'), startColor: color, endColor: color, angle: z.number().min(0).max(360).default(0), transparency }),
  z.strictObject({ type: z.literal('pattern'), pattern: z.enum(Object.keys(backgroundPatterns)), foreground: color, background: color, transparency }),
  z.strictObject({ type: z.literal('image'), path: z.string().min(1).max(32760), transparency }),
  z.strictObject({ type: z.literal('texture'), path: z.string().min(1).max(32760), transparency }),
  z.strictObject({ type: z.literal('inherit') })
]);
const geometry = z.strictObject({
  left: z.number().finite().min(-4000).max(4000).optional(),
  top: z.number().finite().min(-4000).max(4000).optional(),
  width: z.number().positive().max(4000).optional(),
  height: z.number().positive().max(4000).optional(),
  rotation: z.number().finite().min(-360).max(360).optional()
}).refine(v => Object.keys(v).length > 0, 'At least one geometry property is required.');
const style = z.strictObject({
  fontSize: z.number().min(1).max(400).optional(),
  fontFace: shortText.optional(),
  color: color.optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional()
}).refine(v => Object.keys(v).length > 0, 'At least one style property is required.');
const targetRef = z.string().min(20).max(6000);
export const operationSchema = z.discriminatedUnion('type', [
  z.strictObject({ type: z.literal('convert_to_smartart'), targetRef, layout: z.enum(['process', 'cycle', 'hierarchy']), }),
  z.strictObject({ type: z.literal('set_smartart_text'), targetRef, nodeIndex: z.number().int().min(1).max(100), text }),
  z.strictObject({ type: z.literal('replace_text'), targetRef, search: text.refine(s => s.length > 0 && !/[\r\n]/.test(s), 'Search must be nonempty and within one paragraph.'), replacement: text.refine(s => !/[\r\n]/.test(s), 'Replacement must stay within one paragraph.'), expectedMatches: z.number().int().min(1).max(1000), crossRunPolicy: z.enum(['reject', 'first_run']).default('reject') }),
  z.strictObject({ type: z.literal('set_style'), targetRef, style }),
  z.strictObject({ type: z.literal('set_geometry'), targetRef, geometry }),
  z.strictObject({ type: z.literal('set_table_cell'), targetRef, row: z.number().int().positive().max(1000), column: z.number().int().positive().max(1000), text, formatPolicy: z.literal('first_run') }),
  z.strictObject({ type: z.literal('set_slide_background'), targetRef, color: color.optional(), fill: backgroundFillSchema.optional() })
    .refine(v => (v.color !== undefined) !== (v.fill !== undefined), 'Specify exactly one of color or fill.')
]);

const itemGeometry = { left: z.number().finite(), top: z.number().finite(), width: z.number().positive().max(4000), height: z.number().positive().max(4000) };
const shapeGeometry = {
  left: z.number().finite(),
  top: z.number().finite(),
  width: z.number().nonnegative().max(4000),
  height: z.number().nonnegative().max(4000)
};
const slideItem = z.discriminatedUnion('type', [
  z.strictObject({ type: z.literal('chart'), ...itemGeometry, name: shortText.optional(),
    chartType: z.enum(['bar', 'line', 'area', 'pie', 'doughnut', 'radar', 'combo']),
    series: z.array(z.strictObject({ name: shortText, labels: z.array(shortText).min(1).max(100), values: z.array(z.number().finite()).min(1).max(100),
      type: z.enum(['bar', 'line', 'area']).optional(), secondaryAxis: z.boolean().optional() })).min(1).max(12),
    colors: z.array(color).min(1).max(12).optional(),
    fontFace: shortText.optional(), title: text.default(''), showLegend: z.boolean().default(true),
    showValue: z.boolean().default(false), stacked: z.boolean().default(false), horizontal: z.boolean().default(false)
  }).superRefine((v, ctx) => {
    const invalid = message => ctx.addIssue({ code: 'custom', message });
    if (v.series.some(s => s.labels.length !== s.values.length || JSON.stringify(s.labels) !== JSON.stringify(v.series[0].labels))) invalid('Chart series require matching categories and equal label/value counts.');
    if (['pie','doughnut'].includes(v.chartType) && (v.series.length !== 1 || v.series[0].values.some(n => n < 0) || !v.series[0].values.some(n => n > 0))) invalid('Pie/doughnut requires one nonnegative series with a positive total.');
    if (v.chartType === 'combo' && (v.series.length < 2 || v.series.some(s => !s.type) || !v.series.some(s => !s.secondaryAxis))) invalid('Combo requires two series, explicit types and a primary-axis series.');
    if (v.chartType !== 'combo' && v.series.some(s => s.type || s.secondaryAxis)) invalid('Per-series types/axes require combo.');
    if (v.horizontal && v.chartType !== 'bar') invalid('Horizontal is supported for bar charts only.');
    if (v.stacked && !['bar','area'].includes(v.chartType)) invalid('Stacked is supported for bar or area charts only.');
  }),
  z.strictObject({ type: z.literal('text'), text, ...itemGeometry, fontSize: z.number().min(1).max(400).default(24), fontFace: shortText.optional(), color: color.optional(), bold: z.boolean().default(false), italic: z.boolean().optional(), underline: z.boolean().optional(), name: shortText.optional() }),
  z.strictObject({
    type: z.literal('shape'),
    shape: z.enum([
      'rect', 'roundRect', 'ellipse', 'triangle', 'rtTriangle',
      'diamond', 'line', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow'
    ]),
    ...shapeGeometry,
    flipV: z.boolean().optional(),
    flipH: z.boolean().optional(),
    rotation: z.number().finite().min(-360).max(360).optional(),
    fill: color.optional(),
    line: color.optional(),
    lineWidth: z.number().positive().max(100).default(1),
    arrow: z.enum(['none', 'end', 'begin', 'both']).default('none'),
    text: text.optional(),
    fontSize: z.number().min(1).max(400).default(14),
    fontFace: shortText.optional(),
    color: color.optional(),
    bold: z.boolean().default(false),
    italic: z.boolean().optional(),
    underline: z.boolean().optional(),
    align: z.enum(['left', 'center', 'right']).default('center'),
    valign: z.enum(['top', 'mid', 'bottom']).default('mid'),
    name: shortText.optional()
  }).refine(v => v.width > 0 || v.height > 0, 'At least one of width or height must be positive.'),
  z.strictObject({ type: z.literal('table'), rows: z.array(z.array(text).min(1).max(30)).min(1).max(100), ...itemGeometry, fontSize: z.number().min(1).max(200).default(14), name: shortText.optional() })
]);
export const deckSchema = z.strictObject({
  theme: z.enum(Object.keys(designThemes)).optional(),
  title: text.default(''),
  width: z.number().min(72).max(4000).default(960),
  height: z.number().min(72).max(4000).default(540),
  slides: z.array(z.strictObject({ background: z.union([color, backgroundFillSchema]).optional(), items: z.array(slideItem).max(300), notes: text.optional() })).min(1).max(100)
});
const mode = z.enum(['file', 'native-copy']).default('file');
export const contracts = {
  ppt_relayout: z.strictObject({ documentId, expectedRevision: z.number().int().nonnegative(), operationId,
    slide: z.number().int().min(1).max(100), targetRefs: z.array(targetRef).min(1).max(12), titleRef: targetRef.optional(),
    layout: z.enum(['two-column','grid','stack']).default('two-column'), margin: z.number().min(0).max(200).default(40),
    gap: z.number().min(0).max(100).default(24), theme: z.enum(Object.keys(designThemes)).optional(),
    fontSize: z.number().min(14).max(40).default(22), dryRun: z.boolean().default(true) }),
  ppt_compose: z.strictObject({
    templatePath: z.string().min(1).max(32760), templateSlide: z.number().int().min(1).max(100).default(1),
    sourcePath: z.string().min(1).max(32760).optional(), contentDeck: deckSchema.optional(),
    sourceSlides: z.array(z.number().int().min(1).max(100)).min(1).max(100).optional(),
    copyDecorations: z.boolean().default(true), beautify: z.boolean().default(true), margin: z.number().min(0).max(200).default(36),
    fontFace: shortText.optional(), textColor: color.optional(), titleSize: z.number().min(8).max(100).default(30), bodySize: z.number().min(8).max(100).default(20),
    allowOffice: z.boolean().default(false), operationId, layoutCheck: z.boolean().default(true)
  }).refine(v => (v.sourcePath !== undefined) !== (v.contentDeck !== undefined), 'Supply either sourcePath or contentDeck.'),
  ppt_start: z.strictObject({ taskId: z.string().uuid() }),
  ppt_open: z.strictObject({ path: z.string().min(1).max(32760), mode, allowOffice: z.boolean().default(false), visible: z.boolean().default(false), operationId }),
  ppt_inspect: z.strictObject({ documentId, text: text.optional(), slide: z.number().int().positive().optional(), kind: z.enum(['text', 'table', 'image', 'shape', 'group', 'notes', 'background', 'chart', 'smartart']).optional(), offset: z.number().int().nonnegative().default(0), limit: z.number().int().min(1).max(300).default(100), detail: z.enum(['full','summary']).default('full') }),
  ppt_apply: z.strictObject({ documentId, expectedRevision: z.number().int().nonnegative(), operationId, operations: z.array(operationSchema).min(1).max(500), dryRun: z.boolean().default(false), layoutCheck: z.boolean().default(true) }),
  ppt_build: z.strictObject({ deck: deckSchema, mode, allowOffice: z.boolean().default(false), visible: z.boolean().default(false), operationId, layoutCheck: z.boolean().default(true) }),
  ppt_render: z.strictObject({ documentId, reviewId: z.string().uuid().optional(), expectedRevision: z.number().int().nonnegative().optional(), slides: z.array(z.number().int().positive()).min(1).max(100), allowOffice: z.literal(true), width: z.number().int().min(320).max(3840).default(1280), detail: z.enum(['full','summary']).default('full'), layoutCheck: z.boolean().default(false) })
    .refine(v => !v.layoutCheck || v.reviewId, 'Render layout checks require a committed reviewId.'),
  ppt_validate: z.strictObject({ documentId, nativeReadback: z.boolean().default(false), allowOffice: z.boolean().default(false),
    layoutCheck: z.boolean().default(false), slides: z.array(z.number().int().positive()).min(1).max(100).optional(),
    expectedRevision: z.number().int().nonnegative().optional(),
    allowedOverlapPairs: z.array(z.tuple([z.string().min(1).max(6000),z.string().min(1).max(6000)])).max(300).optional(),
    checks: z.enum(['full','layout']).default('full'), detail: z.enum(['full','summary']).default('full')
  }).refine(v => (!v.slides && !v.allowedOverlapPairs?.length && v.checks !== 'layout') || v.layoutCheck, 'slides/allowedOverlapPairs/layout-only checks require layoutCheck: true.')
    .refine(v => !v.allowedOverlapPairs?.length || v.expectedRevision !== undefined, 'Design overlap exceptions require expectedRevision.'),
  ppt_commit: z.strictObject({ documentId, expectedRevision: z.number().int().nonnegative(), outputPath: z.string().min(1).max(32760), operationId }),
  ppt_close: z.strictObject({ documentId, preserveCheckpoint: z.boolean().default(false) }),
  ppt_finish: z.strictObject({ reviewIds: z.array(z.string().uuid()).max(100).default([]), preserveCheckpoints: z.boolean().default(false), requireAccepted: z.boolean().default(false) }),
  ppt_status: z.strictObject({ taskId: z.string().uuid().optional(), operationId: operationId.optional(), reviewId: z.string().uuid().optional() }),
  ppt_diagnose: z.strictObject({})
};

export const descriptions = {
  ppt_relayout: 'Plan or apply a bounded layout for selected fresh references on a problem slide. Defaults to dryRun. Preserves text and data. Applied changes return a page audit; preview unexpected overlap now, otherwise continue batch drafting and review the final output later.',
  ppt_compose: 'Create a file-mode composition in a multi-page batch: preserve template pages and append source PPTX or contentDeck pages. Use templateSlide as layout reference. Return page layout reports by default. Publish separately with ppt_commit, then preview the committed output.',
  ppt_start: 'Start a new task after finishing the current task. Supply a new UUID and reuse it when retrying. The initial server task is created automatically.',
  ppt_open: 'Open a protected file session or explicitly allowed PowerPoint task copy. Never attaches to the active user presentation.',
  ppt_inspect: 'Query objects and revision-bound target references. Use detail:summary for batch editing; full for rich text/chart details. Geometry is in points. Optional slide restricts file-mode object parsing.',
  ppt_apply: 'Execute one bounded batch across one or more pages using fresh references and an idempotency key. Automatically return page layout reports for changed slides; fix unexpected overlap before the next batch. No automatic screenshots.',
  ppt_build: 'Generate multiple editable pages from a declarative deck in one batch. Return compact page layout reports by default. Continue normal pages; visually inspect and repair unexpected overlap on affected pages.',
  ppt_render: 'Preview selected pages with Office consent. reviewId renders exact committed bytes and accumulates page evidence; layoutCheck:true reuses that full native readback for final whole-deck layout checking and returns reviewBundlePath for code inspection. In file mode omit reviewId and supply expectedRevision for draft preview. Use detail:summary for small responses. Images never approve themselves.',
  ppt_validate: 'layoutCheck:true returns reports by page. checks:layout skips a repeated structural pass; detail:summary reduces output. Omit slides for final whole-deck checking. Known design pairs require expectedRevision; omitted pairs reuse declarations only while page dependencies are unchanged; [] clears selected declarations. Review unexpected overlap immediately before the next batch. Native readback can measure ordinary text bounds, not all diagram internals.',
  ppt_commit: 'Publish to a new output path and create a review bundle. Existing files are never overwritten.',
  ppt_close: 'Close a task document after publication or an explicit decision to preserve its checkpoint.',
  ppt_finish: 'Finish and clean resources; requireAccepted:true requires current whole-deck checks without stale pages or known geometry errors plus caller-reviewed preview evidence for every output page. Unresolved geometry is recorded as visual fallback. Without strict acceptance cleanup can finish as incomplete. IDs alone never certify unrendered visuals. Inspect the shutdown report.',
  ppt_status: 'Read task documents, page issue summaries, stale checks and next actions without Office. Optional operationId/reviewId retrieves durable receipts or preview records. This is recorded state, not fresh verification of external files.',
  ppt_diagnose: 'Read runtime and capability information without starting Office.'
};
