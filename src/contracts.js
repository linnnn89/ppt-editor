import { z } from 'zod';

const text = z.string().max(20000).refine(s => s.isWellFormed() && !/[\x00-\x08\x0B\x0C\x0E-\x1F]/.test(s), 'Text contains invalid XML characters.');
const shortText = z.string().min(1).max(300);
const operationId = z.string().min(1).max(100).regex(/^[a-zA-Z0-9_-]+$/);
const documentId = z.string().uuid();
const color = z.string().regex(/^[0-9a-fA-F]{6}$/);
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
  italic: z.boolean().optional()
}).refine(v => Object.keys(v).length > 0, 'At least one style property is required.');
const targetRef = z.string().min(20).max(6000);
export const operationSchema = z.discriminatedUnion('type', [
  z.strictObject({ type: z.literal('replace_text'), targetRef, search: text.refine(s => s.length > 0 && !/[\r\n]/.test(s), 'Search must be nonempty and within one paragraph.'), replacement: text.refine(s => !/[\r\n]/.test(s), 'Replacement must stay within one paragraph.'), expectedMatches: z.number().int().min(1).max(1000), crossRunPolicy: z.enum(['reject', 'first_run']).default('reject') }),
  z.strictObject({ type: z.literal('set_style'), targetRef, style }),
  z.strictObject({ type: z.literal('set_geometry'), targetRef, geometry }),
  z.strictObject({ type: z.literal('set_table_cell'), targetRef, row: z.number().int().positive().max(1000), column: z.number().int().positive().max(1000), text, formatPolicy: z.literal('first_run') })
]);

const itemGeometry = { left: z.number().finite(), top: z.number().finite(), width: z.number().positive().max(4000), height: z.number().positive().max(4000) };
const slideItem = z.discriminatedUnion('type', [
  z.strictObject({ type: z.literal('text'), text, ...itemGeometry, fontSize: z.number().min(1).max(400).default(24), fontFace: shortText.default('Arial'), color: color.default('000000'), bold: z.boolean().default(false), name: shortText.optional() }),
  z.strictObject({ type: z.literal('shape'), shape: z.enum(['rect', 'ellipse', 'line']), ...itemGeometry, fill: color.optional(), line: color.default('000000'), name: shortText.optional() }),
  z.strictObject({ type: z.literal('table'), rows: z.array(z.array(text).min(1).max(30)).min(1).max(100), ...itemGeometry, fontSize: z.number().min(1).max(200).default(14), name: shortText.optional() })
]);
export const deckSchema = z.strictObject({
  title: text.default(''),
  width: z.number().min(72).max(4000).default(960),
  height: z.number().min(72).max(4000).default(540),
  slides: z.array(z.strictObject({ background: color.default('FFFFFF'), items: z.array(slideItem).max(300), notes: text.optional() })).min(1).max(100)
});
const mode = z.enum(['file', 'native-copy']).default('file');
export const contracts = {
  ppt_open: z.strictObject({ path: z.string().min(1).max(32760), mode, allowOffice: z.boolean().default(false), visible: z.boolean().default(false), operationId }),
  ppt_inspect: z.strictObject({ documentId, text: text.optional(), slide: z.number().int().positive().optional(), kind: z.enum(['text', 'table', 'image', 'shape', 'group', 'notes']).optional(), offset: z.number().int().nonnegative().default(0), limit: z.number().int().min(1).max(300).default(100) }),
  ppt_apply: z.strictObject({ documentId, expectedRevision: z.number().int().nonnegative(), operationId, operations: z.array(operationSchema).min(1).max(500), dryRun: z.boolean().default(false) }),
  ppt_build: z.strictObject({ deck: deckSchema, mode, allowOffice: z.boolean().default(false), visible: z.boolean().default(false), operationId }),
  ppt_render: z.strictObject({ documentId, slides: z.array(z.number().int().positive()).min(1).max(100), allowOffice: z.literal(true), width: z.number().int().min(320).max(3840).default(1280) }),
  ppt_validate: z.strictObject({ documentId, nativeReadback: z.boolean().default(false), allowOffice: z.boolean().default(false) }),
  ppt_commit: z.strictObject({ documentId, expectedRevision: z.number().int().nonnegative(), outputPath: z.string().min(1).max(32760), operationId }),
  ppt_close: z.strictObject({ documentId, preserveCheckpoint: z.boolean().default(false) }),
  ppt_finish: z.strictObject({ reviewIds: z.array(z.string().uuid()).max(100).default([]), preserveCheckpoints: z.boolean().default(false) }),
  ppt_status: z.strictObject({ operationId: operationId.optional() }),
  ppt_diagnose: z.strictObject({})
};

export const descriptions = {
  ppt_open: 'Open a protected file session or explicitly allowed PowerPoint task copy. Never attaches to the active user presentation.',
  ppt_inspect: 'Query slides and objects, with exact revision-bound target references. Geometry is in points.',
  ppt_apply: 'Preflight and execute one bounded batch in the current backend. Requires fresh references and an idempotency key.',
  ppt_build: 'Generate an editable presentation from a bounded declarative JavaScript/JSON deck specification.',
  ppt_render: 'Explicitly use PowerPoint to export selected slide images for the current revision. Images are not automatically visually approved.',
  ppt_validate: 'Validate package structure and changed parts; optionally perform explicitly allowed native readback.',
  ppt_commit: 'Publish to a new output path and create a review bundle. Existing files are never overwritten.',
  ppt_close: 'Close a task document after publication or an explicit decision to preserve its checkpoint.',
  ppt_finish: 'Finish after reviewing the latest review bundles, clean up owned resources and exit. The caller must verify actual process exit.',
  ppt_status: 'Read task or operation state without starting PowerPoint.',
  ppt_diagnose: 'Read runtime and capability information without starting Office.'
};
