# Tool reference

Read the section for the operation you need. The [editing workflow](workflow.md) defines batch checks and final review.

## Open Document Session (`ppt_open`)

- **File Mode (`mode: "file"`) [Default]**: Edits the OOXML package directly without starting PowerPoint.
- Prefer file batches for supported edits and defer Office to final readback/render. For top-level text placeholders with inherited geometry, supply all five absolute transform values (`left`, `top`, `width`, `height`, `rotation`); partial inherited changes are rejected, and group children remain unsupported. Use native-copy when the edit requires an unsupported file capability or unresolved effective properties.
- **Native-Copy Mode (`mode: "native-copy"`)**: Uses independent PowerPoint COM worker process on an isolated task copy. Requires explicit `allowOffice: true`.
- Use a unique `operationId` for each new request (e.g. `op_open_01`).

```json
{
  "path": "path/to/presentation.pptx",
  "mode": "file",
  "operationId": "op_open_01"
}
```

## Query Target Elements (`ppt_inspect`)

- Retrieves slides, shapes, text bodies, tables, and slide backgrounds.
- Returns **`revision`** (current document version) and **`targetRef`** (encoded context carrying task, document, mode, revision, generation, and shape fingerprint; not a cryptographic signature).
- **Rule**: You must extract the `targetRef` and note the current `revision` for the subsequent `ppt_apply` step.

## Apply Batch Operations (`ppt_apply`)

- Pass the current version as `expectedRevision`. If the document was modified concurrently, the server will reject the request with `REVISION_MISMATCH`.
- Supported operation types:
  - `replace_text`: Exact text search and replacement. By default, `crossRunPolicy: "reject"`. Use `"first_run"` if replacing text spanning multiple styled runs.
  - `set_style`: Set `fontSize`, `fontFace` (Latin and CJK), `color` (6-char hex like `"17365D"`), `bold`, `italic`, `underline`; booleans support true and false.
  - `set_geometry`: Set `left`, `top`, `width`, `height`, `rotation` in points (pt).
  - `set_table_cell`: Update text in a specific table cell (`row`, `column`, 1-indexed).
  - `set_slide_background`: Use legacy `color` OR `fill` (never both). See [Background fills](#background-fills) for fields and limits.
  - `convert_to_smartart` (native-copy): Convert an inspected top-level text box with `layout: process | cycle | hierarchy`. The layout must exist in Office or an explicitly opened same-task template. Never substitute another layout after SMARTART_LAYOUT_UNAVAILABLE. The regression fixture verifies hierarchy only.
  - `set_smartart_text` (native-copy): Set a 1-based `nodeIndex` from the fresh SmartArt snapshot. Long text may overflow; no automatic font shrinking is promised.

### Background fills

`fill.type` selects the required fields: `solid` uses `color`; `gradient` uses
`startColor`, `endColor` and optional `angle` (0–360 degrees, default 0);
`pattern` uses `pattern`, `foreground`, `background`; `image` and `texture` use
`path`; `inherit` has no other fields. Colors are six-digit hex. All types except
`inherit` accept `transparency` from 0 to 100 (default 0).
Images and textures require a local PNG/JPEG file of at most 20 MiB. Use pattern
names from `ppt_diagnose.capabilities.backgroundPatterns`; do not invent names.

## Commit and Finish (`ppt_commit` & `ppt_finish`)

- **`ppt_commit`**: Atomically publishes the edited deck to a **new output path**. **Overwriting the source file is strictly prohibited by design.** Returns a unique `reviewId`.
- **`ppt_validate`**: Checks supported package structure and relationships. This is not a full OOXML schema or visual validation.
- File-mode native readback requires `allowOffice: true` at both open and validate, plus `nativeReadback: true`. The returned `nativeValidation` binds the read-only candidate hash to the current revision; compare with the actual committed output before reusing that evidence.
- **`ppt_finish`**: For delivery use `requireAccepted:true` with reviewed IDs. Inspect `acceptance`, `shutdownReport` and `checkpointCleanupError`; cleanup completion alone is not successful acceptance.
- **`ppt_render`**: For final review use `documentId` + committed `reviewId`; `layoutCheck:true` also records a final whole-deck layout check from that readback. Office permission is required at open/build/compose and render. Inspect the PNGs before passing reviewIds to finish. Without reviewId, file drafts require expectedRevision; native-copy can preview its current session. Draft images never replace final output evidence.

## Template composition

Prefer `ppt_compose` for a template plus an old deck, or for appending new pages.
Keep all template pages, append source pages, and use templateSlide (default 1)
as the background/layout reference. Use sourcePath OR contentDeck, exactly one.
Optional sourceSlides selects old pages. copyDecorations defaults true and copies
non-text, non-placeholder page graphics behind content. Textual page branding is
not duplicated. beautify defaults true, applies uniform text/table typography,
and fits objects into margins; it does not guarantee visual reflow. Check warnings.
Do not erase source content to resolve overflow or large-object background cover.
Inspect the composed file session and perform the page checks in the [editing workflow](workflow.md). At the
final pass, commit to a new path, render the committed reviewId with explicit Office
consent, inspect the PNGs, then finish. For later
pages, use the previous output as templatePath with a fresh contentDeck/task.
Animations and unsupported cross-page links are rejected. A separate advanced
fixture verifies chart workbook and SmartArt data preservation through composition;
this does not cover every imported chart/diagram or linked data source.
Use `ppt_diagnose` to verify that `version` matches `diskVersion`,
`restartRequired` is false, and `templateComposition` is available. Current
diagnostics also expose `queuedRequestCancellation` and `pageTextMeasurementCoverage`.
`checkpointPartReuse:"sha256-verified"`, `progressNotifications:true` and
`requestStageTimings:true` identify verified part reuse and request progress support.
Refresh the MCP client if its version, capabilities or 14-tool list is outdated.

## Charts and relayout

`ppt_build` accepts editable bar/line/area/pie/doughnut/radar/combo charts with
explicit categories and values. Combo series declare their own type and optional
secondaryAxis. Only static chart-owned XLSX data may be opened natively; formulas,
external data, macros and arbitrary embeddings remain rejected.
Build or compose.contentDeck may use executive/editorial/academic themes; explicit
styles override theme defaults. These are original palettes, not downloaded templates.

Use `ppt_relayout` with fresh same-slide `targetRefs`, optional `titleRef`, and
two-column/grid/stack. dryRun defaults true. Inspect placements and warnings, then
apply with a fresh operationId and dryRun false. Reinspect and code-check the page;
review suspected overlap or clipping immediately, otherwise continue
drafting and review the committed result at the end. Text and data are retained.
Do not claim automatic visual approval.

## Unsupported operations

Macros, DRM and encrypted presentations are rejected. Animation editing, 3D models, internal SmartArt reshaping and video editing are unsupported. OCR is not provided; text inside an image cannot be edited as slide text. Long text replacements do not automatically reflow; check them in the rendered result.
