---
name: ppt-editor
description: Structured, deterministic, and safe PowerPoint editing via standard MCP tools. Supports pure OOXML file manipulation and native Office automation with strict state consistency and revision protection.
---

# PowerPoint Editor Skill (`ppt-editor`)

Use this skill when interacting with the PowerPoint editing core via standard MCP tools.

## Core Philosophy: Inspect-First & Revision Integrity

All edits must follow the inspected, revision-bound lifecycle. **Never guess existing object keys or target references.**

```
[ppt_open] -> [ppt_inspect] -> [ppt_apply] -> [ppt_commit] -> [ppt_finish]
```

Version: `0.4.0-rc.4` (local candidate; configuration and discovery do not prove a completed workflow).
Codex discovers this skill through the local `.agents/skills/ppt-editor` link.
Configure the project MCP server as `ppt_editor` following the repository README.
After changing configuration, refresh the MCP connection and verify the available
tools. If the tools are absent, report that limitation instead of claiming direct
MCP execution. Starting the server or listing tools does not start PowerPoint.
The server creates the first task automatically. After `ppt_finish`, call
`ppt_start` with a fresh UUID in `taskId` before starting the next task. Reuse
that UUID when retrying the start request. An active task cannot be replaced.
Use `ppt_status` with an optional historical `taskId` and `operationId` to query
completed receipts. Finished tasks and their references cannot be reused for edits.
If the connection is lost, keep the taskId, documentId and operationId. A new
server connection creates a new task by default; `ppt_status` can query the old
task but cannot switch back to editing it. Explicit recovery requires restarting
the server with `--task <previous-task-uuid>` and the same `--base-dir`, or using
the documented CLI with that task ID. Inspect durable state before retrying.
Graceful EOF cleanup is verified; forced process termination is a separate case.

---

## 1. Standard Editing Lifecycle

### Work in multi-page batches; report and refine by page

Default workflow: **file-mode batch drafting → returned page reports → targeted
page refinement → final whole-deck check and visual review**. Prefer `mode:"file"`
for supported edits. Use native-copy when a required capability or effective
property cannot be handled reliably in file mode; final Office rendering does not
require moving ordinary editing into native-copy.

For a short deck whose changes fit the tool limits, prefer one whole-deck batch.
Split larger or heterogeneous work by layout, dependency or payload size. Do not
hard-code five pages per batch or make one page the default production unit.
`ppt_apply` accepts at most 500 operations; `ppt_inspect` returns at most 300
objects per request. Check pagination and obtain all required references before
applying a batch. Reinspect after a revision change; do not reuse old references
for the next batch.

Use `ppt_build` / `ppt_compose` for multiple pages and `ppt_apply` for bounded
operations across multiple pages. Do not split ordinary work into one tool cycle
per page. These tools return `layoutAudit` by default after a successful batch;
apply checks the changed pages. A failed follow-up audit does not undo a completed
batch: inspect its outcome, then validate; never regenerate with a fresh operationId
merely because the audit failed. `layoutCheck:false` disables this follow-up only.
File drafting does not start Office or generate images.

- Await the actual tool response and inspect its outcome, revision and page
  reports before issuing dependent edits. Never treat a pending request as
  completed or plan a repair from an assumed result. If a response is missing or
  ambiguous, inspect status/receipts before retrying with the same operationId.
- Read the reports by page. Continue drafting when checked pages are clear.
  Refine identified problem pages with fresh references; several already-known,
  independent fixes may share an apply batch. "Refine by page" does not require
  one tool call per page or inventing a change on every page.
- Non-design overlap is an immediate exception: if `screenshotRequiredNow: true`
  or the content otherwise indicates unintended overlap, visually inspect the
  affected page/region now. Fix confirmed unintended overlap, then rerun that
  page's code check before continuing. Do not defer it until the end.
- Also review suspected unintended page overflow or text clipping immediately,
  even when `screenshotRequiredNow` is false: that flag currently tracks overlaps,
  not every visual problem. Unresolved geometry alone is not proof of an error;
  retain it in the worklist for final visual review, escalating sooner if the
  content suggests an actual defect. Never label unresolved or stale pages clear.
- Code cannot infer design intent. For known intentional layering (for example
  text over a panel), pass `allowedOverlapPairs: [[objectKeyA, objectKeyB]]` from
  current inspected objects together with `expectedRevision`. Never exempt an
  unknown overlap just to suppress review. Omitted pairs reuse saved declarations
  only while the page, layout dependencies and generation remain unchanged;
  explicit `allowedOverlapPairs: []` clears declarations on the checked pages.
- Use `ppt_inspect(detail:"summary")` for ordinary batch changes, requesting full
  detail only for rich-text/chart information. Inspect a specific slide when fixing
  it. `ppt_status.documents[].layout` restores page issues and next actions after
  interruptions. Status is persisted evidence, not a fresh check of external files.
- Keep LLM-facing updates compact: group clear page numbers, and report issue
  pages with object identifiers, the measured problem and next action. Use summary
  responses where supported; do not repeatedly load whole-deck snapshots or echo
  duplicate pending-page lists. Keep the underlying evidence, unresolved states
  and failures intact. Current tools may still return duplicate `pendingSlides`;
  this guidance does not imply that server response compression is implemented.
- For an extra quick check use `ppt_validate(layoutCheck:true, checks:"layout",
  slides:[...], detail:"summary")`. File-mode checks parse selected object trees;
  native mode retains whole-document identity guards. Full validation is the default.
  Reuse a successful automatic check of the same changed pages instead of adding
  a redundant validate call; explicitly validate when the check is missing,
  failed, stale or a different scope is needed.
- Geometry `clear` only covers the reported checks. Native readback can measure
  ordinary unrotated text bounds; diagram internals, strokes, effects and grouped
  child coordinates remain incompletely covered. Inspect them in the final review.

Once all pages are drafted, review the accumulated worklist and fix issues by page.
A final whole-deck check is required after edits. Prefer committing and rendering
the final reviewId with `layoutCheck:true, detail:"summary"`: this reuses the full
native readback of committed bytes for the final layout check, while rendering the
requested pages. It also returns `reviewBundlePath` so code can inspect the stored
full snapshot without returning it all to the LLM. Alternatively run `ppt_validate`
with `layoutCheck:true` and omit `slides` for a separate whole-deck pass. Avoid
duplicating those full reads when the same evidence already exists. Unchanged page
evidence survives unrelated edits; affected dependencies invalidate relevant pages.
Use final page previews/contact sheets and detailed views where needed; do not
generate a screenshot after every ordinary page-edit operation.
For an immediate overlap exception in file mode, call `ppt_render` with the current
`expectedRevision`, affected `slides`, `allowOffice:true`, and `detail:"summary"`.
Omit reviewId to preview checkpoint bytes without a temporary publication. Office
permission must already be present on the session. Final output review still uses
the committed reviewId; draft previews never count as final evidence.

Render final pages in convenient batches. Preview evidence accumulates by page for
the same committed output hash. After inspecting every final page, finish with
`reviewIds` and `requireAccepted:true`. This verifies a current final layout check,
no stale/unchecked pages or unresolved known geometry errors, and all-page preview
evidence. Unresolved geometry is recorded as visual fallback when those pages were
reviewed. Cleanup without strict acceptance can finish with `acceptance:incomplete`;
never report that as accepted delivery. Passing IDs without previews records caller
confirmation, not successful visual review.

After a final-review repair, commit again and use the new reviewId for final checks
and previews. Do not reuse stale output evidence. Review every final page using
contact sheets and readable detailed views as needed; code checks or generated
image paths alone are not visual approval.

### Workflow measurement boundaries

When comparing workflows, freeze the input, target changes and acceptance criteria
before execution. Record actual request/response ordering and measured time; keep
operator waits separate from execution time and retain gross elapsed time. Label
scripted, predefined corrections as such. They do not measure LLM discovery of
unknown errors or independent reasoning time. Response bytes are not tokens or
money, and a small single-deck sample is not a universal speed ranking. Manual
benchmark gates, mandatory per-page changes and initial-deck rendering are not
requirements for ordinary editing unless the user requests them.

### Step 1: Open Document Session (`ppt_open`)
- **File Mode (`mode: "file"`) [Default]**: Pure OOXML XML DOM parsing. Safe, fast, zero Office dependency. Does not start PowerPoint.
- Prefer file batches for supported edits and defer Office to final readback/render. For top-level text placeholders with inherited geometry, supply all five absolute transform values (`left`, `top`, `width`, `height`, `rotation`); partial inherited changes are rejected, and group children remain unsupported. Use native-copy when the edit requires an unsupported file capability or unresolved effective properties.
- **Native-Copy Mode (`mode: "native-copy"`)**: Uses independent PowerPoint COM worker process on an isolated task copy. Requires explicit `allowOffice: true`.
- Always generate an alphanumeric `operationId` (e.g. `op_open_01`).

```json
{
  "path": "path/to/presentation.pptx",
  "mode": "file",
  "operationId": "op_open_01"
}
```

### Step 2: Query Target Elements (`ppt_inspect`)
- Retrieves slides, shapes, text bodies, tables, and slide backgrounds.
- Returns **`revision`** (current document version) and **`targetRef`** (encoded context carrying task, document, mode, revision, generation, and shape fingerprint; not a cryptographic signature).
- **Rule**: You must extract the `targetRef` and note the current `revision` for the subsequent `ppt_apply` step.

### Step 3: Apply Batch Operations (`ppt_apply`)
- Pass the current version as `expectedRevision`. If the document was modified concurrently, the server will reject the request with `REVISION_MISMATCH`.
- Supported operation types:
  - `replace_text`: Exact text search and replacement. By default, `crossRunPolicy: "reject"`. Use `"first_run"` if replacing text spanning multiple styled runs.
  - `set_style`: Set `fontSize`, `fontFace` (Latin and CJK), `color` (6-char hex like `"17365D"`), `bold`, `italic`, `underline`; booleans support true and false.
  - `set_geometry`: Set `left`, `top`, `width`, `height`, `rotation` in points (pt).
  - `set_table_cell`: Update text in a specific table cell (`row`, `column`, 1-indexed).
  - `set_slide_background`: Use legacy `color` OR `fill` (never both): solid, two-color linear gradient, image, texture, pattern, inherit. See README for exact fields and limits.
  - `convert_to_smartart` (native-copy): Convert an inspected top-level text box with `layout: process | cycle | hierarchy`. The layout must exist in Office or an explicitly opened same-task template. Never substitute another layout after SMARTART_LAYOUT_UNAVAILABLE. This machine's fixture verifies hierarchy only.
  - `set_smartart_text` (native-copy): Set a 1-based `nodeIndex` from the fresh SmartArt snapshot. Long text may overflow; no automatic font shrinking is promised.

### Step 4: Commit and Finish (`ppt_commit` & `ppt_finish`)
- **`ppt_commit`**: Atomically publishes the edited deck to a **new output path**. **Overwriting the source file is strictly prohibited by design.** Returns a unique `reviewId`.
- **`ppt_validate`**: Checks supported package structure and relationships. This is not a full OOXML schema or visual validation.
- File-mode native readback requires `allowOffice: true` at both open and validate, plus `nativeReadback: true`. The returned `nativeValidation` binds the read-only candidate hash to the current revision; compare with the actual committed output before reusing that evidence.
- **`ppt_finish`**: For delivery use `requireAccepted:true` with reviewed IDs. Inspect `acceptance`, `shutdownReport` and `checkpointCleanupError`; cleanup completion alone is not successful acceptance.
- On `CLEANUP_UNCONFIRMED`, inspect `ppt_status` and retained checkpoints. An unknown shutdown is persisted and blocks further close/finish/native startup, including after reconnect. Do not erase the record or blindly retry; preserve recovery data and report the unresolved cleanup.
- **`ppt_render`**: For final review use `documentId` + committed `reviewId`; `layoutCheck:true` also records a final whole-deck layout check from that readback. Office permission is required at open/build/compose and render. Inspect the PNGs before passing reviewIds to finish. Without reviewId, file drafts require expectedRevision; native-copy can preview its current session. Draft images never replace final output evidence.

---

## 2. Hard Boundaries & Rules

### Template workflow

Prefer `ppt_compose` for a template plus an old deck, or for appending new pages.
Keep all template pages, append source pages, and use templateSlide (default 1)
as the background/layout reference. Use sourcePath OR contentDeck, exactly one.
Optional sourceSlides selects old pages. copyDecorations defaults true and copies
non-text, non-placeholder page graphics behind content. Textual page branding is
not duplicated. beautify defaults true, applies uniform text/table typography,
and fits objects into margins; it does not guarantee visual reflow. Check warnings.
Do not erase source content to resolve overflow or large-object background cover.
Inspect the composed file session and perform the page code checks above. At the
final pass, commit to a new path, render the committed reviewId with explicit Office
consent, inspect the PNGs, then finish. For later
pages, use the previous output as templatePath with a fresh contentDeck/task.
Animations and unsupported cross-page links are rejected. A separate advanced
fixture verifies chart workbook and SmartArt data preservation through composition;
this does not cover every imported chart/diagram or linked data source.
Use ppt_diagnose to verify version 0.4.0-rc.4 and templateComposition capability;
refresh the MCP client if it does not expose the current 14 tools.

### Advanced generation and layout

`ppt_build` accepts editable bar/line/area/pie/doughnut/radar/combo charts with
explicit categories and values. Combo series declare their own type and optional
secondaryAxis. Only static chart-owned XLSX data may be opened natively; formulas,
external data, macros and arbitrary embeddings remain rejected.
Build or compose.contentDeck may use executive/editorial/academic themes; explicit
styles override theme defaults. These are original palettes, not downloaded templates.

Use `ppt_relayout` with fresh same-slide `targetRefs`, optional `titleRef`, and
two-column/grid/stack. dryRun defaults true. Inspect placements and warnings, then
apply with a fresh operationId and dryRun false. Reinspect and code-check the page;
render immediately only for suspicious/non-design overlap, otherwise continue
drafting and review the committed result at the end. Text and data are retained.
Do not claim automatic visual approval.

### Task, document and process binding

Read `ppt_status.bindings` to identify the exact source and task copy. Records
include source hash and observed worker/Office identity (PID, creation time and
executable path), where available; historical bindings are not live process checks.
Never reuse references across tasks or documents. CROSS_TASK_REFERENCE,
TASK_BINDING_MISMATCH, NATIVE_PROCESS_MISMATCH or DOCUMENT_IDENTITY_MISMATCH require
stopping the edit and inspecting the bound state, not switching the active window.
Two file tasks use isolated copies. Only one native task per user holds the Office
lease; NATIVE_BUSY means preserve both tasks and wait for the owning task to finish.

1. **Source Protection**: Never attempt to write directly back to the original source presentation path. Always specify a distinct destination path in `ppt_commit`.
2. **No Speculative Modifications**: If a target element is not found during `ppt_inspect`, stop and report to the user instead of trying alternative uninspected shapes.
3. **Typography & Overflow Warning**: Text modifications do not execute visual auto-reflow. If replacement text is significantly longer than original text, warn the user about potential bounding box overflow.
4. **Unsupported Features**:
   - Macros (VBA), DRM, encrypted presentations are rejected at open.
   - Animations, 3D models, complex SmartArt internal reshaping, and embedded video trimming are out of scope.
   - OCR is not provided; text embedded inside bitmap images cannot be replaced.
