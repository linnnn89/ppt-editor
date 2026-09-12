# Editing workflow

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

## Benchmarking

When comparing workflows, freeze the input, target changes and acceptance criteria
before execution. Record actual request/response ordering and measured time; keep
operator waits separate from execution time and retain gross elapsed time. Label
scripted, predefined corrections as such. They do not measure LLM discovery of
unknown errors or independent reasoning time. Response bytes are not tokens or
money, and a small single-deck sample is not a universal speed ranking. Manual
benchmark gates, mandatory per-page changes and initial-deck rendering are not
requirements for ordinary editing unless the user requests them.
