# Efficient code orchestration

Use this path when the host can call MCP tools from JavaScript and retain values
between calls. The examples use Codex's `tools`, `store`, `load`, `text` and `exit`;
use the host's documented equivalents elsewhere. With direct tool calls only,
use `detail:"summary"`, scoped queries and the returned `targetRef` normally.

Keep the complete tool response in code state. Show one business result to the
model: prefer `structuredContent`, falling back to the JSON text in `content`.
Inspect `isError` first and preserve the error rather than presenting it as success.
Do not print both representations or append routine timing diagnostics.

## Inspect once, retain references

Store `ppt.inspectArgs` with the actual `documentId` and the required `slide`,
`text`, `kind`, `offset` or `limit` filters. Omit filters that would exclude needed
objects; a batch may span several pages. This example defaults to summaries and
honors an explicit `detail`. Use `detail:"full"` when relevant text is truncated
or rich-text/chart details are needed. Read the next result page only when
required objects are still missing.

```javascript
const request = { detail: "summary", ...load("ppt.inspectArgs") };
store("ppt.inspection", null);
const response = await tools.mcp__ppt_editor__ppt_inspect(request);
store("ppt.inspectResponse", response);
if (response.isError) { text(response); exit(); }
const inspection = response.structuredContent
  ?? JSON.parse(response.content.find(item => item.type === "text").text);
store("ppt.inspection", inspection);
const { objects, ...context } = inspection;
text({ ...context, objects: objects.map(({ targetRef, ...object }) => object) });
```

Only the opaque references are hidden from the model's view. Keep object keys,
text, truncation flags, geometry, capabilities, revision and pagination information.
The original references remain available in `ppt.inspection`; never reconstruct
them from shortened strings or decode them to bypass the server's checks.

The example holds one inspection page. If a batch needs several result pages,
retain them together only when document, revision and generation match, with
unique object keys. An ordinary reinspection replaces the earlier cache. Preserve
the same rule when changing query scope; do not silently mix old and new results.

## Assemble one bounded edit batch

After reading the view, store `ppt.editPlan` with its `documentId`, `revision`
and `generation`, plus `{ operationId, edits }`. These bind the selection to the
view used to make the decision; a later inspection must not rebind an old plan.
Each edit contains the selected object's exact `key` and the normal operation
fields, such as `type`, `search`, `replacement` and `expectedMatches`.
Names and display order are not unique identifiers. Use a new operation ID for a
new request, then retain the assembled request for any uncertain-result retry.

```javascript
const inspection = load("ppt.inspection");
const plan = load("ppt.editPlan");
if (!inspection || !plan?.operationId || !plan.edits?.length) {
  throw new Error("A current inspection and an explicit edit plan are required.");
}
if (plan.documentId !== inspection.documentId || plan.revision !== inspection.revision
    || plan.generation !== inspection.generation) {
  throw new Error("The edit plan no longer matches the inspection. Rebuild it from the current view.");
}
const request = {
  documentId: inspection.documentId,
  expectedRevision: inspection.revision,
  operationId: plan.operationId,
  operations: plan.edits.map(({ key, ...operation }) => {
    const matches = inspection.objects.filter(object => object.key === key);
    if (matches.length !== 1) throw new Error("Inspect the exact target before editing.");
    return { ...operation, targetRef: matches[0].targetRef };
  })
};
store("ppt.applyCall", { request, state: "pending" });
store("ppt.inspection", null);
const response = await tools.mcp__ppt_editor__ppt_apply(request);
store("ppt.applyCall", { request, state: "returned", response });
if (response.isError) { text(response); exit(); }
const result = response.structuredContent
  ?? JSON.parse(response.content.find(item => item.type === "text").text);
text(result);
```

The cached inspection is invalidated before dispatch: an interrupted wait may
still have changed the document. Reinspect for the next new edit batch. A
`returned` call means a response arrived; check `isError`, the business outcome
and any follow-up audit before deciding whether the edit succeeded.

If the response is uncertain, keep `ppt.applyCall.request`. Use `ppt_status` with
the original task ID and operation ID to inspect its durable receipt before
retrying those exact arguments. A confirmed revision mismatch requires a fresh
inspection, not a loop replaying stale references. Code state is temporary; after
it is lost, recover through status/receipts and fresh inspection as described in
[Recovery](recovery.md). Retain task/document/operation IDs in the handoff when needed.

## Reuse the returned evidence

Apply/build/compose already return summary page audits. Output that business
result once, retaining failures, pending or stale pages, coverage limitations and
`nextAction`. Reuse a successful current audit instead of adding a validate call.
Read detailed evidence only for a decision it can change. Render the committed
review with `layoutCheck:true, detail:"summary"` to share the final readback.
Complete the checks and visual review in [Workflow](workflow.md); compact output
does not change acceptance requirements. Compare model-visible JSON size and
tool-call counts separately from tokens, model latency or whole-task speed.
