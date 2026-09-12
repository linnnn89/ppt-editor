---
name: ppt-editor
description: Create, inspect, edit and review PowerPoint presentations with the project's ppt-editor MCP tools or CLI. Use for file-based batch editing, native Office operations and slide layout checks.
---

# PowerPoint editing

Use the project's ppt-editor tools. Default to file-mode batch editing, check results by slide, refine affected slides and review the saved presentation.

## Select a manual

Read only the manual or section needed for the current task. Do not load every reference at startup.

| Task | Read |
|---|---|
| Create, edit or review a deck | [Editing workflow](references/workflow.md) |
| Inspect objects without editing | [Inspect](references/tools.md#query-target-elements-ppt_inspect); the editing workflow is unnecessary for a read-only query |
| Choose an operation or its parameters | The relevant section of [Tool reference](references/tools.md) |
| Use templates, charts, SmartArt or relayout | [Template composition](references/tools.md#template-composition) or [Charts and relayout](references/tools.md#charts-and-relayout), alongside the editing workflow |
| Connect, resume a task, retry an uncertain operation or handle cleanup failure | [Connection and recovery](references/recovery.md) |
| Compare workflow performance | [Benchmarking](references/workflow.md#benchmarking) |

## Rules for every task

- Use inspected object references and the current revision. Wait for actual tool results before dependent actions; do not guess targets or assume a pending edit succeeded.
- Preserve the source and save to a new path. Use Office only when authorized; existing authorization remains valid.
- For editing, keep multi-page batches. Inspect suspected unintended overlap or clipping immediately; ordinary page edits do not need individual screenshots.
- Final delivery requires current whole-deck checks, inspection of every final page and strict acceptance. Draft previews and successful cleanup alone do not establish acceptance.
- If tools are unavailable or a requested capability is unsupported, report that limit. Do not claim unperformed editing, rendering or verification.
