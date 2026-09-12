# Connection and recovery

## Connection and task lifecycle

Version: `0.4.0-rc.4` (local candidate; configuration and discovery do not prove a completed workflow).
Codex discovers this skill through the local `.agents/skills/ppt-editor` link.
Configure the project MCP server as `ppt_editor` following the [repository README](../../../README.md#installation-and-mcp-setup).
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

## Document and process identity

Read `ppt_status.bindings` to identify the exact source and task copy. Records
include source hash and observed worker/Office identity (PID, creation time and
executable path), where available; historical bindings are not live process checks.
Never reuse references across tasks or documents. CROSS_TASK_REFERENCE,
TASK_BINDING_MISMATCH, NATIVE_PROCESS_MISMATCH or DOCUMENT_IDENTITY_MISMATCH require
stopping the edit and inspecting the bound state, not switching the active window.
Two file tasks use isolated copies. Only one native task per user holds the Office
lease; NATIVE_BUSY means preserve both tasks and wait for the owning task to finish.

## Uncertain results and cleanup

If an operation response is missing, inspect the persisted operation receipt and current document revision before retrying. A retry of the same request uses the same operationId. A failed layout audit does not undo a completed edit. Do not regenerate or repeat the edit with a new ID just to obtain a successful check.

On `CLEANUP_UNCONFIRMED`, inspect `ppt_status`, its cleanup report and retained checkpoints. An unknown shutdown is persisted and blocks further close, finish and native startup, including after reconnect. Do not erase the record or blindly retry. Preserve recovery data and report what remains unconfirmed.

Inspect `acceptance`, `shutdownReport` and `checkpointCleanupError` before reporting completion. Confirmed resource cleanup alone does not mean that the output passed review.
