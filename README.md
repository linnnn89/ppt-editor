# ppt-editor

Create and edit PowerPoint presentations through an MCP server or command-line interface. The project supports batch changes to `.pptx` files, checks slide layouts, and uses Microsoft PowerPoint to render the saved result for review. Original files are preserved; edited presentations are saved to a new path.

**Version:** `0.4.0-rc.4` — release candidate for Windows x64 and Node.js 24.

## Features

| Area | Supported operations |
|---|---|
| Text | Replace text; change fonts, size, color, bold, italic and underline; edit speaker notes |
| Objects | Change position, size and rotation; edit table cells; generate slides with editable shapes, lines and arrows |
| Backgrounds | Solid colors, two-color gradients, images, textures, patterns and inherited backgrounds |
| Templates | Keep template slides and append an existing deck or newly generated slides |
| Charts | Create bar, line, area, pie, doughnut, radar and combination charts with embedded static data |
| SmartArt | Convert text and edit node text through PowerPoint, subject to available layouts |
| Review | Report possible overlaps and objects outside slide bounds; render drafts or saved outputs; track review status |

There are two editing modes:

- **`file`** edits the OOXML package directly. Use it for supported batch operations. Editing and geometry checks do not start PowerPoint.
- **`native-copy`** edits a separate copy through PowerPoint COM. Use it when an operation needs native Office support. It does not attach to the user's active presentation.

Rendering requires PowerPoint in either mode. File-mode editing can therefore use native rendering for final review without moving all edits into COM.

## Installation and MCP setup

Requirements:

- Windows x64 and Node.js 24.
- Python and MSVC C++ build tools for the `winax` native module.
- Microsoft PowerPoint for native editing, rendering and Office integration tests.

```powershell
git clone https://github.com/linnnn89/ppt-editor.git
cd ppt-editor
npm ci
```

The install script builds `winax` and runs a native-module smoke test. It uses Node 24.18.1 headers to avoid a known addon issue with the 24.19.0 headers; it does not replace the installed Node runtime. See [build-native.js](scripts/build-native.js) for the build configuration.

### Connect Codex

Add the following to the project's `.codex/config.toml`. Replace the Node executable and repository paths with the paths on your machine.

```toml
[mcp_servers.ppt_editor]
command = "C:/Program Files/nodejs/node.exe"
args = ["C:/path/to/ppt-editor/src/mcp.js", "--base-dir", "C:/path/to/ppt-editor/work/codex-tasks"]
cwd = "C:/path/to/ppt-editor"
startup_timeout_sec = 15
tool_timeout_sec = 180
enabled = true
```

Run the server directly with Node so package-manager output does not interfere with STDIO. Starting the server does not start PowerPoint.

The [project Skill](skills/ppt-editor/SKILL.md) provides editing instructions for the agent. To make it discoverable in this project, run the following from the repository root. If the destination already exists, inspect it before making changes.

```powershell
New-Item -ItemType Directory -Path .agents/skills -Force | Out-Null
New-Item -ItemType Junction -Path .agents/skills/ppt-editor -Target (Resolve-Path skills/ppt-editor).Path
```

Refresh the MCP connection, then call `ppt_diagnose` to check the running version and capabilities. Updating files on disk does not reload an existing server process. Local MCP configuration, the Skill link and task data are excluded from Git.

## Recommended workflow

**Edit in batches, check reports by slide, fix affected slides, then review the saved presentation.**

1. **Inspect and edit.** Start in file mode. A short deck may fit in one batch; split larger tasks by layout or operation count. Use `ppt_inspect` with `detail:"summary"` and obtain the current object references before calling `ppt_apply`. Inspection returns at most 300 objects per request; apply accepts at most 500 operations. Handle pagination where needed.
2. **Read the result before continuing.** Build, compose and apply return layout reports by default. Wait for the actual response and check whether the edit and its follow-up check succeeded. Continue drafting when the checked slides are clear. An audit failure does not undo an edit that already completed.
3. **Review problems when they occur.** Render and inspect a slide immediately if the report or content suggests unintended overlap, objects outside the page or clipped text. Fix confirmed problems and check the affected slide again before continuing. Do not take a screenshot after every ordinary edit. Only declare an overlap intentional when the design supports that decision.
4. **Keep repairs focused.** Query the affected slide with fresh references. Several known, independent fixes may share a batch. Reuse current check results instead of repeating the same validation or loading the entire deck again. Keep unresolved and stale results visible in the worklist.
5. **Review the final output.** Commit to a new path, then call `ppt_render` with the returned `reviewId`, `layoutCheck:true` and `detail:"summary"`. This combines the final whole-deck layout check with the readback used for rendering. Inspect every final page, using contact sheets and detailed views as needed. After further edits, commit again and review the new output. Finish with `requireAccepted:true` only when the checks and visual review are complete.

`screenshotRequiredNow` currently identifies overlap reports; a false value does not rule out other visual problems. Code checks cannot determine design intent or fully assess readability. A generated image is not evidence that someone has inspected it.

For a targeted code check, use `ppt_validate` with `layoutCheck:true`, `checks:"layout"`, `slides:[...]` and `detail:"summary"`. For a file-mode draft preview, use `ppt_render` with the current `expectedRevision` and omit `reviewId`. Draft previews do not count toward final acceptance. Office use must be authorized at open/build/compose and render.

Detailed rules for overlap declarations, revision handling and final review are in the [Skill](skills/ppt-editor/SKILL.md).

## Tools and examples

The MCP server exposes 14 tools. The CLI provides the same commands without the `ppt_` prefix.

| Purpose | Tools |
|---|---|
| Manage tasks | `ppt_start`, `ppt_status`, `ppt_finish`, `ppt_diagnose` |
| Open or create a presentation | `ppt_open`, `ppt_build`, `ppt_compose` |
| Inspect and edit | `ppt_inspect`, `ppt_apply`, `ppt_relayout` |
| Check and render | `ppt_validate`, `ppt_render` |
| Save and close | `ppt_commit`, `ppt_close` |

The MCP server creates the first task automatically. After finishing it, call `ppt_start` with a fresh UUID to begin another. For CLI usage, start a task explicitly. Completed task references cannot be used for further edits.

### Inspect a presentation from PowerShell

```powershell
$task = node src/cli.js start | ConvertFrom-Json
$opened = node src/cli.js open --task $task.taskId --path ./sample.pptx --mode file --op open_1 | ConvertFrom-Json
node src/cli.js inspect --task $task.taskId --doc $opened.documentId --slide 1 --detail summary
node src/cli.js close --task $task.taskId --doc $opened.documentId
node src/cli.js finish --task $task.taskId --reviews '[]'
```

This example only inspects the file. It does not start Office, edit the source or approve an output. For command names and basic usage, run `node src/cli.js --help`. Complex arguments can be supplied as JSON through `--stdin`.

### Apply changes

Send the following to `ppt_apply`, replacing the placeholders with values from the current inspection. Add operations to the array to edit multiple objects or slides in one request.

```json
{
  "documentId": "<document UUID>",
  "expectedRevision": 0,
  "operationId": "format_1",
  "operations": [
    {
      "type": "set_style",
      "targetRef": "<current targetRef>",
      "style": { "fontSize": 24, "color": "17365D" }
    }
  ]
}
```

Use the returned revision for subsequent operations. Geometry is expressed in points and rotation in degrees. A top-level placeholder with inherited geometry requires all five transform values: `left`, `top`, `width`, `height` and `rotation`.

### Append a deck to a template

Send the following to `ppt_compose`:

```json
{
  "templatePath": "inputs/template.pptx",
  "sourcePath": "inputs/content.pptx",
  "templateSlide": 1,
  "beautify": true,
  "allowOffice": true,
  "operationId": "compose_1"
}
```

All template slides are retained and source slides are appended. `templateSlide` selects the background and layout reference. To append newly generated slides, replace `sourcePath` with `contentDeck`; supply exactly one of them. Review the composition warnings: `beautify` applies typography and margins, but does not guarantee that all content fits.

### Editing reference

- Text and table operations, geometry, themes, charts and tool parameters: [contracts.js](src/contracts.js).
- Background fill fields: `backgroundFillSchema` in [contracts.js](src/contracts.js). Images and textures accept local PNG/JPEG files up to 20 MiB; supported pattern names are listed in [background.js](src/background.js).
- Layout planning: `ppt_relayout` supports `two-column`, `grid` and `stack` for selected objects on one slide. It defaults to `dryRun:true`; inspect the plan before applying it.
- SmartArt: layouts must be available in Office or a template opened in the same task. The regression fixture verifies `hierarchy`; other layouts depend on the installation.

## File safety and recovery

Edits use task copies and revision-bound object references. `ppt_commit` requires a new output path and refuses to overwrite the source or an existing file. Native tasks only close documents they own and preserve pre-existing PowerPoint sessions.

| Result | Action |
|---|---|
| `REVISION_MISMATCH` or a stale reference | Inspect the current document and rebuild the edit using current references. |
| Missing or uncertain response | Query `ppt_status` with the operationId. Retry the same request with the same ID only after checking its receipt. |
| `NATIVE_BUSY` | Wait for the task holding the Office lease to finish. Do not close another task's Office session. |
| `CLEANUP_UNCONFIRMED` | Inspect status and retained checkpoints. Do not delete recovery records or assume cleanup succeeded. |
| `acceptance:incomplete` | Treat the output as not yet accepted. Resource cleanup and final review are separate results. |

`ppt_status` reads recorded state; it does not revalidate external files. A new MCP connection normally starts a new task. To recover an unfinished task, restart the server with the original `--task` UUID and `--base-dir`, then inspect its state. The [Skill](skills/ppt-editor/SKILL.md) describes recovery and reference handling in more detail.

## Limitations

- Layout checks cover object bounds and possible overlaps. Native readback can measure some text bounds, but chart and SmartArt internals, grouped child coordinates, strokes and effects are not fully covered.
- Text changes do not automatically resize or reflow all content. Review long replacements in the rendered slides.
- Shapes can be generated on new slides; arbitrary shape insertion into an existing slide is not supported. Lines use fixed coordinates and do not follow moved shapes.
- Existing chart series cannot be edited. Native workbook access is limited to static chart data; formulas, external links, macros and unrelated embedded files are rejected.
- Macro-enabled, encrypted, DRM-protected and digitally signed documents are unsupported. OCR, animation editing, 3D objects and audio/video editing are also outside the current scope.
- Recovery tests cover orderly STDIO closure and retained checkpoints, not every forced termination or Office failure. Native startup may briefly change window focus.

## Development and testing

```powershell
npm run check          # JavaScript syntax
npm run check:privacy  # Candidate files and staged content
npm test               # Unit and integration tests, including real Office operations
```

Run the full suite on a desktop without an existing PowerPoint session; some ownership tests require that condition. The suite uses synthetic presentations. Privacy scanning checks known patterns and file types, and does not replace review of the files being committed.

Performance comparisons are described in [Benchmark methodology](docs/benchmarks/methodology.md). Measure the same work and acceptance criteria across methods; distinguish tool execution time from model reasoning, and response bytes from token usage or cost.

Production code is in `src/`, tests in `tests/`, and validation scripts and fixtures in `scripts/`. The [changelog](CHANGELOG.md) records version changes. Private presentations, generated outputs, task records and local archives belong in Git-ignored directories such as `private/`, `outputs/`, `work/` and `TEMP/`.
