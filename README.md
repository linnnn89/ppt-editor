# ppt-editor

Current source version: `0.4.0-rc.4` (release candidate; not a stable release). Default to file-mode batch drafting, page-level check reports, targeted refinement and final whole-deck visual review. Inspect suspected unintended overlap immediately. See the [recommended LLM workflow](#recommended-llm-workflow) and [project Skill](skills/ppt-editor/SKILL.md). Refresh the desktop MCP connection and verify the runtime version and `batchPageReports` capability; updated source files do not prove that the connected host has reloaded.

PowerPoint 结构化编辑与自动化任务核心。提供面向 AI Agent 与工具链的确定性文档编辑能力，支持 Windows 环境下的 Node 24 运行时。

## 核心能力与架构

系统采用“纯文件解析（OOXML）”与“原生 Office 自动化（COM）”双引擎分层架构，由统一任务层（TaskHost）与 CLI 入口串联：

- **双引擎编辑**：
  - **文件模式（File Mode）**：基于 OOXML XML DOM 与 ZIP 流直接操作，具备原稿物理隔离与原子持久化检查点，无需启动 Office 即可处理；
  - **原生副本模式（Native-Copy Mode）**：调用独立 PowerPoint COM 子进程操作副本，严格限制权限；仅在冷启动独占时执行安全退出，不关闭或影响系统预先存在的外部 Office 会话。
- **已验证的编辑能力**：
  - 文本内容：段落/Run 级文本精确查找与替换（支持单 Run 严格匹配或 `first_run` 策略）；
  - 文本格式：字号、字体（中西文字体族）、十六进制颜色、粗体、斜体、下划线设置与取消；
  - 表格内容：普通表格单元格文本修改（暂不支持单元格拆分与合并）；
  - 幻灯片背景：纯色、双色线性渐变、图片、平铺纹理、11 种图案及恢复母版继承；文件与原生引擎均支持；
  - 模板合并：保留模板现有页，将旧稿或新内容接在后面，沿用指定参考页的背景和版式，可统一字体、颜色、字号及页面留白；
  - 备注信息：幻灯片备注（Notes）文本检查与修改；
  - 几何图形与示意结构图：
    - 多图元绘制：矩形（`rect`）、圆角矩形（`roundRect`）、等腰三角形（`triangle`）、直角三角形（`rtTriangle`）、菱形（`diamond`）、方向箭头块（`rightArrow` / `downArrow` 等）、线条（`line`）；
    - 图形内置文本：支持在图形 `<p:sp>` 内部直接生成 `<p:txBody>` 文本，字号、字体、颜色、加粗及居中/左对齐受控，图形与文本物理一体化；
    - 导向箭头与连线：支持末端/始端箭头标记（`arrow: 'end' | 'begin' | 'both'`）与线宽控制，已在原生 PowerPoint 下完成 1920×1080 流程架构图实测。
- **任务事务与状态一致性**：
  - 状态版本控制：每次操作绑定唯一 `targetRef` 与版本号（`expectedRevision`），拦截过期引用或并发状态冲突；
  - 幂等性与防重放：记录 `operationId` 执行回执，保证重复请求不会导致状态漂移；
  - 跨进程互斥：通过文件锁确保多进程并发调用下的事务安全。
- **通用写包性能**：
  - 正式接入增量流复用（`writePackageOptimized`），未改动部件直接复用源压缩流；
  - 媒体资源（图片等）采用 `STORE` 存储，规避无效的重压缩 CPU 开销。

## 明确边界与非目标（当前不支持项）

为确保生产环境下的确定性与文稿安全，本项目明确设定了以下能力边界：

- **高级图形与排版边界**：
  - 结构图连线目前基于绝对坐标定位，暂未绑定 Office 动态磁吸锚点（Connection Site Binding），在 PowerPoint 中手动拖拽节点方块时连线不会自动跟随延展；
  - 可复用模板版式和背景，但不支持任意母版重建、复杂图表系列绑定；
- **排版不承诺自适应**：文本修改不包含基于视觉模型的自动重排算法。文字长度超长可能导致超出文本框边界，需人工核对；
- **不受支持的格式与对象**：
  - 暂不支持宏（VBA）、数字签名、加密或 DRM 保护的文档；
  - 图表支持按明确数据新建，暂不修改旧图表的数据系列或连接外部实时数据；SmartArt 内部结构变形、3D 模型、复杂动画序列及音视频剪辑不支持；
  - 不包含 OCR 能力，无法识别或修改图片内嵌文字；
- **平台与运行时约束**：原生模式强依赖 Windows x64 与已安装的 PowerPoint；启动 COM 时可能短暂产生前台焦点变化。

## 环境要求与验证

- 操作系统：Windows x64
- 运行时：Node 24（推荐 Node 24.19.0）
- 本地依赖：Microsoft PowerPoint（仅原生模式需要）、本机 Python 与 MSVC C++ 构建环境（用于编译 `winax` 原生绑定）。

### 常用命令

```powershell
# 依赖安装（项目内编译原生绑定，不修改全局系统环境）
npm ci

# 源码与隐私扫描
npm run check          # 检查全量 JavaScript 语法
npm run check:privacy  # 扫描敏感凭证与本地主机私人路径

# 自动化测试
npm test               # 运行全量单元与集成测试（包含并发锁争用与外部会话隔离测试）
npm run test:native    # 运行原生 COM 读写闭环验证

# 性能实验基准测试（须自备授权的样本路径）
npm run benchmark:fast -- --source "path/to/sample.pptx" --out "outputs/run"

# 启动标准 STDIO MCP 协议服务（供 Codex / AI 客户端直接连接）
node src/mcp.js
```

## 标准 MCP 服务与 AI Agent Skill

### Recommended LLM workflow

**File-mode batch drafting → actual page reports → targeted page refinement → final whole-deck check and visual review.** The [Skill](skills/ppt-editor/SKILL.md) contains the operational rules; this is the default for supported editing tasks, subject to the user's explicit requirements.

1. **Choose the backend and batch size.** Prefer `mode:"file"` for supported edits. Use native-copy for a required unsupported capability or unresolved effective properties, rather than ordinary per-page read/write cycles. A short deck can use one whole-deck batch when its operations fit the tool limits; split larger or heterogeneous tasks by layout, dependencies or payload size. There is no mandatory five-page batch size. `ppt_apply` permits up to 500 operations; `ppt_inspect` permits up to 300 objects per request, so handle pagination before editing.
2. **Read once for the current batch, then apply.** Use `ppt_inspect(detail:"summary")` and current `targetRef` / `expectedRevision`; fetch full detail only when needed for rich text or charts. `ppt_build`, `ppt_compose` and `ppt_apply` return automatic `layoutAudit` reports by default, with apply checking the changed pages. Await the actual response and inspect its outcome before dependent edits. Never assume a pending operation succeeded. A failed follow-up audit does not roll back a completed edit; inspect status/receipts before retrying, preserving the operationId.
3. **Read reports by page and keep drafting when clear.** Keep automatic checks enabled in the normal workflow. Reuse a successful current check rather than immediately repeating validate. Summarize clear page numbers together; report problem pages with object identifiers, measured issues and next actions. Preserve stale, unresolved and failed states. Current responses may duplicate pending-page lists; avoid echoing them or reloading full snapshots unnecessarily. This is caller guidance, not an implemented server-side response compression change.
4. **Inspect visual exceptions immediately.** When `screenshotRequiredNow` is true, or content suggests unintended overlap, render the affected draft pages, inspect them, fix confirmed defects and code-check the changed pages before continuing. Apply the same response to suspected unintended page overflow or text clipping even if the flag is false: the flag currently tracks overlaps. Geometry that is merely unresolved stays on the final visual-review worklist unless there is evidence of a defect. Do not take a screenshot after each ordinary page edit or exempt unknown overlaps just to silence a report.
5. **Refine locally, then finish the whole deck.** Query only the affected page when possible and use fresh references after revisions change. Several known independent fixes may share an apply batch; page-level refinement does not require inventing a change on every page. Commit to a new output path, then use `ppt_render` with the committed reviewId, `layoutCheck:true` and `detail:"summary"` to combine final readback, whole-deck layout checking and previews. Review every final page, using contact sheets and readable detailed views as needed. After repairs, commit again and review the new output; finish with `requireAccepted:true` only after current checks and visual review are complete.

The local 15-page comparison with predefined per-page corrections supports choosing the file backend first. Whole-deck and five-page file batches were close in elapsed time; whole-deck used fewer calls and less response text. Native merged readback returned the least text but took much longer. Two runs per method do not establish a universal ranking or the lowest model bill. Known-correction execution does not measure discovery of unknown defects or independent LLM reasoning.

For future comparisons, freeze the source, target and acceptance criteria; retain request/response ordering and gross elapsed time, with operator waits reported separately. Label scripted corrections and measured metrics accurately. Do not present response bytes as tokens or cost. Manual benchmark gates, initial-deck rendering and mandatory per-page changes are not routine workflow requirements.

### Layout checks and visual exceptions

Use a targeted check when automatic evidence is missing, failed or stale, or when a different scope is needed:

```json
{"documentId":"<document UUID>","layoutCheck":true,"checks":"layout","slides":[1,3,5],"detail":"summary"}
```

返回 `layoutAudit`，包含按页的超界量（pt）、可能重叠的对象对、无法解析的几何、`pendingSlides` 与 `screenshotRequiredNow`。未声明的交叠触发 `review_overlap_before_next_batch`，由调用方立即核查和修正；检查器不会自行截图、移动对象或删除内容。已知设计叠放可传 `allowedOverlapPairs: [[keyA,keyB]]`，key 来自当前查询，并提供 `expectedRevision`。省略例外参数时，只有页面、布局依赖及 generation 未变才复用声明；显式 `[]` 清除本次所查页面的声明。

省略 `slides` 即检查全稿。报告保存在 `worklistPath`；`ppt_status.documents[].layout` 可恢复每页检查状态、待办和下一步。修改不相关页面不再让所有页面检查过期；母版、主题等依赖变化会使相关页面过期。后续编辑仍需要最终整稿复查。文件模式不会启动 Office；`checks:"layout"` 跳过重复的结构检查，按所选页解析对象，首次恢复仍验证持久检查点。原生模式保留整稿身份守卫。`checks` 默认 full；也可以复用下面的正式成品渲染读回完成最终检查。

普通排版修改优先使用文件批量流程。继承母版位置的顶层文本占位符，可在 `set_geometry` 中明确提供 `left/top/width/height/rotation` 五项，写入完整绝对坐标；缺项返回 `GEOMETRY_REQUIRES_FULL`，不会猜测继承值。组合子对象仍不支持此路径。

最终可调用 `ppt_render({documentId,reviewId,slides:[...],allowOffice:true,layoutCheck:true,detail:"summary"})`：复用对已提交字节的完整只读原生快照，返回整稿 `layoutAudit` 与 `reviewBundlePath`，无需额外读取一次整稿。结构检查、输出哈希校验和原生读回仍会执行；完整对象信息保存在 review bundle 中，代码可按需读取，不必全部返回给 LLM。即使只渲染部分页面，布局检查也覆盖整个快照，但最终验收仍要求所有页面都有视觉证据。`layoutCheck` 在 render 中默认 false，只允许与 reviewId 一起使用；CLI 对应 `render --layout-check true`。

文件模式检查对象边框；原生读回还可测量普通未旋转文本的实际边界，报告 `TEXT_OUTSIDE_FRAME`。图表/SmartArt 内部、线宽、阴影与组合子对象仍未完整覆盖。旋转对象采用保守外接矩形，需核查设计意图；clear 不等于视觉验收。普通查询用 `ppt_inspect(detail:"summary")`，需要字体 runs 或图表数据时再取 full。CLI 支持 `--checks layout --detail summary --slides '[1,3]'`，设计例外通过 stdin JSON 传入。

异常页可直接调用 `ppt_render({documentId,expectedRevision,slides:[...],allowOffice:true,detail:"summary"})` 预览任务检查点，不必先发布临时成品；文件会话在 open/build/compose 时也必须允许 Office。草稿预览不计入最终验收。正式成品继续使用 reviewId；分批渲染的页面证据按同一输出哈希累积。完成全稿检查、修复已知问题并查看所有最终预览后，调用 `ppt_finish({reviewIds:[...],requireAccepted:true})`。检查失败会保留修正机会；不可解析几何记录为需要视觉接管的页面。省略严格验收时仍可清理资源，但 `acceptance:incomplete` 不能当作验收通过。没有预览的 caller confirmation 不再写成 visualReviewed。

### Advanced objects and visual relayout

- `ppt_build` 的 `type: "chart"` 支持 `bar/line/area/pie/doughnut/radar/combo`。系列使用 `name/labels/values`；组合图每个系列指定 `type: "bar" | "line" | "area"`，可用 `secondaryAxis: true` 指定副轴。所有系列的分类必须一致；饼图/环图只接受一个非负且总和大于零的系列。数据以可编辑的内嵌 XLSX 保存。原生打开只接受图表引用的静态工作簿，拒绝公式、宏、外链数据及其他嵌入对象。
- `ppt_build` 和 `ppt_compose.contentDeck` 可指定 `theme: "executive" | "editorial" | "academic"`。三套原创配色提供文字、背景和图表默认样式，显式指定的样式优先。设计参考 [Microsoft 演示文稿模板](https://powerpoint.cloud.microsoft/create/en/presentation-templates/)和[可读性建议](https://support.microsoft.com/en-us/powerpoint/tips-for-creating-and-delivering-an-effective-presentation)，不包含下载的整套网络模板。
- 原生副本中，`convert_to_smartart` 将已查询的顶层文本框转换为 SmartArt，`layout` 可选 `process/cycle/hierarchy`；`set_smartart_text` 按 `nodeIndex`（从 1 开始）修改节点文字。布局必须在本机 Office 目录或**同一任务已打开的模板**中可用；不可用时返回 `SMARTART_LAYOUT_UNAVAILABLE`，不会替换成其他布局。当前真实验收使用模板中的 hierarchy 布局，未证明本机 process/cycle 创建可用。
- `ppt_relayout` 对同一页显式选中的对象规划 `two-column/grid/stack` 排列，默认 `dryRun: true`；返回新旧位置及溢出/未选中对象重叠提示。取得新鲜引用后，以新的 operationId 和 `dryRun: false` 应用。保留文字和数据，不删减内容。图表、表格与 SmartArt 不写普通形状旋转属性。`two-column` 与 `grid` 当前都按双列等宽排列，`stack` 为单列。
- SmartArt 保持原生对象，图表工作簿和 SmartArt 依赖可随模板合并保留；这不等于支持任意动态链接、动画或 SmartArt 内部结构重建。长节点文字不自动缩小，仍可能溢出，必须查看最终预览。

重排参数示例（引用来自本次 `ppt_inspect`，单位 pt）：

```json
{
  "documentId": "<document UUID>",
  "expectedRevision": 0,
  "operationId": "layout_preview",
  "slide": 1,
  "titleRef": "<fresh title targetRef>",
  "targetRefs": ["<fresh body targetRef>", "<fresh chart targetRef>"],
  "layout": "grid",
  "theme": "executive",
  "dryRun": true
}
```

CLI 同样支持 `ppt-task relayout --task <id> --stdin`，标准输入为上述 JSON。修改后重新提交、渲染并查看成品，预演成功不是视觉验收。

### 双任务防串台

`ppt_status.bindings` 保留 taskId、documentId、原稿绝对路径和 SHA-256；原生模式另记录工作副本路径、Node 工作进程及可确认的 PowerPoint 进程身份（PID、创建时间、可执行文件路径）。这是历史记录；实时请求另行检查，不把历史 PID 当作当前存活证据。共享或无法唯一识别的 Office 会话可能没有可确认的 Office 身份，此时保持文稿级权限。

修改必须同时匹配任务、文稿、后端、revision、generation 和对象指纹；跨任务引用返回 `CROSS_TASK_REFERENCE`。原生 IPC 使用独立私有标识，调用前检查工作进程身份，文稿操作核对实际打开路径和可确认的 Office 身份。同一用户同时只允许一个原生任务持有 Office 租约，第二个收到 `NATIVE_BUSY`；两个纯文件任务仍可各自编辑独立副本。不会通过切换活动窗口选择 PPT，也不强制结束其他任务的 PowerPoint。

### 模板接稿与继续加页

`ppt_compose` 默认保留模板全部现有页，将旧稿按顺序接在后面。模板文件和旧稿均只读，新文稿进入文件模式任务，再由 `ppt_commit` 另存到新路径：

```json
{
  "templatePath": "inputs/template.pptx",
  "templateSlide": 2,
  "sourcePath": "inputs/content.pptx",
  "beautify": true,
  "copyDecorations": true,
  "fontFace": "Microsoft YaHei",
  "textColor": "17365D",
  "titleSize": 30,
  "bodySize": 20,
  "margin": 36,
  "allowOffice": true,
  "operationId": "compose_01"
}
```

`templateSlide` 是背景/版式参考页，默认 1；封面不适合作为正文模板时选正文页。可用 `sourceSlides: [2,3]` 选择旧稿页面；默认全部，禁止重复。总页数最多 100。

继续新增页面：将上次输出作为 templatePath，改传 `contentDeck: {"slides":[...]}`，结构与 ppt_build 的 deck 相同；sourcePath 和 contentDeck 必须且只能提供一个。原有页保持原顺序，新增页仍使用指定参考页的背景。CLI 对应 `compose`，可用 `Get-Content compose.json -Raw | node src/cli.js compose --task TASK_ID`；JSON 文件使用上述工具参数，任务须先 start。

模板复用包含母版/版式、参考页直接背景和默认复制的无文字、非占位装饰形状。模板页内有文字的页脚/品牌不会自动复制，可放在母版中复用。旧稿对象缩放到目标页面留白内；beautify 统一文本框和表格文字，不删除内容或保证自动重排。返回 composition.warnings，包括大对象遮挡背景、文字溢出风险和默认颜色回退。动画及无法保留的跨页链接会拒绝导入；复杂图表、SmartArt、特殊母版仍需单独验收。当前真实样本覆盖文本、普通表格与备注。

### 背景、字体和形状

先用 ppt_inspect（kind: "background"）取得背景 targetRef，然后批量 ppt_apply。旧的 `color: "FFFFFF"` 保持兼容；新用法在 set_slide_background 中传 fill：

| 效果 | fill 示例 |
|---|---|
| 纯色 | `{"type":"solid","color":"EDF4FA","transparency":10}` |
| 双色线性渐变 | `{"type":"gradient","startColor":"DDEEFF","endColor":"FFFFFF","angle":35}` |
| 图片拉伸铺满 | `{"type":"image","path":"inputs/background.png"}` |
| 原尺寸平铺纹理 | `{"type":"texture","path":"inputs/texture.png"}` |
| 图案 | `{"type":"pattern","pattern":"diagonalCross","foreground":"CDDDEE","background":"FFFFFF"}` |
| 恢复母版背景 | `{"type":"inherit"}` |

除 inherit 外支持 transparency 0–100，默认 0。图片/纹理只接受本地 PNG/JPEG，单文件最多 20 MiB；纹理从左上角按 100% 比例平铺。11 种图案名：horizontal、vertical、cross、downDiagonal、upDiagonal、diagonalCross、dots5、dots10、dots20、dots25、dots50。暂不包含多色/径向渐变、自定义图片裁剪及母版本身的编辑。OOXML 与 COM 的渐变栅格结果不承诺逐像素一致。

`set_style` 支持 fontFace、fontSize、color、bold、italic、underline；后三者可 true/false 开关。`set_geometry` 支持文本框/形状的 left、top、width、height（pt）与 rotation（度）。新建形状支持 flipH、flipV 和 rotation，文本和带文字形状支持粗体、斜体、下划线。

### A 路线最终成品预览

Follow the [recommended LLM workflow](#recommended-llm-workflow): build/compose/open(file, allowOffice:true) → batch inspect/apply with page reports → targeted refinement → commit → render(reviewId, layoutCheck:true, detail:"summary") → inspect all final pages → finish(reviewIds, requireAccepted:true). Use a separate whole-deck validate only when that evidence is needed independently; do not duplicate the complete readback by default. Output SHA-256, revision and image hashes bind to the reviewId. `visualReviewed` requires all-page preview evidence and caller confirmation; images are not automatically approved. Offline status is historical evidence, not a fresh check of external output files.

本项目提供官方标准 MCP（Model Context Protocol）协议服务入口：

- **服务端实现**：[`src/mcp.js`](src/mcp.js)，基于 `@modelcontextprotocol/server`（v2.0.0）与 STDIO 本地管道传输；
- **暴露的 14 项业务工具**：`ppt_start`, `ppt_open`, `ppt_inspect`, `ppt_apply`, `ppt_commit`, `ppt_validate`, `ppt_close`, `ppt_finish`, `ppt_status`, `ppt_diagnose`, `ppt_build`, `ppt_compose`, `ppt_relayout`, `ppt_render`；
- **Agent Skill 指南**：详见 [`skills/ppt-editor/SKILL.md`](skills/ppt-editor/SKILL.md)，指导 Agent 严格遵循“打开 -> 检查与版本绑定 -> 批量修改 -> 另存发布 -> 验收完成”的受控生命周期。

## CLI 工具调用示例 (`ppt-task`)

下面使用 PowerShell 与 JSON 标准输入传参；输入文稿由调用者显式提供，输出必须为新路径。

```powershell
$inputPath = (Resolve-Path ./sample.pptx).Path
$task = node src/cli.js start | ConvertFrom-Json
$opened = node src/cli.js open --task $task.taskId --path $inputPath --mode file --op open_1 | ConvertFrom-Json
$view = node src/cli.js inspect --task $task.taskId --doc $opened.documentId --kind text | ConvertFrom-Json
$target = $view.objects | Where-Object { $_.text -match 'Original' } | Select-Object -First 1
if (-not $target) { throw 'No matching inspected object.' }
$payload = @{
  documentId = $opened.documentId
  expectedRevision = $view.revision
  operationId = 'apply_1'
  operations = @(@{
    type = 'replace_text'; targetRef = $target.targetRef
    search = 'Original'; replacement = 'Revised'; expectedMatches = 1
  })
} | ConvertTo-Json -Depth 6
$edited = $payload | node src/cli.js apply --task $task.taskId --stdin | ConvertFrom-Json
$committed = node src/cli.js commit --task $task.taskId --doc $opened.documentId --rev $edited.revision --out ./output.pptx --op commit_1 | ConvertFrom-Json
node src/cli.js validate --task $task.taskId --doc $opened.documentId
# 检查输出和审阅记录后，再结束任务：
@{ reviewIds = @($committed.reviewId) } | ConvertTo-Json | node src/cli.js finish --task $task.taskId --stdin
```

也可使用 `--operations` 或 `--operations-file`；不支持 `--ops`。修改与发布必须传入对应的 `--rev`。

## MCP 连续任务

服务启动时自动创建首个任务。先使用 `ppt_status` 取得当前 taskId；完成 `ppt_finish` 后调用 `ppt_start`，传入新的 UUID：`{"taskId":"<new-uuid>"}`。重试启动必须复用同一 UUID，不会重复创建任务。当前任务未完成时拒绝切换。

`ppt_status` 可传入历史 `taskId` 和可选 `operationId` 查询回执。已结束的任务不会重新变为可编辑状态；旧对象引用不能用于新任务。

客户端应直接以 Node 可执行文件启动 `src/mcp.js`，工作目录设为项目根目录，避免包管理器额外输出干扰 STDIO。

### Codex 项目级接入

在已信任的项目根目录创建 `.codex/config.toml`，将下例中的 Node 和项目路径替换为本机实际绝对路径。Node 路径可用 `node -p process.execPath` 查询。该配置只作用于本项目：

```toml
[mcp_servers.ppt_editor]
command = "C:/Program Files/nodejs/node.exe"
args = ["C:/path/to/ppt-editor/src/mcp.js", "--base-dir", "C:/path/to/ppt-editor/work/codex-tasks"]
cwd = "C:/path/to/ppt-editor"
startup_timeout_sec = 15
tool_timeout_sec = 180
enabled = true
```

在项目根运行下列 PowerShell 命令，将 Codex 的项目 Skill 发现目录链接到现有 Skill 文件。若目标已经存在，先检查其内容，不覆盖：

```powershell
New-Item -ItemType Directory -Path .agents/skills -Force | Out-Null
New-Item -ItemType Junction -Path .agents/skills/ppt-editor -Target (Resolve-Path skills/ppt-editor).Path
codex mcp get ppt_editor --json
```

本机配置、Skill 链接和 `work/` 验证产物均被 Git 忽略；公开维护的 Skill 仍是 `skills/ppt-editor/SKILL.md`。配置成功后，在 Codex 的 MCP 设置中刷新/重启连接；若当前任务仍未出现工具，重新启动 Codex 并回到该项目。配置读取、工具发现、实际业务调用须分别验证。

此前已验证 Codex 能发现旧版工具和项目 Skill；新版为 14 个工具，刷新后请用 `ppt_diagnose` 核对版本和 capabilities。两项真实 STDIO 测试证明：输入 EOF 后，空闲原生会话被清理且可从检查点恢复；已接收的编辑完成并持久化回执后再清理。运行方式：

```powershell
node --test --test-concurrency=1 tests/mcp-disconnect.test.js
```

这些测试仅使用虚构样本，在检测到已有 PowerPoint 会话时跳过。覆盖的是允许服务完成清理的受控 EOF，不代表客户端强制终止、系统崩溃或所有 Office 异常都能无损恢复。当前桌面任务刷新后的两轮业务调用与预览验收仍待完成。

配置依据：[Codex MCP](https://learn.chatgpt.com/docs/extend/mcp?surface=cli)、[本地 Skill 发现](https://learn.chatgpt.com/docs/build-skills)。

## 候选版验证边界

Validate generated bytes and package contents before delivery; structural validation is neither full ECMA schema validation nor visual review. File-mode final previews use `ppt_render` with the committed reviewId; draft previews omit reviewId and require the current expectedRevision. Office permission is required at open/build/compose and render. Draft previews do not satisfy final acceptance. Geometry generation creates new slides or appends them through composition; arbitrary shape insertion into an existing slide remains unsupported.

Privacy checks cover candidate paths and staged file contents; passing only means no configured rule matched and still requires manual review. Production FileEngine uses the JSZip optimized writer. Historical plans, process notes and one-off experiments are kept in the ignored local `TEMP/` archive; private decks, previews and run data remain in ignored input/work/output directories. These local artifacts are not required to install or run the repository.

自行车创作对照可运行 `node --expose-gc scripts/benchmark-bicycle.js`。脚本仅生成虚构新文稿，检测到已有 PowerPoint 时停止；A 使用正式 buildDeck 的翻转能力，B 使用独立 COM 实验脚本创建原生形状；MCP 支持新建图形，尚不支持向既有页直接插入新形状。两条路径各运行 3 次，均包含 50 个形状、元数据清理、最终文件只读重开、逐对象核验、1600×900 渲染和自有进程退出确认。输出及分阶段计时位于新建的 `outputs/bicycle-ab-*` 目录。完整耗时包含每轮启动与退出；不包含模型思考、MCP 往返，也未控制操作系统缓存。

## 目录结构

- `src/contracts.js`：参数 Schema 与操作定义
- `src/file-engine.js`：基于 OOXML 的纯文件解析、编辑与检查点管理
- `src/native-host.js` & `src/native-worker.js`：PowerPoint COM 隔离执行与安全退出的原生工作进程
- `src/task.js`：TaskHost 统一任务层、跨进程互斥锁与 `targetRef` 上下文管理
- `src/cli.js`：面向终端与脚本的命令行桥接入口
- `src/mcp.js`：基于官方 SDK 的标准 STDIO MCP 协议服务端入口
- `src/ooxml.js`：OOXML 规范校验、DOM 序列化与流复用优化写包
- `skills/ppt-editor/SKILL.md`：面向 AI Agent 的规范调用 Skill 描述
- `scripts/check-privacy.js`：发布前隐私合规扫描工具
- `docs/benchmarks/methodology.md`：通用性能实验方法与证据边界


## 正常保存护栏

原生保存前检查目标是否被文件、目录或链接占用，并探测父目录可写性；检查点预检发生在实际修改之前。文件与原生输出均先生成候选、验证，再防覆盖发布。其他进程在预检后抢占输出路径时，发布仍拒绝覆盖。

冲突时选择新的输出路径；结果未知时先查询状态和检查点，不更换 operationId 盲目重试。系统保留原生调用超时与退出观察，不自动点击 Office 对话框，也不全局关闭共享应用提示。护栏可防止已知的路径冲突，不能保证 Office 的所有内部异常都不弹窗。

原生关闭必须收到成功回执才写入 closed；finish 在关闭文稿期间保留全部检查点，确认宿主清理后才删除恢复数据。清理失败或结果未知返回 `CLEANUP_UNCONFIRMED`，任务保持 active，`ppt_status` 可读取 `cleanupReport`，文稿关闭错误另记为 `closeErrors`。已确认保留的外部 Office 会话不作为清理失败。

未确认的宿主清理记录会阻止再次关闭、finish 或启动新原生宿主，重启命令不会自动解除。先核对进程和保留数据；需要恢复时，原生模式可从 meta.json 指向的检查点副本启动新任务，文件模式仍可另存当前修订。不要删除失败记录来强制通过清理。

文件模式可在打开时及 `ppt_validate` 时均显式设置 `allowOffice: true`，再使用 `nativeReadback: true`。系统在文稿锁内读取最新持久修订、生成独立候选，以只读且无窗口方式由 PowerPoint 打开检查并关闭；返回 `nativeValidation`，含候选 SHA-256、revision、generation、原生快照及 readOnly 标记。成功后删除临时候选，失败时保留。此校验不改变编辑修订，也不等于视觉审阅；它的哈希仅对应本次生成的候选，后续提交仍需核对实际输出。
