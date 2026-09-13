# ppt-editor

[English](README.md) | 简体中文

通过 MCP 服务或命令行创建、编辑 PowerPoint 演示文稿。支持批量修改 `.pptx` 文件、检查幻灯片布局，并使用 Microsoft PowerPoint 渲染保存后的文件以供审查。原文件始终保留，修改结果另存到新路径。

**版本：** `0.4.0-rc.7`，候选发布版。

**运行 MCP 服务的电脑必须安装 Windows x64 和 Node.js 24.x。** 当前支持范围为 `>=24 <25`，项目不内置 Node.js。原生编辑和渲染还需要安装 Microsoft PowerPoint。

[安装与配置](#安装与-mcp-配置) · [推荐工作流](#推荐工作流) · [工具与示例](#工具与示例) · [功能限制](#功能限制)

## 功能

| 类别 | 支持的操作 |
|---|---|
| 文字 | 替换文字；设置字体、字号、颜色、粗体、斜体和下划线；编辑演讲者备注 |
| 对象 | 调整位置、尺寸和旋转角度；编辑表格单元格；生成包含可编辑形状、线条和箭头的幻灯片 |
| 背景 | 纯色、双色渐变、图片、纹理、图案和继承背景 |
| 模板 | 保留模板页面，追加现有演示文稿或新生成的页面 |
| 图表 | 创建柱形图、折线图、面积图、饼图、圆环图、雷达图和组合图，使用内嵌静态数据 |
| SmartArt | 通过 PowerPoint 将文字转换为 SmartArt、编辑节点文字，受可用布局限制 |
| 审查 | 报告疑似重叠和越界对象；渲染草稿或保存后的文件；记录审查状态 |

提供两种编辑模式：

- **`file`**：直接编辑 OOXML 文件包，适合支持范围内的批量操作。编辑和几何检查不会启动 PowerPoint。
- **`native-copy`**：通过 PowerPoint COM 编辑独立副本，适合需要 Office 原生功能的操作。不会接管用户正在编辑的演示文稿。

两种模式的渲染都需要 PowerPoint。因此，可以用文件模式批量编辑，再用 PowerPoint 渲染最终结果，无须将全部编辑操作转为 COM 调用。

## 安装与 MCP 配置

### 环境要求

请在运行 MCP 服务或 CLI 的电脑上准备以下环境：

| 组件 | 要求 | 用途 |
|---|---|---|
| 操作系统 | Windows x64 | 所有安装，包括仅使用文件模式的情况 |
| Node.js | **24.x**（`>=24 <25`） | 运行 MCP 服务、CLI 和安装脚本 |
| 原生模块编译工具 | Python 和 MSVC C++ 编译工具 | 在 `npm ci` 期间编译 `winax` |
| Microsoft PowerPoint | 已安装的桌面应用 | 原生编辑、渲染和 Office 集成测试 |

安装前检查 Node：

```powershell
node --version
node -p "process.platform + ' ' + process.arch"
```

输出应为 `v24.x.x` 和 `win32 x64`。Node 22、Node 25 及更高版本均不在当前支持范围内。安装脚本会检查平台和 Node 主版本，不符合要求时停止安装。

### 从源码安装

```powershell
git clone https://github.com/linnnn89/ppt-editor.git
cd ppt-editor
npm ci
```

`npm ci` 只安装项目依赖，不会安装 Node.js、Python、MSVC 或 PowerPoint。当前源码安装方式会编译原生模块，即使后续仅使用文件模式，也需要编译工具。

安装脚本会编译 `winax` 并执行原生模块冒烟测试。构建使用 Node 24.18.1 头文件，以避开 24.19.0 头文件的已知扩展模块问题；这不会替换本机 Node 运行时，也不要求运行时必须为 24.18.1。具体配置见 [build-native.js](scripts/build-native.js)。

### 连接 Codex

仅在某个项目中使用时，将以下配置加入该项目的 `.codex/config.toml`；Codex 只加载受信任项目的配置。需要跨项目使用时，写入用户级 `~/.codex/config.toml`。目录或文件不存在时创建；已有 `ppt_editor` 时更新同名配置表并保留其他设置。配置作用域见 [Codex MCP 官方指南](https://learn.chatgpt.com/docs/extend/mcp?surface=cli)。

将 Node 可执行文件和仓库路径替换为本机路径。`command` 必须指向 **Node.js 24.x 的可执行文件**；终端中的 Node 版本正确，并不代表此处配置的可执行文件也是同一版本。

```toml
[mcp_servers.ppt_editor]
command = "C:/Program Files/nodejs/node.exe"
args = ["C:/path/to/ppt-editor/src/mcp.js", "--base-dir", "C:/path/to/ppt-editor/work/codex-tasks"]
cwd = "C:/path/to/ppt-editor"
startup_timeout_sec = 15
tool_timeout_sec = 180
enabled = true
```

直接用 Node 启动服务，避免包管理器的输出干扰 STDIO 通信。启动 MCP 服务本身不会启动 PowerPoint。

### 添加 Agent Skill

[Skill](skills/ppt-editor/SKILL.md) 提供操作入口和三份专项手册。安装 Skill 不会安装项目依赖或注册 MCP 服务。用户和本地 AI Agent 均可按照以下命令操作。

按 [Codex Skill 官方指南](https://learn.chatgpt.com/docs/build-skills)选择一种发现范围：

| 范围 | Skill 位置 | 生效范围 |
|---|---|---|
| `project` | `<项目>/.agents/skills/ppt-editor` | 当前项目及其子目录；下方示例使用本仓库 |
| `user` | `~/.agents/skills/ppt-editor` | 该用户的各个项目 |

在仓库根目录运行下面的命令，将 `$skillScope` 设为 `project` 或 `user`。命令将 Skill 子目录作为 Windows 目录联接，指向仓库中维护的版本；已有目标会保留并显示出来，供检查后决定如何更新。

```powershell
$skillSource = (Resolve-Path -LiteralPath 'skills/ppt-editor' -ErrorAction Stop).Path
$skillScope = 'project'
$skillParent = switch ($skillScope) {
    'project' { Join-Path (Get-Location).Path '.agents/skills' }
    'user' { Join-Path $HOME '.agents/skills' }
    default { throw 'Choose project or user for skillScope.' }
}
$skillDestination = Join-Path $skillParent 'ppt-editor'
$existingSkill = Get-Item -LiteralPath $skillDestination -Force -ErrorAction SilentlyContinue
if ($existingSkill) {
    $existingSkill | Format-List FullName, LinkType, Target
    throw 'Destination already exists. Inspect it before updating.'
}
New-Item -ItemType Directory -Path $skillParent -Force -ErrorAction Stop | Out-Null
New-Item -ItemType Junction -Path $skillDestination -Target $skillSource -ErrorAction Stop | Out-Null
Get-Item -LiteralPath $skillDestination | Format-List FullName, LinkType, Target
Get-Content -LiteralPath (Join-Path $skillDestination 'SKILL.md')
```

目录联接会跟随仓库更新，已有正确联接无须再复制一份。请保留仓库的配置路径。如果已有目标是独立副本，先备份到 Skill 发现目录之外，再比较并更新完整的 `skills/ppt-editor` 文件夹，包括 `references/`。只复制 `SKILL.md` 会缺少操作手册。

### 验证安装

1. 核对命令显示的安装目录和联接目标。安装目录应包含 `SKILL.md` 及 `references/workflow.md`、`references/tools.md`、`references/recovery.md`。
2. 确认 Codex 的可用技能中出现 `ppt-editor`。文件变化被检测后仍未出现时，重启 Codex；用户级安装还应在另一个项目中核对。
3. 刷新 MCP 连接并调用 `ppt_diagnose`，确认 `version` 与 `diskVersion` 一致、`restartRequired:false`，并且当前 14 个工具可用。用户级 Skill 要跨项目调用这些工具，还需要用户级 MCP 配置。

Skill 被发现、MCP 连接成功、实际文稿编辑成功是不同的检查。更新磁盘文件不会自动重载已启动的 MCP 进程。本地项目 MCP 配置、Skill 联接和任务数据均已排除在 Git 提交范围之外。

## 推荐工作流

**批量编辑，按页读取检查报告，精修有问题的页面，最后审查保存后的完整演示文稿。**

1. **检查内容并批量编辑。** 优先使用文件模式。短文稿可以一次批量处理，较大任务按布局或操作数量分批。使用 `ppt_inspect` 的 `detail:"summary"` 获取摘要，在调用 `ppt_apply` 前取得当前对象引用。单次检查最多返回 300 个对象，单次编辑最多接受 500 个操作；需要时处理分页。
2. **收到结果后再继续。** build、compose 和 apply 默认返回逐页布局报告。等待真实响应，确认编辑和后续检查是否成功。已检查的页面没有问题时，继续处理下一批。检查失败不会撤销已经完成的编辑。
3. **发现问题立即视觉核查。** 如果报告或内容提示非设计性重叠、越界或文字裁切，立即渲染并查看对应页面。修正已确认的问题，重新检查该页后再继续。不必每次普通编辑后都截图。只有设计本身能够支持这一判断时，才将重叠声明为有意设计。
4. **集中修复相关页面。** 重新查询有问题的页面，取得最新引用。多个已明确、互不依赖的修复可以合并到同一批次。复用仍然有效的检查结果，避免重复验证或重新加载全文稿。未解决的问题和过期结果应继续保留在待办记录中。
5. **审查最终文件。** 提交到新路径后，使用返回的 `reviewId` 调用 `ppt_render`，设置 `layoutCheck:true` 和 `detail:"summary"`。这会在渲染所需的回读过程中完成最终全稿布局检查。查看每一页最终结果，按需使用缩略图总览和局部放大。继续修改后，应重新提交并审查新文件。只有代码检查和视觉审查均完成，才能以 `requireAccepted:true` 结束任务。

逐页报告会保留当前有效证据，并标明尚未完整检查的内容；几何结果为 `clear` 不代表文字边界已经完整测量。覆盖范围、证据有效性及最终验收规则见[工作流手册](skills/ppt-editor/references/workflow.md)。

MCP 支持取消处理、可选进度通知和阶段计时；文件检查点会复用已验证的落盘部件。取消、进度和检查点故障的细则见[连接与恢复手册](skills/ppt-editor/references/recovery.md#uncertain-results-and-cleanup)，计时字段和解释见[性能比较](skills/ppt-editor/references/workflow.md#benchmarking)。

`screenshotRequiredNow` 当前用于标记重叠报告；值为 false 并不表示没有其他视觉问题。代码检查无法判断设计意图，也无法完整评估可读性。生成了图片不等于已经查看过图片。

针对指定页面执行代码检查时，调用 `ppt_validate`，设置 `layoutCheck:true`、`checks:"layout"`、`slides:[...]` 和 `detail:"summary"`。预览文件模式草稿时，使用当前 `expectedRevision` 调用 `ppt_render`，不传 `reviewId`。草稿预览不计入最终验收。使用 Office 需要在 open/build/compose 和 render 调用中明确授权。

重叠声明、版本处理和最终审查的详细规则见 [Skill](skills/ppt-editor/SKILL.md)。

## 工具与示例

MCP 服务提供 14 个工具。CLI 提供同名命令，去掉 `ppt_` 前缀即可。

| 用途 | 工具 |
|---|---|
| 管理任务 | `ppt_start`、`ppt_status`、`ppt_finish`、`ppt_diagnose` |
| 打开或创建文稿 | `ppt_open`、`ppt_build`、`ppt_compose` |
| 检查与编辑 | `ppt_inspect`、`ppt_apply`、`ppt_relayout` |
| 验证与渲染 | `ppt_validate`、`ppt_render` |
| 保存与关闭 | `ppt_commit`、`ppt_close` |

MCP 服务会自动创建第一个任务。任务结束后，使用新的 UUID 调用 `ppt_start` 开始下一项任务。CLI 需要显式启动任务。已经结束的任务引用不能用于继续编辑。

### 在 PowerShell 中检查演示文稿

```powershell
$task = node src/cli.js start | ConvertFrom-Json
$opened = node src/cli.js open --task $task.taskId --path ./sample.pptx --mode file --op open_1 | ConvertFrom-Json
node src/cli.js inspect --task $task.taskId --doc $opened.documentId --slide 1 --detail summary
node src/cli.js close --task $task.taskId --doc $opened.documentId
node src/cli.js finish --task $task.taskId --reviews '[]'
```

此示例仅读取文件，不会启动 Office、修改源文件或批准输出。运行 `node src/cli.js --help` 查看命令和基本用法。复杂参数可以通过 `--stdin` 以 JSON 传入。

### 批量修改

将以下参数发送给 `ppt_apply`，用当前检查结果替换占位值。在数组中增加操作，即可一次修改多个对象或页面。

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

后续操作应使用返回的新版本号。几何尺寸使用磅，旋转角度使用度。对于继承几何属性的顶层占位符，修改变换属性时需同时提供 `left`、`top`、`width`、`height` 和 `rotation` 五个值。

### 将文稿追加到模板

将以下参数发送给 `ppt_compose`：

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

所有模板页面都会保留，源文稿页面追加在其后。`templateSlide` 指定背景和布局参考页。若要追加新生成的页面，将 `sourcePath` 替换为 `contentDeck`；两者必须且只能提供一个。请查看合成过程中的警告：`beautify` 会应用字体排版和页边距调整，但不保证所有内容都能完整容纳。

### 编辑参数参考

- 文字、表格、几何属性、主题、图表和工具参数：[contracts.js](src/contracts.js)。
- 背景填充字段：[contracts.js](src/contracts.js) 中的 `backgroundFillSchema`。图片和纹理支持最大 20 MiB 的本地 PNG/JPEG 文件；可用图案名称见 [background.js](src/background.js)。
- 布局规划：`ppt_relayout` 可对单页中选定的对象规划 `two-column`、`grid` 和 `stack` 布局。默认 `dryRun:true`，应先查看方案再应用。
- SmartArt：所用布局必须存在于 Office 或同一任务中打开的模板。回归测试验证了 `hierarchy`，其他布局取决于安装环境。

## 文件保护与恢复

编辑使用任务副本，对象引用绑定到文稿版本。`ppt_commit` 要求使用新输出路径，拒绝覆盖源文件或已有文件。原生任务只关闭自己持有的文档，并保留此前已经打开的 PowerPoint 会话。

| 返回结果 | 处理方式 |
|---|---|
| `REVISION_MISMATCH` 或引用过期 | 检查当前文稿，使用最新引用重新构造编辑请求。 |
| 响应缺失或结果不确定 | 使用 operationId 查询 `ppt_status`。先检查操作记录，再决定是否以相同 ID 重试相同请求。 |
| `NATIVE_BUSY` | 等待持有 Office 使用权的任务结束，不要关闭其他任务的 Office 会话。 |
| `CLEANUP_UNCONFIRMED` | 检查状态和保留的恢复记录，不要删除记录或假定清理已经成功。 |
| `acceptance:incomplete` | 输出尚未通过验收。资源清理与最终审查是两项独立结果。 |

`ppt_status` 读取已记录的状态，不会重新验证外部文件。新的 MCP 连接通常会创建新任务。要恢复未完成的任务，请用原来的 `--task` UUID 和 `--base-dir` 重启服务，再检查其状态。恢复和引用处理的详细规则见 [Skill](skills/ppt-editor/SKILL.md)。

## 功能限制

- 布局检查覆盖对象边界和疑似重叠。原生回读能够测量部分文字边界，但无法完整检查图表和 SmartArt 内部、组合内子对象坐标、描边及效果。
- 文字替换不会自动调整所有内容的尺寸和排版。较长的替换内容需要在渲染结果中检查。
- 可以在新生成的页面中创建形状；暂不支持向已有页面任意插入形状。线条使用固定坐标，不会跟随形状移动。
- 不支持编辑已有图表的数据系列。原生工作簿访问仅限静态图表数据；公式、外部链接、宏和无关嵌入文件会被拒绝。
- 不支持含宏、加密、DRM 保护或数字签名的文稿。OCR、动画编辑、3D 对象及音视频编辑也不在当前范围内。
- 恢复测试覆盖正常关闭 STDIO 和保留恢复记录的情况，不覆盖所有强制终止或 Office 故障。原生启动可能短暂改变窗口焦点。

## 开发与测试

```powershell
npm run check          # JavaScript syntax
npm run check:privacy  # Candidate files and staged content
npm test               # Unit and integration tests, including real Office operations
```

完整测试应在没有打开 PowerPoint 会话的桌面环境中运行，部分会话归属测试依赖这一条件。测试使用合成文稿。隐私扫描只检查已知模式和文件类型，不能代替对实际提交文件的审查。

性能比较方法见 [基准测试方法](docs/benchmarks/methodology.md)。不同方法应完成相同工作并采用相同验收标准；工具执行时间应与模型推理时间区分，响应字节数不能直接视为 token 用量或费用。

业务代码位于 `src/`，测试位于 `tests/`，验证脚本和测试素材位于 `scripts/`。版本变化见 [更新日志](CHANGELOG.md)。私人文稿、生成文件、任务记录和本地归档应放在 `private/`、`outputs/`、`work/`、`TEMP/` 等已被 Git 忽略的目录中。
