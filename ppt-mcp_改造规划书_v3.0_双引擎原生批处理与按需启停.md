# ppt-editor 实施基线｜2026-09-12 JS 新建项目修订

本节依据用户本轮明确要求和复核后的决定，优先于下方保留的 v3.0 历史规划。历史正文中的 Python、C#、旧仓库和“现有资产”不能作为当前项目实施事实。

## 当前范围与验收

- 在 `linnnn89/ppt-editor` 从零建立独立项目，自行维护的代码使用 JavaScript；用户已允许含原生二进制的 Node 扩展。
- Windows x64、已登录桌面、本地普通 `.pptx` 为第一条实际验证路径。Node 24 为当前构建基线；不改变全局环境。仅使用当前已有系统工具构建锁定的项目依赖。
- 保留 OOXML 文件编辑和 PowerPoint 原生副本批量编辑。用户原稿的直接附着仍按 P4 另设验收，不因发现运行中的 PowerPoint 自动附着。
- 交付物是供 Codex 调用的 Skill、CLI、MCP 和编辑引擎。没有独立 GUI 的规格；不建立网页/Electron 编辑器。
- 默认原稿保护、另存结果；修改、保存、验证、审阅与退出分别给出证据。未实现的对象组合必须预检拒绝，不能降级成截图。

## JS 模块与任务生命周期

短命 CLI 连接本任务的 Node 任务进程；核心领域逻辑和文件引擎先采用进程内模块。原生 COM 执行器使用独立 Node 子进程。CLI 与任务进程使用官方 MCP SDK 和受限本地管道传输；另提供按调用启动的标准 STDIO 桥接入口。协议握手交给 SDK，业务状态与 MCP 连接状态分开。

只在实际任务显式 start 后运行，同任务复用；finish 发布最新验收回执后清理，外部调用方核实退出。无任务不运行服务、托盘、计划任务或全局 MCP 空壳。status 离线读取且不自启。失联、租约到期、调用超时和崩溃采用独立处理，保留唯一检查点，不把超时视为验收通过。

Windows Job 仅接纳已知自有 Node 进程，使用明确的归属与退出观察；PowerPoint 不因进程名、PID 或 COM 创建成功被授予强杀权限。现有用户文稿必须保持。原生路径的消息处理、调用拒绝、进程共享及非激活行为均以 P0 实测确定。

## 第一轮能力表与阶段

| 阶段 | 实施内容 | 门槛 |
|---|---|---|
| P0 | 锁定依赖；安全任务启动/退出；小样例文件与原生文字、格式、几何修改；保存、重开、读回、故障与前台观察 | 不以 COM 连通代替编辑通过；依赖或原生行为不成立时记录阻塞，不自动换语言或改系统配置 |
| P1/P2 | 统一查询与对象引用；文本替换、文本样式、简单几何、固定表格/备注；文件候选层、原生检查点、持久回执、另存、验证、finish | 原稿安全、引用不过期、批次预检、部分失败可定位、退出可核验 |
| P3 | JS 新建文稿、模板修改、生成后继续编辑、渲染审阅修订；复杂 SVG/DrawingML roundtrip 单独标覆盖 | 旧项目能力不能冒充已经继承；生成文件需要独立打开和结构验证 |
| P4/P5 | 原稿附着、复杂分组/图表/动画/媒体及发布兼容矩阵 | 保持原 v3.0 的独立验收门槛，不因基础闭环完成宣称全面支持 |

文本替换默认保留既有 runs；跨不同格式 run 的替换需要明确格式策略。几何以 pt 为外部单位；分组内部坐标不隐式当作页面坐标。严格文件任务不启动 Office，包括验证环节。原生渲染是单独允许的能力，未运行的验证层按实报告。

依赖候选为官方 MCP SDK、ZIP/XML 库、PptxGenJS、winax 和 Koffi；实际版本记录于 `package-lock.json`。winax 需本机编译，已发现本机现有构建工具；能否满足 Node 24 和原生生命周期要求仍须验证。没有授权自动安装新的系统构建工具或推送远端。

实施和验证结果持续追加到 `docs/codex_worklog.md`。

---

# 历史规划：ppt-mcp 改造规划书
## v3.0｜Windows 专精：OOXML 文件编辑 + PowerPoint 原生批处理 + 任务级按需启停

**资料核查时间：北京时间 2026 年 9 月 12 日。**\
**用途：可供 Codex／开发者实施的架构决策、工作包和验收依据。**\
**状态：规划与静态核查，不是已实现功能，不包含本机 Windows／Office 性能实测。**

本版替代 v2.0 中“原生层首先仅作渲染、广泛原生编辑后置”的产品范围；保留其按需启停、原稿保护、源保真、结构化查询和真实退出验证。前面聊天中的示意延迟、语言性能排序、COM／OOXML 快慢预测均不作为本项目基准。文中工具、模块及状态名是拟议契约，不是现有命令。

本次重新核对的仓库主线：

| 仓库 | main 引用 |
|---|---|
| `linnnn89/ppt-mcp` | `b3073f8fcf5d2ebeda621a424097dde9dd771cea` |
| `linnnn89/WinCode` | `af9c1b6a7284cf8fa909badedcc1c05a34cc7ee1` |

核查覆盖已有 v2.0 计划、两仓引用及已读关键源码、Microsoft／OpenAI 官方文档、公开 COM 项目代码与一手 issue。未查看用户所说的具体 Mac 演示视频，未执行全仓审计，也未复现任何外部项目公布的测试结果。来源见文末。

---

## 1. 决策摘要

### 1.1 产品定位

> 让 Codex 像访问结构化文档对象一样编辑 PowerPoint：未打开的文件可直接修改 OOXML；需要原生实时编辑时，在明确绑定的 PowerPoint 文稿上批量执行 COM 操作。两条路径共享对象级命令、权限、验证和交付规则，仅在 PPT 任务期间运行，Codex 验收后关闭。

要实现的是“批量、精确、可验证”，不是复制视频中的鼠标动作。窗口是否可见、是否抢焦点、是否持久运行，是三个独立问题。

### 1.2 本版冻结的五项方向

1. **COM 是正式编辑引擎，而非仅渲染附属工具。** 第一轮可行性实验必须包含真实批量修改；首个 Windows 双引擎版本必须交付经过测试的原生编辑子集。
2. **OOXML 继续作为不启动 Office 的独立路径。** 不为了支持实时编辑而废弃现有 Python、DrawingML、roundtrip 和交付检查资产。
3. **共享命令，不共享不受控的写状态。** 一个文档会话同一时刻只允许一个写后端；切换必须经过检查点和重新读取，不能逐操作自动来回切换。
4. **按任务运行。** Codex 主动启动；同任务复用；交付后提示 Codex 验收；验收结束默认关闭；只有用户明确连续工作才保留。
5. **技术栈沿用 v2.0：Python 核心 + C#／.NET Windows 控制与原生 Host。** 不因未经实测的语言性能判断再次重写网关。TypeScript 不是性能禁区，但当前没有必要为它另加一层必需运行时；C++ 仅在测量证明需要时考虑局部使用。

---

## 2. 自我审核：这次修正什么，不再推断什么

| 先前表述或潜在误解 | 本版修正 | 实施影响 |
|---|---|---|
| COM 不适合普通文本框、位置和样式编辑 | 普通原生对象操作是 COM 的正常用途；性能需区分冷启动、批处理和逐次模型往返 | 增加原生 batch executor，不再把全部简单修改强制分流到 OOXML |
| Mac 上逐框变化足以确定 AX、AppleScript 或 Office.js | 现象不足以区分这些路径，也不足以恢复私有实现 | 借鉴用户体验目标，不把推测写成技术前提 |
| 一批 MCP 操作就等于一次 COM 调用或原子事务 | 一次高层请求内部仍可能有许多 COM 调用，也可能部分成功 | 单独统计互操作调用；设计 checkpoint、部分结果和失败恢复 |
| PowerPoint 已打开就优先使用它 | “有一个 PowerPoint 进程”不等于目标已绑定，更不等于拥有该进程 | 只对已授权、准确绑定的文稿选择 native 路径 |
| UIA 与 PowerPoint 原生对象树可以等同 | UIA 控件不保证与每个 Slide／Shape 一一对应 | 编辑定位用 COM／OOXML；UIA 仅诊断窗口和阻塞状态 |
| 文件修改后，保持 PowerPoint 打开即可看到新内容 | 已加载的文稿有自己的内存状态 | 文件转原生、原生转文件都显式保存／关闭／重读 |
| 完工返回“已关闭”即可 | 返回消息时进程可能仍活着 | 外部观察进程退出；状态查询不重启服务 |
| COM 是“更完整”就必然保真 | PowerPoint 保存也可能规范化文件或更新部分结构 | 原生输出做语义与视觉保留验证，不承诺部件字节不变 |

**已证实的外部事实：** OpenAI 官方在 2026-04-16 描述了 Mac 后台 computer use，包括观察、点击、输入和用户继续使用其他应用；微软提供 ShapeRange、TextRange 等原生对象接口；公开项目确有会话内串行 COM 执行实现。[S03–S07]

**证据不足：** 用户看到的那段演示具体走什么 API、是否使用脚本批处理、逐帧行为是否等于逐次模型决策、相对 Windows COM 快多少，均不能据现有材料判定。本项目也不能提前声称达到同等速度。

---

## 3. 外部项目与网友经验：采用什么、排除什么

| 来源 | 本次可确认的内容 | 采用 | 不照搬 |
|---|---|---|---|
| `linnnn89/WinCode` | 统一工具注册、字段校验、定向结构化结果，以及已有只读 UIA／原生适配思想 | 紧凑查询、明确身份、证据边界、资源生命周期、有界受理 | 不合并整仓，不带 Roslyn、MSBuild、源码索引；不把只读 UIA 宣传成已验证的后台写入 |
| `sbroenne/mcp-server-powerpoint` | `PresentationBatch.cs` 使用专用 STA、Channel 队列、消息泵／OLE message filter；有重开文稿和退出处理 | 串行原生执行、同会话复用、实际身份记录、真实 Office 集成测试 | 已读代码有无界队列、全局 Designer 注册表修改、强制可见等选择；这些不满足本项目默认策略 |
| 同项目 `CONTINUATION.md` | 作者记录了 Untitled 参数导致 Open→Save 失败、只测 Create 不足以发现问题等经历；文档保留不同阶段记录 | 保存／重开、异常、图表、焦点分开测试，不能以编译或单次创建成功代替验收 | 不把作者测试当成本项目实测，也不把历史记录当成全仓当前事实 |
| `Ayushmaniar/powerpoint-mcp` | README 列出运行中 PowerPoint 的实时读写、对象快照和复合内容工具 | 丰富对象快照、整页规划、减少格式修改往返 | 不采用任意 Python evaluate、前台公式编辑或自动扩权；README 中对其他库能力的笼统比较不能作技术结论 |
| `pptx-automizer` | 已读关系助手处理关联部件与资源复制 | 关系图闭包和未知内容保留测试 | 不整体换引擎，不照抄依赖具体路径假设的解析逻辑 |
| `python-pptx` 官方文本 API | text frame／paragraph／run 层次的赋值行为与格式继承不同 | run／字符范围级编辑、显式格式策略 | 不以整框 `.text = ...` 冒充无损替换 |
| Codex issue #34614 及评论 | 特定 Windows 构建下的 MCP 进程链残留报告；跟帖给出特定源码轮廓的 Job Object 实验 | 自有子进程归属、根进程先退出的回归、退出独立验证 | 不推断所有 Codex 版本都泄漏，也不把网友的修复建议当成已合入官方版本 |

这些资料的作用是形成实现和测试要求，而非“某个项目有星标或演示就直接采用”。尤其要区分**知道 PID**与**有权终止该进程**。[S02、S07–S11]

---

## 4. 用户可见的工作方式

### 4.1 三种会话模式

| 模式 | 编辑真相来源 | 场景 | 默认保护 |
|---|---|---|---|
| `file` | 不变源包 + 文件修改层 | 未打开文稿的后台修改、批量文件工作、生成与 roundtrip | 不启动 Office，原稿只读，另存 |
| `native-copy` | PowerPoint 内本任务工作副本 | 需要真实 PowerPoint 对象、实时排版和快速多轮编辑 | 使用专属副本；可隐藏或可见但不主动激活；检查点恢复 |
| `native-attached` | 用户明确指定的已打开文稿 | 用户要求“直接改我当前打开的这份 PPT” | 单独授权、精确绑定、不默认保存原稿、不默认关闭用户文稿；不承诺整体原子回滚 |

`native-copy` 中“副本属于任务”，不表示整个 PowerPoint 进程一定独占。进程可能共享，必须另核验应用 lease。[S12]

### 4.2 后台策略独立于会话类型

**严格文件策略：** 完全不创建 Office、浏览器或预览 UI，不使用键鼠、剪贴板。泛化的“后台改这个文件”默认选择此路径。

**原生无主动激活策略：** 允许明确需要的 PowerPoint 原生会话；不调用前台激活、选中、键鼠或剪贴板进行编辑。可见窗口不等于取得前台。对具体 Office 构建、操作和状态验证，不承诺任意插件／弹窗环境绝无干扰。

用户要求在 PowerPoint 内编辑或观看实时结果时，选择经过验证的 `native-copy`；用户要求修改自己的已打开原稿时，才进入 `native-attached`。不能因发现 POWERPNT.EXE 就自动附着。

`WithWindow=false` 只描述演示文稿窗口是否显示；不能把它解释为 PowerPoint 全部窗口、加载项和安全提示都不可能出现。[S13]

### 4.3 不做“表演式”编辑

普通 COM 修改不需要逐次选择文本框，也不需要不断切换当前页。模型可以一次提交多页计划，本地连续执行，窗口是否逐帧刷新由应用决定。

默认不为了模仿演示而调用 Select、Activate、GotoSlide 或抢鼠标。需要进度时返回“已改页／对象”和最新缩略图；需要可见效果时由用户自行查看文稿。自动翻页演示不是核心编辑能力，也不能成为后台路径的隐含副作用。

---

## 5. 架构与职责：共享核心，隔离原生故障

```text
Codex：PPT Skill、内容规划、语义验收
        │ 仅 PPT 操作任务时启动
        ▼
Windows TaskController / CLI（C#/.NET）
  task 身份、受保护 IPC、MCP 客户端、进程归属、退出观察
        │ 标准 STDIO MCP
        ▼
Python MCP Core
  工具契约、能力路由、文档会话、对象引用、批次、回执与交付
        │
        ├── FileEngine
        │   OOXML 局部补丁、现有生成/roundtrip、关系与保真检查
        │
        └── NativeAdapter ── 独立 RPC ── OfficeHost（C#/.NET）
                                      STA、COM batch executor
                                      原生对象快照、检查点、渲染
                                      按需只读 UIA/Win32 诊断
```

Controller 管理 Core 和 OfficeHost 的生命周期；NativeAdapter 与 OfficeHost 使用独立的内部通信，控制通道不能因等待编辑结果而阻塞。不得形成“Controller 持锁等待 Core，Core 又等待 Controller 创建 Host”的循环等待。

### 5.1 模块职责

| 模块 | 职责 | 不负责 |
|---|---|---|
| TaskController | 启动、权限作用域、本地连接、OS 归属、中止和真实退出 | 不实现文本或图表业务规则，不判断论文／讲稿内容是否完成 |
| ToolRegistry | 单一输入／输出契约、组合校验、能力说明、工具分发 | 不暴露任意 COM 成员或脚本执行 |
| DocumentSession | documentId、模式、revision、来源、原生 generation、检查点 | 不让同一 revision 被两个后端同时写 |
| TargetResolver | 从范围查询定位对象，产出精确引用和歧义 | 不靠当前选中对象或全局 active presentation 猜目标 |
| BatchPlanner | 预检、依赖排序、同页聚合、读回规则、风险与恢复方式 | 不重新调用模型，不自行改用户事实 |
| FileEngine | 文件部件修改、源保留、关系闭包和输出组包 | 不在原生文稿仍打开时改同一路径的 ZIP |
| NativeExecutor | 已批准命令的 COM 解释执行、STA、有限重试、局部读回 | 不自发路由到 GUI，不绕过安全设置 |
| Validator / Publisher | 内容和保留检查、固定交付物与报告、形成验收包 | 不把成功导出图片当成已视觉审阅 |

### 5.2 为什么不再调整语言

本次修正的是**COM 的业务地位和批处理方式**，不是要再换一次基础设施。Python 保留现有资产，.NET 集中承载 Windows 进程与 Office 对象操作。COM 批次在 OfficeHost 内解释执行，不经过 Python 对每个属性往返调用。

若将来确有成熟 TypeScript 网关需要保留，可通过相同契约替换入口；它不应该迫使文件引擎重写，也不应形成另一套任务和权限状态。当前首版不引入 C++ 重构、额外 JS 层或多套 MCP 服务器。

---

## 6. Codex 接入与工具契约

### 6.1 默认使用 Skill + CLI 桥接真实 MCP

安装只放置程序、隔离运行环境和 Skill；不设置开机启动、Windows Service、托盘、计划任务、全局空载 MCP 或长期网络端口。

Codex 处理实际 PPT 文件或生成任务时，调用 CLI 启动本次 Controller；后者持有真正的 MCP 连接。后续 CLI 调用连接同一任务，任务内复用。仅讨论本规划或解释术语不启动。

OpenAI 官方公开了本地 STDIO MCP 配置，桌面／IDE 配置流程也包含重新启动连接步骤。这不是模型在任意运行中会话里都能热插拔工具的保证；因此不把动态挂接作为首版前提。[S14]

支持原生直接 MCP 接入的客户端以后可增加入口，但须通过同样的启动、结束、再启动和无残留测试；不能为了出现在工具列表而长期保留空进程。

### 6.2 建议保持 11 个聚合业务工具

| 工具 | 拟议职责 |
|---|---|
| `ppt_open` | 明确文件／原生副本／已授权附着来源，返回模式、身份、能力和保存政策 |
| `ppt_inspect` | 概览、查找、对象详情、格式来源、运行状态；按需返回缩略图 |
| `ppt_apply` | 一次提交多对象／多页批次；支持预检和持久回执 |
| `ppt_build` | 调用既有生成、模板与 roundtrip，输出纳入当前任务 |
| `ppt_render` | 按 revision 导出指定页，区分原生渲染和近似预览 |
| `ppt_validate` | 结构、语义保留、原生读取／渲染和覆盖报告 |
| `ppt_commit` | 发布经验证候选，明确保存原稿还是另存，生成 reviewBundle |
| `ppt_close` | 关闭任务文档或解除附着，不默认丢弃未保存修改 |
| `ppt_finish` | Codex 以最新 reviewId 确认工作结束，默认清理退出 |
| `ppt_status` | 已知状态与 operationId 结果，不暗中启动 Office |
| `ppt_diagnose` | 显式需要时检查环境、文件占用、窗口和阻塞状态 |

CLI 另有 start、call、离线 status、finish、abort、resume 等控制职责；这些不是 MCP 标准命令。

### 6.3 契约约束

Python 核心维护唯一的业务 Schema 来源，导出契约供 .NET 请求映射与验证；原生 Host 仍独立检查目标身份、会话代际、允许命令和当前权限，以防边界绕过。两处校验不是两套不一致业务规则。

写请求拒绝未知字段、隐式类型转换、过期目标和歧义。结果保留实际后端、revision、变更范围、失败阶段、持久化状态与未覆盖项。CLI 与 MCP 使用同一语义，不各自实现“更方便”的绕过路径。

当前官方规范页面指向 2026-07-28；实施时锁定已发布且相互兼容的 SDK／协议组合，不等同于使用最新 main。业务 taskId、revision、finish 与标准协议状态分离，不自行猜测握手或关闭消息。[S15]

---

## 7. 引擎选择：按任务约束和已有会话，不按语言信仰

### 7.1 建议决策顺序

1. 检查用户限制：不得启动 Office、必须在当前窗口修改、保留格式、原稿保存政策等。
2. 检查已绑定的文档会话。当前会话已经是 `native-copy` 时，后续支持的操作继续走原生，不因某个文本替换“可能文件更快”而切换。
3. 检查能力矩阵和恢复语义。包括 nativeWithoutActivation、preservesTextRuns、supportsCheckpoint、requiresWorkbookActivation 等。
4. 没有活动会话、只是后台文件编辑时选择 `file`；明确需要 PowerPoint 原生行为时选择 `native-copy`；只有明确授权才附着原稿。
5. 在授权和能力均满足后，才参考本机基准决定可等价操作的优先级。

不写“99% 都走 OOXML”或“COM 一定更快”。选择结果及原因应可观察，例如“当前原生会话已打开，支持该批次，免除后端切换”。

### 7.2 一个批次只能属于一个写后端

预检发现操作不支持时，在执行前拒绝或返回拆分计划。不能先执行 7 项 COM，再静默将第 8 项转到文件层。

**文件 → 原生：** 停止新写入 → 组装有效候选 → 校验并保存 → 以新的原生文稿打开 → 验证内容映射 → 更新 generation 和引用。

**原生 → 文件：** 保存当前内容到受控检查点 → 确认原生不再写该候选 → 关闭对应工作文稿 → 文件层重读实际输出 → 重建索引和 revision。

用户附着文稿通常转成新的文件副本继续工作，不擅自关闭用户文稿。后端切换后，原先的字节保真等级也需要重新说明。

### 7.3 索引是投影，不是唯一真相

`file` 模式的权威数据是源包与修改层；`native` 模式的权威数据是当前绑定的 PowerPoint 内存文稿，最近持久检查点另列。不能用旧磁盘哈希证明原生内存未发生变化。

---

## 8. COM Batch Executor：怎样真正实现快速原生编辑

### 8.1 从“逐框问模型”改成“模型提交一批，本地连续执行”

推荐链路：

```text
读取页面结构与目标属性
→ Codex形成一页或多页的编辑计划
→ 一次 ppt_apply
→ 原生 Host 预检并解析目标
→ 在同一 STA 顺序执行所有动作
→ 局部读回 + 检查点
→ 返回一份变更与错误回执
```

批次可以涵盖文字范围替换、字号／颜色、几何、对齐／分布和固定结构表格。不要在每个属性之间要求模型重新决策。

**注意：** 这是减少模型、MCP 和 Host 往返，不是让所有 COM 属性写入自动变成一次跨进程调用；也不是自动事务。批次数、COM 调用数和保存次数必须分别统计。[S04–S06]

### 8.2 一批原生操作的实施步骤

| 步骤 | 实现要求 |
|---|---|
| 接收与准入 | 有界队列；限制操作数和响应大小；排队时间计入总 deadline |
| 整体预检 | 校验类型、权限、文档代际、目标唯一性、能力、依赖、检查点策略；dry-run 不修改文稿 |
| 目标解析 | 在 STA 内一次取得本批必要的 Slide／Shape／TextRange；同页聚合，避免反复沿 COM 对象链查询 |
| 编译本地计划 | 尊重操作顺序和依赖；可合并独立、同类且语义相同的属性设置，不能跨结构依赖重排 |
| 连续执行 | 常见同类形状可使用 ShapeRange；普通赋值直接访问对象，不 Select／Activate；混合目标分别处理 |
| 验证 | 读回本次改变的字段和必要依赖，记录实际成功／失败对象；不为每项导出整页图 |
| 持久化 | 按批次形成有效检查点并更新持久回执；没有写入磁盘的状态须标为 volatile |
| 返回 | 返回新 revision、目标映射、实际变更、未完成项、验证覆盖与后续动作 |

微软提供了 ShapeRange 操作和 TextRange.Replace 等方法，可作为直接对象操作基础；但字体 runs、继承、分组与复杂对象的保留效果必须按样本验证，不由 API 存在自动推导。[S04、S05]

### 8.3 原生线程与引用管理

专用 STA 线程承担 COM 对象创建、读取、修改和释放，维护消息泵；控制、状态和取消入口不得与长 COM 调用共用阻塞路径。异步包装不能让 COM continuation 跑回线程池。

同一 PowerPoint 实例只有一条串行执行权。多个 task 即便文稿不同，也不能在同一实际实例上各建互不协调的执行器。首版可对同用户会话的原生写任务实行全局互斥，冲突返回忙，而不是建长期守护进程。

有限缓存文稿和本批／当前页引用；关闭、复制、结构变化或 Host 重启使相应引用失效。不长期保留全稿巨大 COM 对象图，不每改一框就强制 GC，也不对不明共享 RCW 盲目 FinalRelease。

官方 Office 线程说明支持 STA／串行与忙状态处理的必要性；它不提供本项目性能数字。[S06]

### 8.4 重试、超时与结果未知

对确认属于 Office busy／拒绝受理的错误，在原 deadline 内有限退避重试；不统一重试所有异常。新增、删除、复制、保存等可能已产生副作用的操作，在结果未知时先核查，不能自动重放整个 batch。

取消只保证阻止尚未执行的后续动作，不能保证立即打断挂起的 COM 调用。超时后隔离该 lease，保留持久回执、操作阶段和最后检查点；Host heartbeat 只证明进程活着，不证明 Office 可响应。

`outcome_unknown`、`partial`、`restored_from_checkpoint` 与 `completed` 必须分开。该批已经修改了部分对象但尚未持久化时，也不能报告“完全没改”。

### 8.5 不复制 Excel 的优化习惯

不假设 PowerPoint 存在 Excel 式 ScreenUpdating／Calculation 开关。不为减少弹窗而写全局注册表、禁用用户加载项、关闭安全策略、自动启用宏，或者把窗口挪到屏幕外冒充后台支持。

公开项目中的可见化、禁 Designer 或任意脚本执行是其自身取舍，不是本项目的隐含前提。[S07、S08]

---

## 9. COM 的两种写入恢复语义

### 9.1 任务副本：可恢复检查点，不称为数据库事务

初次打开前建立原稿只读快照和任务副本；每个已确认批次都有最后有效检查点。执行前确认恢复基线存在；成功后读回并保存新检查点，更新 journal 后才返回 durable success。

失败时，若文稿和应用仍响应，可关闭受影响任务文稿并从最后有效检查点重开；失效所有旧原生引用。恢复成功返回明确状态；无法恢复时保留检查点并终止该原生会话。

PowerPoint 内存中可能短暂出现部分修改，副本恢复并不是对屏幕、事件或外部链接副作用的通用撤销。因此首版排除自动外链刷新和其他不可控副作用。

`SaveCopyAs` 是创建副本的接口；具体 checkpoint 保持 active 文稿绑定、dirty 状态与格式的行为在 P0 验证。**`Saved=true` 不是保存，它会抑制提示并可能导致未保存更改在关闭时丢失。** 禁止用它替代交付保存；只允许在已验证可丢弃的任务工作副本上执行明确 discard。[S16、S17]

### 9.2 用户原稿：best-effort guarded，不保证整批回滚

进入 `native-attached` 前明确目标路径／文稿身份、是否存在未保存修改、是否允许保存原稿、是否只允许另存结果。初次读取若来自内存，应标注 sourceState=native-memory。

可用 SaveCopyAs 获取包含当前内存内容的恢复副本，但它不是用户编辑会话的完整事务快照。任何自动回滚都可能抹掉同期人工改动，所以批次失败默认停止、报告实际范围，并提供恢复副本，不擅自全局 Undo、重载原稿或丢弃内容。

使用文稿／目标指纹、原生事件和局部读回检测外部修改；这些不是与 PowerPoint 人工操作共享的一把事务锁。要求执行批次期间不要同时手工编辑同一份文稿，其他应用仍可继续使用。无法接受这一边界时改用任务副本。

任务结束释放本项目引用，默认不关闭、不 Quit 用户的 PowerPoint；用户原稿尚未保存时，报告真实状态，不因自动化进程已结束就称内容已安全交付。

---

## 10. OOXML 引擎：独立可用，也能与原生编辑串联

### 10.1 保留现有资产

继续使用现有生成、SVG／DrawingML、roundtrip、source-ref／哈希和交付校验。局部微调新增文件补丁路径，不强制全量 SVG 往返；新页面与大幅重设计仍可通过原有作者工作流完成。[S01]

### 10.2 存储与原子工作状态

采用不可变源包、懒加载索引、已修改部件 overlay、journal 和有效 revision 清单。每批在候选修改层执行并校验，成功后发布新清单；失败不替换上一个有效状态。

最终发布或需要完整候选用于原生处理时再组包；不要每改一个字段都重新解压／压缩整稿，也不在原 ZIP 内进行高风险就地写入。每批的磁盘持久化成本与最终组包成本分别测量。

### 10.3 必须保护的细节

| 类别 | 约束 |
|---|---|
| 文本 | 保留段落、runs、域、项目符号和超链接；跨格式替换有明确策略 |
| 图片 | 默认只改指定实例；共享媒体写时复制；纵横比和裁剪明确 |
| 几何 | 外部明确 pt 等单位，内部统一；分组坐标／旋转单独支持，不用屏幕像素冒充 |
| 样式 | 显式值、母版／版式继承、主题引用分开；缺省不是黑色或“没有字体” |
| 关系 | 复制、删除和重排维护关联部件、动画目标、连接器与内部跳转 |
| 未知扩展 | 保留未知 XML 与依赖闭包，注意 namespace 和 mc:Ignorable 等语义；不激进清理资源 |

未修改部件可以按解压后字节验证；修改部件按批准的语义差异验证。整个 ZIP 文件字节相同不是保真指标。经过 PowerPoint 原生保存后，则不能继续宣称原包部件字节完全不变。[S09、S10]

### 10.4 图表单独列能力包

简单内嵌图表数据同时维护工作簿、类别／系列、范围、缓存和关系；复杂公式、外链和特殊图表不默默降级。

原生图表格式与数据编辑分开。微软明确 ChartData.Activate 会激活关联工作簿窗口，因此这条常见数据路径不能纳入默认无激活能力。需要文件侧更新数据时按第 7 节完成安全切换，不能同时改一个已在内存中的嵌入工作簿。[S18]

---

## 11. 对象身份、设计语义与可验证修改

### 11.1 身份模型

`targetRef` 内部绑定 taskId、generation、documentId、后端及 revision、SlideID／part URI、ShapeID、分组路径和指纹。页码与对象名称是显示和搜索线索，不单独授权写入。

原生重开／复制／重建时更新 generation；文件／原生映射必须实际核对。不在查询时自动注入永久 tag 改写原稿。第一版先安全拒绝旧 revision，后续才引入经过验证的细粒度引用复用。

### 11.2 快照同时服务文字模型与视觉模型

默认返回概览、目标文本、几何、类型、显式／继承格式、可执行操作、限制和必要来源。图片另取；完整对象树按需分页，不每次返回所有媒体。

标注“标题”“结论区”“装饰元素”等语义角色时区分来源：placeholder 类型可直接读取，模型推断的角色只用于建议，不自动等于可执行目标。

### 11.3 让编辑工具支持真正的排版工作

高层命令应能表达对齐、等距分布、保持比例缩放、统一局部样式、替换指定范围、按约束移动；不是只开放 left／top 的机械 setter。

每次 layout 操作写清相对谁对齐、哪个坐标系、保留什么尺寸和哪些对象不动。重叠检查区分有意叠放与风险；容纳文字失败时先报告，不擅自改临床数字、结论或大幅缩小全稿字号。

保留原 PPT Master 的内容策略、设计和模板能力，避免新项目退化成“只能挪文本框”的工具。

---

## 12. 验证、交付与原生渲染

| 验证层 | 实施内容 | 证据边界 |
|---|---|---|
| L1 包与关系 | ZIP、XML、部件、内容类型、引用闭包、页面列表 | 不证明内容或显示正确 |
| L2 变更与保留 | 请求对象是否正确改变，未授权对象是否保持，格式和关系差异 | 原生保存允许必要结构规范化，但必须说明 |
| L3 原生读回／渲染 | 属性读回、指定页 Slide.Export、Office／字体／分辨率环境 | 导出成功不等于图像已审阅 |
| L4 视觉与语义验收 | Codex或用户实际核查排版、内容、遗漏和任务要求 | 明确查看了哪些页和哪些图 |

渲染绑定精确 revision；只渲染脏页及其依赖影响页。主题、母版或共享资源变化需要扩大范围。纯文件路径组包后打开只读候选进行原生渲染；原生编辑路径从绑定的内存文稿导出，并关联实际保存检查点。[S19]

对源文稿预先记录错误与警告，不把所有原有警告当成本次修改缺陷，也不顺手修复无关问题。自动样式／容量检查不能冒充视觉通过。

发布默认到新文件；目标同卷准备、完成检查后使用可用的原子发布方式。失败时保留候选，不能“先删原稿再复制”。同步盘／UNC／网络盘能力单列，首版以本地工作区为参考。

reviewBundle 包含交付文件及哈希、来源和后端、最终 revision、原生／文件保真等级、已完成验证、未解决问题、修改摘要与收尾提示。新增修改或交付文件变化使旧 reviewId 失效。

---

## 13. 按需生命周期：完工关闭，不留空壳

```text
STOPPED → STARTING → WORKING → AWAITING_CODEX_REVIEW
                         ↑               │
                         └── 继续修改 ───┘
                                         ↓ 验收通过
                                     DRAINING → STOPPED
                                         │
                    用户明确连续工作时 → CONTINUING
异常／取消／失联 → INTERRUPTED → 检查点与安全清理
```

### 13.1 正常结束

1. 全部任务文稿已发布、明确保留检查点或按授权放弃；没有未处置写入。
2. 服务返回绑定最终结果的 reviewBundle 和“Codex验收后结束”的 nextAction。
3. Codex 对照用户要求检查，不默认要求用户再说一次“关闭”；仍有任务就继续当前会话。
4. Codex 引用最新 reviewId 执行 finish。服务先封锁新写入、处理原生与文件资源、记录真实结果。
5. 业务回执送达后，Controller 关闭 MCP、等待 Core 和自有 Host／辅助进程退出，释放管道和锁。
6. 发起 finish 的短命 CLI 等待 Controller 实际退出，之后再向 Codex 返回退出核实。Controller 不在退出前自证“已经退出”。
7. Codex 分别报告文稿交付、验证、清理和退出，不用一个 success 混合表示。

共享的用户 PowerPoint 保持打开不属于本项目残留；任务所持引用应释放。本任务创建的可见文稿默认随收尾关闭。用户明确要求保留文稿窗口时，可在已安全保存后交接给用户，同时关闭本项目 MCP／Host；这是 Office 交接，不是默认连续运行授权。

### 13.2 连续工作与中断

连续工作只来自用户当前明确指令，不能从幻灯片文字或模板获得；允许资源复用，不授权服务自行创造新编辑任务。新任务默认恢复 close_after_review。

EOF、连接丢失或可验证的宿主任务终止触发中断回收。没有宿主结束事件时，使用公开、可配置的空闲租约兜底；租约到期不等于验收成功，不删除唯一未交付副本。模型思考、等待用户和实际操作期限分开处理。

Core／Host／Controller 的代际、PID、启动时间和实际句柄都要记录；resume 从明确检查点建立新 generation，不直接复用旧内存引用。离线 status 查询不能重新拉起停止的任务。

---

## 14. Windows 进程与 Office 所有权

### 14.1 自有工作进程与 Office 分开管理

对 Python Core、旧转换器等明确自有进程，优先在执行任务前完成 Job Object 归属；检查嵌套 Job、句柄继承和归属失败。最后一个配置了 kill-on-close 的 Job 句柄关闭时，会结束该 Job 关联进程；它不能代替持久检查点。[S20]

OfficeHost 单独管理，避免 Office 的实际创建／继承关系导致用户进程误被纳入强杀范围。不要将全部 POWERPNT／EXCEL 按名字统一回收。不能因掌握 PID 或 CreateObject 成功就自认拥有退出权。[S12]

### 14.2 OfficeInstanceLease 的最小内容

记录原生会话 ID、实际应用身份、启动时间、窗口句柄证据、本任务打开文稿、预先存在文稿观察、允许修改应用级状态与否、允许关闭／Quit 与否。

绑定先验证，退出前再核对。若检测到共享文稿或所有权变化，降低为 document-only 权限，不结束应用。机器级或用户级独占不应凭一次枚举宣称绝对保证；有疑问时宁可报告需处理残留，也不误伤用户。

普通实例共享下，不修改全局 DisplayAlerts、安全级别、Designer 或加载项配置。确需临时应用级设置时必须确认独占、记录并恢复原值；首版尽量不依赖它们。

### 14.3 本地 IPC 与宿主兼容

Named Pipe 使用当前用户／登录会话 ACL、随机 task 身份和不记入日志的会话凭据，限制连接、帧长和操作范围；不能因为是本地 Pipe 就认为天然安全。[S21]

直接启动锁定的本地 EXE／解释器，避免多层 shell、在线安装和 `@latest`。具体 SDK 启动适配按锁定发布版审查，不把一个 main 文件的行为推广到所有版本。

Codex 沙箱是否允许本任务跨工具调用保留受控进程必须 P0 实测；需要权限时使用明确授权，不绕过用户安全边界。该环境确实不支持任务运行器时，返回兼容性限制，不偷偷部署服务替代。

---

## 15. 性能验证：先证明瓶颈，别再凭语言判断

### 15.1 对照设计

固定同一输入副本、同一对象集合、同一修改清单及同一验证／保存等级。初始可用 30 页、100 个文本目标作为可复现实验样例；这是测试规模，不是速度承诺。

| 组别 | 路径 | 目的 |
|---|---|---|
| A | OOXML 单批次，检查点与一次最终发布 | 文件路径基线 |
| B | 同一 PowerPoint 会话，COM 单批次，同等持久化／验证 | 原生批处理基线 |
| C | 同一原生会话，逐工具请求，无额外人为等待 | 分离协议与对象解析重复成本 |
| D | 每次操作重新启动／打开／保存／关闭，仅作小规模反例 | 衡量冷启动和重复工作 |
| E | 真实 Codex 规划／编辑／修订全链路 | 衡量模型调用、返回量及人工体验 |

C 不应人为插入模型延迟以保证它慢；E 的模型时间与工具执行时间分开报告。不能用热 COM 比冷文件，或只为其中一组省略保存／检查点。

### 15.2 记录指标

记录冷启动、源快照、解析／索引、COM 激活与 Open、目标解析、纯 apply、局部读回、checkpoint、ZIP 组装、保存、渲染、退出和总耗时。

同时记录实际模型调用数、MCP／IPC 往返数、互操作调用估计／计数、成功／部分失败率、输出字节、CPU、峰值私有内存、进程残留、前台事件与恢复成本。无法准确统计某类 COM 调用时标为部分观测，不伪造精确计数。

### 15.3 发布指标

先冻结行为门槛：目标正确、原稿安全、支持对象保真、失败可定位、明确退出、无工具导致的前台干扰。速度门槛由 P0 基线形成，并与固定样例、硬件、Office build、字体、冷／热状态绑定。

热查询可重复 100 次；端到端及冷启动单列样本数，报告 p50／p95和范围。正常启停初始做 100 次循环，异常单独注入；这些是测试方案，不是已有成绩。

只有 profile 发现真实 Python 层热点，才优化索引、底层库或考虑 C++ 小模块；不能以 C++ 通常适于计算为理由重写整个项目。

---

## 16. 分阶段实施与验收

### P0：先做 COM 编辑实验与按需生命周期实验

**实施：** 固定仓库和依赖；选真实样例；完成 A／B／C 基准原型；测试 Native Open、混合格式文字、几何、SaveCopyAs／Save／重开、隐藏／遮挡／可见非前台、Office 已开其他文稿、模态窗口及退出。并验证 CLI→MCP→finish→实际退出→再次启动。

**交付物：** `capability-matrix`、可重复基准脚本、环境记录、COM 故障与焦点事件报告、入口兼容结论。

**门槛：** 不以“COM 能连接”或“能创建文件”作为成功；至少证实直接对象编辑、持久化、重开和无误关。原生后台子集不能成立时禁用该子集，文件路径仍可继续，不转为键鼠补救。

### P1：任务控制与统一查询

**实施：** Controller、Python MCP、Skill、单一 Schema；task／document／generation／revision；file 与 native 快照；受限原生文稿绑定；只读 UI 诊断；finish 和离线状态。

**门槛：** 非 PPT 任务零自有进程；重复请求不重复启动；定向查询没有隐式 Office 启动；正常结束真实退出，旧引用拒绝。

### P2：Windows 双引擎编辑 MVP

**实施：** 两引擎共同支持选定的文本、字号／颜色、简单几何和固定表格／备注子集；原生副本 batch executor；文件 overlay；检查点、operationId、局部读回、另存、reviewBundle。

**门槛：** 原生模式必须实际改正在打开的工作副本，而非仅导出图；一次批次多对象执行；对共享支持集进行语义对照；故障不会污染原稿；原生部分失败与文件回滚分别报告。

**范围：** 用户原稿的任意实时附着尚不默认开放，但工作副本可以呈现真实 PowerPoint 实时变化。不得将“逐框实时编辑”再全部推回纯文件路线。

### P3：整合原有生成、设计和验证

**实施：** 现有 Generate／Template／roundtrip 接入统一任务；生成后能在原生副本中继续精修；渲染—审阅—修订闭环；检查遗留确认网页、素材下载和转换器启停。

**门槛：** 原有代表样例不退化；没有默认弹网页；跨引擎切换只在安全边界；最终交付包含内容与验证覆盖。

### P4：受控附着用户文稿

**实施：** 指定已打开文稿、含未保存内容的快照、独立保存政策、指纹与外部改动检测、有限批次、best-effort 部分结果和解除附着。

**门槛：** 同名窗口不改错；不把旧磁盘文件当 live 状态；人工冲突停止；失败不全局 Undo；finish 不关闭用户 PowerPoint。该阶段按实际需求发布，但从 P1 就设计相关身份与恢复差异。

### P5：高级能力与正式发布

**实施：** 图片裁剪替换、分组与对齐、页面复制／重排、主题／版式、更多表格、图表、动画和媒体按能力包推进；兼容矩阵、预构建安装、版本与契约校验。

**门槛：** 每一能力包都附带关系、保留、无前台和失败测试；不因支持面扩展而恢复宏／任意脚本／全局设置改动。

**完整首版建议：P1 + P2 + P3。** 安全与生命周期从 P0 随功能建设，不留到发布前补做；不需要先完成 P4／P5 才能获得原生快速编辑体验。

---

## 17. 仓库落点与工作包

### 17.1 复用和改造位置

| 位置 | 行动 |
|---|---|
| `skills/ppt-master/` | 保留设计、模板、生成和原生恢复资产；新入口按需加载其参考 |
| `scripts/pptx_to_svg.py`、`svg_to_pptx.py` 及实现包 | 保留 CLI；提供结构化库适配，避免解析自然语言日志 |
| `scripts/authoring_roundtrip.py` | 复用 source-ref／哈希／恢复；新增对象身份不可与之冲突 |
| `scripts/pptx_opc_validation.py`、`pptx_delivery_check.py` | 复用关系与交付检查，补充双引擎输出基线 |
| `scripts/powerpoint_video.py` | 旧入口过渡保留；新任务统一走 Host，不默认 GetActiveObject 后决定全局退出 |
| `AGENTS.md`、Skill 路由和贡献说明 | 明确新 fork 的任务控制与授权边界，避免旧确认网页／常驻假设冲突；保留许可与归属 |
| WinCode 相关代码 | 借鉴或少量提取 ToolRegistry、OperationContext、原生协议／只读诊断思想；不形成全仓运行依赖 |

### 17.2 建议目录

```text
src/ppt_mcp/
  server/          MCP 契约与入口
  domain/          task、document、revision、target、receipt
  planning/        能力路由、批次预检与依赖
  file_engine/     索引、overlay、局部补丁、关系
  native_adapter/ 原生请求映射、generation 与结果对接
  validation/      结构、保留、原生证据与报告
  adapters/        旧生成/roundtrip
windows/
  TaskController/ CLI、受保护 IPC、MCP client、归属与退出
  OfficeHost/      STA、batch executor、checkpoint、render、lease
contracts/          单一来源导出的 schema 与样例
tests/
  contract/ lifecycle/ file/ native/ preservation/
  recovery/ foreground/ fixtures/ benchmarks/
skills/ppt-mcp/     操作路由、任务开始、验收与收尾
docs/decisions/     架构决策与兼容依据
```

这些是职责边界，不要求建立几十个空服务或引入微服务、消息总线与数据库。

### 17.3 建议 PR／工作包

| 工作包 | 交付 | 同步测试 |
|---|---|---|
| WP1 | 决策记录、真实 fixtures、基准原型 | 对象数／输入哈希／Office 环境可复现 |
| WP2 | Controller 与 task 生命周期 | 根进程先退出、再次启动、状态不自启、并发 start |
| WP3 | Python MCP 与统一契约 | 字段错误、大小限制、编码、旧 schema、协议污染 |
| WP4 | 两引擎定向查询与身份 | 页序、同名对象、分组、已保存／未保存来源 |
| WP5 | 原生 batch executor 与检查点 | 多对象、拒绝调用、部分失败、超时、恢复引用失效 |
| WP6 | 文件局部补丁与保留 | 多 run、未知扩展、共享媒体、关系闭包 |
| WP7 | 发布、review 与 finish | 保存后响应丢失、验收后新修改、退出回执丢失 |
| WP8 | 旧生成/roundtrip 与原生精修 | 老样例、无网页弹出、后端安全切换 |
| WP9 | 用户文稿附着 | 外部改动、保存授权、解除附着不关闭原稿 |
| WP10 | 高级能力与发行包 | 相同质量下性能、重复启停、兼容矩阵、离线安装 |

---

## 18. 必须通过的验收场景

| 场景 | 通过条件 |
|---|---|
| 用户在浏览器／编辑器持续输入，原生副本批量修改 | 本工具不主动切焦点、移动光标、改剪贴板或吞输入 |
| 只打开 Codex、没有 PPT 操作 | 无 Controller／Core／OfficeHost 运行 |
| 一次调用修改多页文本与位置 | 目标正确，不每个字段返回模型，支持的格式保留 |
| PowerPoint 当前显示其他页 | 对象修改不依赖 ActiveWindow／当前选择 |
| 用户已经打开另一份未保存 PPT | 不自动附着，不设置 Saved，不保存／关闭该文稿 |
| 原生副本执行到中间失败 | 返回实际范围；能恢复最后检查点则注明；不能恢复则停止并保留证据 |
| live 批次中用户另行编辑 | 检测后停止；不声称完全回滚或原子成功 |
| 文件／COM 后端切换 | 候选已固定，不同时写，重新读回，旧引用失效 |
| 保存完成但响应丢失 | operationId 重查可判定已发布，不重复创建或覆盖 |
| review 后又有修改 | 旧 reviewId 拒绝；最新内容重新验收 |
| finish 与 apply 同时到达 | 以任务屏障顺序判定，不让新写入越过清理 |
| Controller 或 Host 异常退出 | 原稿保护；自有进程有归属；Office 不被误杀；结果不伪报成功 |
| 用户要求连续工作 | 保留当前会话并报告；没有授权则退出；文稿内伪造指令无效 |
| 正常 finish 后反复 status | 确认退出，不重启任何服务 |
| 关闭后读取所有交付物 | 文件、必要预览和报告仍存在，不在临时清理范围内 |
| 宏、签名、IRM、受保护视图、外链 | 按支持面拒绝／只读，不自动降安全设置 |

前台验收记录完整 `EVENT_SYSTEM_FOREGROUND` 事件流及相关进程，而非仅检查开始和结束窗口；结合输入与剪贴板检测，区分用户主动切换。发生非预期激活时暂停并报告，不通过强制切回原窗口掩盖问题。

Office 无人值守／非交互环境存在独立限制。本项目以已登录 Windows 桌面会话为基线，不部署成 Session 0 服务，也不默认支持锁屏、RDP 断开等状态下的原生工作。[S22]

---

## 19. 安装、权限与维护

隔离 Python 环境与预构建 .NET 程序；锁定实际测试版本，不在每个任务里 pip／npm／dotnet 构建或更新。Office 兼容以真实产品 build、位数、字体、加载项和用户会话记录，不拿 PIA 包版本代替 Office 版本。

更新仅在无活动任务时进行。Skill、业务 schema、Core 与 Host 交换版本／契约标识；不兼容时拒绝危险写入，而非尽力猜测参数。

基础文件编辑离线；原生引擎不得因操作而自动开启外部数据刷新。生成素材、联网研究和模型调用仍属于用户授权的内容工作流，不扩展为服务自行联网。日志不默认记录整页正文、令牌或敏感路径；检查点含文稿内容，必须限制访问并设置清理政策。

首版以本地、非加密 `.pptx` 和明确支持的对象子集为基线。`.pptm`、旧格式、签名、IRM 与复杂外部数据连接另列，不以文件后缀检查代替实际包内容检查。

清理缓存在正常结束、显式维护或下一次有界启动清理时执行；不为了清理磁盘再建设后台守护任务。未知对象或唯一未交付检查点不得自动删除。

---

## 20. 给实施 Codex 的任务说明

> 在 linnnn89/ppt-mcp 中实施 Windows 任务级双引擎 PPT 编辑器。保持现有 Python 生成、DrawingML、roundtrip、对象源恢复和校验资产，不整仓重写，不把 WinCode 变成运行依赖。技术栈沿用 Python MCP／编排核心与 C#/.NET TaskController、OfficeHost；不为性能猜测新增 TS 或 C++ 必需层。
>
> 按 PPT 任务启动，任务内复用，交付后返回验收包并提示 Codex 完工后 finish，只有用户明确连续工作才保留。默认入口为 Skill＋CLI 桥接真实 MCP，不依赖动态热挂接，不建立开机服务、常驻网关或全局后台监听。退出必须由外部调用者观察；status 不得自启。
>
> 首先实施 P0：公平比较文件批量修改与同一 PowerPoint 会话的 COM 批量修改，并验证真实保存、重开、前台事件与退出。COM 是正式写引擎，不只负责渲染。P2 必须交付在已打开任务副本上的真实批量编辑；不要把该能力再无限后置。UIA 仅用于状态诊断，不用键鼠或前台点击补救。
>
> 核心用统一对象命令、targetRef、revision、generation、operationId 和可查询回执。每批只允许一个写后端；安全切换经过检查点、关闭和重新读取。文件路径以候选修改层保护原稿；原生副本以持久检查点恢复；用户已打开原稿采用明确授权和 best-effort 语义，不承诺通用原子回滚。
>
> OfficeHost 使用专用 STA、有界串行队列、消息泵、有限分类重试及按批次读回。缓存必要引用，不每个字段经过模型／Python跨进程请求。不能把一次 MCP batch 说成一次 COM 调用或一个原子事务。禁止任意脚本/evaluate、自动全局 Designer／安全设置修改、Saved=true 冒充保存和按进程名结束 Office。
>
> 每阶段同时交付正确性、保留、冲突、故障、生命周期与无前台干扰测试。只有实际绑定的文稿可编辑；知道 PID 不等于拥有 Quit 权限。最终报告分开记录交付、验证、清理与退出，不用成功导出或关闭回执代替真实验收。

---

## 21. 最终结论与待验证项

本版不再在“只用 COM”与“尽量排除 COM”之间摇摆。**文件路径负责无需应用的精确修改，原生路径负责已绑定 PowerPoint 会话的批量操作与真实排版；选择依据是任务约束、支持能力、会话状态和实测，而不是语言或演示印象。**

尚待实机验证：目标 Codex 允许的任务运行器方式、参考 Office 的隐藏／非激活行为、实例复用和所有权、具体格式保留、检查点恢复、COM 批量与文件路径的端到端性能。这些已安排为 P0 与发布门槛，不是未作决定的架构分叉。

已完成的是新版规划、相关代码／文档与一手报告核查；未修改仓库、未建立分支、未执行 Windows 性能或 COM 测试。

---

## 附录：核查来源与使用边界

下列条目是事实来源；模块划分、步骤、状态名、测试规模和优先级为本规划提出的设计。源码片段不等于完成全仓审计，README 演示与作者测试不等于本项目验证。

**S01｜本项目与已有引擎。** 本次重新核对 main，沿用对相同固定提交下入口、roundtrip 和验证相关文件的既有读取；改造前再固定依赖与测试基线。

历史项目参考已去标识化；不作为当前实现证据。

相关路径：`skills/ppt-master/scripts/authoring_roundtrip.py`、`pptx_opc_validation.py`、`pptx_delivery_check.py`、`powerpoint_video.py`、`workflows/edit-native-pptx.md`。

**S02｜WinCode。** 本次确认 main 并重读 ToolRegistry；参考前轮相同提交的 README、架构与原生适配读取。

历史工具参考已去标识化；不作为当前实现证据。

**S03｜OpenAI 官方 Mac 后台 computer use 说明。** 2026-04-16；只证明公开产品能力，不证明特定 PPT 视频的底层实现。

`https://openai.com/index/codex-for-almost-everything/`

**S04｜Microsoft Shapes.Range。** 同页形状集合上的范围操作。

`https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shapes.range`

**S05｜Microsoft TextRange.Replace。** 范围内查找替换接口；不由文档推导任意复杂格式都无损。

`https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textrange.replace`

**S06｜Microsoft Office threading support。** STA、线程安全、忙状态与调用拒绝。

`https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office`

**S07｜sbroenne/mcp-server-powerpoint。** 本次读取 `PresentationBatch.cs` 1–400 行，返回 blob `74e74920f726ce0480c9e255ea557fdb06d82a3d`：专用 STA、队列、启动、可见性、Designer 注册表修改、进程身份与重开逻辑。跟随 main 的 URL 会变化，实施时应记录实际提交。

`https://github.com/sbroenne/mcp-server-powerpoint/blob/main/src/PowerPointMcp.ComInterop/Session/PresentationBatch.cs`

同时读取作者延续记录中的参数／保存／真实测试经历；其中含早期和后期更新，不能将所有历史段落当作同一时点事实。

`https://github.com/sbroenne/mcp-server-powerpoint/blob/main/CONTINUATION.md`

**S08｜Ayushmaniar/powerpoint-mcp。** 阅读 README 的 COM、实时操作、快照和 evaluate 功能说明；未进行其源码全审或性能复现。不采用 README 关于其他库“只写”等笼统说法。

`https://github.com/Ayushmaniar/powerpoint-mcp`

**S09｜pptx-automizer。** 前轮已读取关系助手的具体源码，作为部件／关系保留经验而非替换引擎依据。

`https://github.com/singerla/pptx-automizer/blob/main/src/helper/xml-relationship-helper.ts`

**S10｜python-pptx 官方文本 API。** 本次重读文本层级和赋值语义。

`https://python-pptx.readthedocs.io/en/latest/api/text.html`

**S11｜Codex issue #34614 与评论。** 创建于 2026-07-21，查询时更新于 2026-08-26；特定 Windows 客户端的进程链报告。评论 #5083356069 为用户提供的限定源码实验，不是本项目复现，也不等于官方已合入修复。

`https://github.com/openai/codex/issues/34614`

`https://github.com/openai/codex/issues/34614#issuecomment-5083356069`

**S12｜Microsoft GetObject／CreateObject 行为。** 包含历史 PowerPoint MultiUse 表，现代构建需实测；不拿旧表保证当前实例隔离。

`https://learn.microsoft.com/en-us/previous-versions/troubleshoot/microsoft-365/microsoft-365-apps/office-suite-problems/getobject-createobject-behavior`

**S13｜Microsoft Presentations.Open。** WithWindow、ReadOnly 和 Untitled 参数。

`https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open`

**S14｜OpenAI 官方 MCP 配置文档。** 本次跳转到 ChatGPT Learn；支持本地 STDIO 配置，配置能力不自动等于任意会话热挂接。

`https://developers.openai.com/codex/mcp/`

**S15｜MCP 官方规范。** 本次 latest 跳转到 2026-07-28；版本支持须另按发布 SDK 与客户端测试。

`https://modelcontextprotocol.io/specification/2026-07-28`

**S16｜Microsoft Presentation.SaveCopyAs。** 创建副本接口；本规划的检查点协议是工程设计，不是该 API 自带事务。

`https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.savecopyas`

**S17｜Microsoft Presentation.Saved。** 显式说明设为 true 会抑制提示，关闭时可能丢失未保存变化。

`https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saved`

**S18｜Microsoft ChartData.Activate。** 激活关联工作簿窗口，以及 Workbook 属性使用前提。

`https://learn.microsoft.com/en-us/office/vba/api/powerpoint.chartdata.activate`

**S19｜Microsoft Slide.Export。** 原生页面导出接口。

`https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.export`

**S20｜Microsoft Job Objects。** 关联范围、继承、嵌套和 kill-on-close 行为。

`https://learn.microsoft.com/en-us/windows/win32/procthread/job-objects`

**S21｜Microsoft Named Pipe Security。** ACL、登录会话和访问检查。

`https://learn.microsoft.com/en-us/windows/win32/ipc/named-pipe-security-and-access-rights`

**S22｜Microsoft 无人值守 Office 自动化说明。** 非交互／无人值守场景限制；不能直接推广为当前用户会话中任何自动化都不可用。

`https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa`
