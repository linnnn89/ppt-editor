# Changelog

## 0.4.0-rc.4 (release candidate)

- Document the default file-mode batch workflow, page-level refinement, immediate visual exceptions, actual-response ordering and final committed-output review in the Skill and README.
- Keep historical plans, process notes and one-off benchmark scripts in an ignored local archive; retain production code, regression tests, reusable fixtures and supported validation tools in the repository.

- 同稿竞速驱动的 batch 优化：文件模式允许为继承位置的顶层文本占位符写入完整绝对变换，拒绝缺项；正式 render 可复用原生读回完成整稿布局检查，并提供完整审阅记录路径，减少重复 COM 检查和响应正文。默认参数兼容旧流程。
- 保持多页 batch：build/compose/apply 默认返回按页检查摘要；按问题页精修，不强制逐页制作或截图。
- 文件查询支持只解析选定页对象；增加 layout-only 检查、summary 查询及可恢复的任务页面摘要。检查和设计声明绑定页面与布局依赖，保留未受影响页面的结果。
- 文件草稿可按 expectedRevision 直接预览检查点；正式成品的分批预览证据按页累积，原生读回补充普通未旋转文本边界测量。
- finish 区分调用方确认、视觉证据与资源清理；requireAccepted 可检查最终整稿审查和完整页面预览。没有预览不会标记 visualReviewed；所有待确认成品均校验输出哈希。
- 旧操作回执可重放；不增加工具或依赖。升级后需刷新 MCP 连接并核对 batchPageReports 等能力，磁盘版本不代表已连接进程已更新。

## 0.4.0-rc.3（本地候选，未发布）

- ppt_validate/CLI validate 新增按页或整稿的 layoutCheck：报告对象边框重叠、超界量及几何未解析项，不生成图片，不修改内容。
- 已知设计叠放可用当前修订绑定的对象对显式声明；其他交叠返回立即视觉核查提示，修正并复查后继续，无可疑重叠页面直接继续制作。
- 持久保留各页问题清单，后续修改使旧修订检查过期；最终统一检查全稿、逐项修复并集中预览。Skill 不再要求每页完成即截图。
- 普通文本实际溢出、图表/SmartArt 内部、线宽/阴影不在几何检查覆盖内；旋转外接框交叠为候选风险。clear 不作为视觉批准。

## 0.4.0-rc.2（本地候选，未发布）

- 新增可编辑多系列、堆叠及柱线双轴组合图，保存静态内嵌数据工作簿；原生护栏继续拒绝公式、宏、外链和非图表嵌入对象。
- 支持原生文本转 SmartArt、节点文字修改及模板内布局复用。隐藏 Office 的节点写入 E_FAIL 仅在原文未变且关联文字形状唯一时使用替代写入，并核对读回。
- 新增 ppt_relayout/CLI relayout：显式对象选择、默认预演、双列或单列布局、溢出提示；图表等框架不写不支持的旋转属性。
- 增加 executive/editorial/academic 原创主题；修复 MCP 参数默认值提前覆盖主题字体、颜色和背景。
- 任务状态持久记录原稿/副本/进程绑定；修改入口新增 taskId 校验，原生 IPC 绑定私有标识，拒绝重新初始化更换任务身份。同一用户原生租约仍互斥。
- MCP 为 14 项工具。真实验收覆盖 hierarchy SmartArt，process/cycle 仍依赖本机或模板布局可用性；长文字不保证自动适配。没有提交、推送或发布。

## 0.4.0-rc.1（本地候选，未发布）

- A 路线正式生成支持水平/垂直翻转、旋转和文本下划线；自行车实验复用正式生成器。
- 文件和原生引擎增加双色线性渐变、图片、纹理、图案及母版背景继承；保留旧纯色参数。
- 新增 ppt_compose：保留模板现有页，追加旧稿或新页面，复用参考页背景/版式和无文字装饰，可统一文字样式与留白。
- ppt_render 支持指定提交 reviewId，以只读独立副本核验与预览实际成品；持久保存预览和调用方确认记录，拦截输出变化及过期审阅。
- MCP 扩充为 13 项工具；CLI 补齐 compose/render/diagnose 和离线审阅查询；diagnose 报告运行版本与能力。
- 使用三个新增自动化场景覆盖兼容、护栏和真实 STDIO/Office 流程；桌面当前连接刷新验收仍待完成。不承诺任意模板、自动视觉美化或通用性能收益。

## 0.3.0-rc.2（本地候选，未发布）

- 原生关闭失败不再标记 closed 或删除检查点；finish 在清理确认前保留恢复材料。
- CLI、MCP 和 finish 共用清理结果判断；未知结果持久化并阻止后续清理绕过，保留合法外部 Office 会话。
- 文件模式原生读回改为校验最新修订的独立只读候选，返回哈希和修订依据，不改变编辑会话。
- 增加三个可靠性回归场景；Codex 当前任务直接调用与插件安装验收仍待完成。

## 0.3.0-rc.1（本地候选，未发布）

- 修正优化写包对 ZIP 目录项的处理，验证实际输出后防覆盖发布。
- 修复原生预检、恢复状态持久化及同一连接的会话复用。
- MCP 新增 ppt_start 与历史任务状态查询；服务版本与 package 一致。
- 实验脚本去除私人默认值；隐私检查包含暂存区实际内容与未跟踪 Skill，匹配值不回显。
- 文档区分候选测试与正式接入验收，修正 CLI 示例。
- 几何图形建稿、背景编辑、MCP 与 Skill 属于本次尚未发布的候选能力；Codex 实际接入验收待完成。

## 0.1.1

- 移除公开文档中的私人样本信息和运行记录。
- 移除实验脚本默认私人路径、样本正文匹配及样本专属布局参数。
- 所有性能实验要求显式 --source，验证使用虚构样本。
- 扩大私人文稿、截图和运行产物的版本控制排除规则。
- 重建去标识化的公开历史。

## v0.1.1 历史能力范围

OOXML 文件编辑、原生 COM 工作进程、原稿保护、持久检查点及真实退出验证。保留快速 XML 实验与四方法性能对照。CLI/MCP/Skill 尚未实现，通用 FileEngine 写包策略尚未切换到实验优化实现。

版本采用 MAJOR.MINOR.PATCH，记录实际验证与适用边界。
