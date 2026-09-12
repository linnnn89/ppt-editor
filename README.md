# ppt-editor

Windows / Node 24 的 PowerPoint 编辑核心，当前版本 **v0.1.1（核心预览版）**。

包含 OOXML 文件编辑、独立 COM 工作进程、原稿保护、检查点及保存/重开验证。CLI、MCP 服务和 Skill 尚未实现；没有全局命令或 npm start 入口。

## 环境与验证

需要 Windows x64、Node 24、已安装 PowerPoint，以及本机已有的 Python / MSVC C++ 构建工具。安装脚本在项目内构建 winax，不安装系统工具。当前本机已验证 Node 24.19.0；winax 使用 Node 24.18.1 headers 编译以规避已复现的原生 GC 问题，不改变系统运行时。

```powershell
npm ci
npm run check
npm test
npm run test:native
npm run test:lifecycle
```

原生验证使用自动生成的虚构内容，临时文件位于 work/。启动 PowerPoint 仍可能短暂进入前台，不承诺完全静默。其他 Node/Office 版本须独立验证。

## 批量编辑与性能实验

必须显式传入自己授权使用的样本；仓库不包含私人样本或默认私人路径。

```powershell
npm run benchmark:fast -- --source "path/to/sample.pptx" --out "outputs/new-run"
npm run benchmark -- --source "path/to/sample.pptx" --out "outputs/comparison"
```

快速入口使用 XML 批量修改并避免媒体重新压缩；完整入口对比逐片段 COM、按文本框 COM、XML 全包压缩、XML 避免媒体重压缩。生成文件后仍执行 PowerPoint 读回验证。

实验对每个非空段落加测试标记，设置字体、字号、文字颜色及幻灯片背景，并包含使用中母版的非占位文本。字号规则是基础测试规则，不保证任意文稿排版；应检查输出中的布局标记和实际预览。图片内部文字不做 OCR，模板提示和动态字段不作为正文修改。

结果只写入本地 outputs/。报告含样本内容、文件路径、指纹或预览，应视为私人运行产物，不加入 Git。源码中的性能实验优化尚未替换通用 FileEngine.bytes() 的写包策略。

详见 [实验方法](docs/benchmarks/methodology.md)。源文件、输出文件和私人运行日志默认排除版本控制。

## 结构与版本

- src/file-engine.js：文件索引、编辑、检查点。
- src/native-host.js 与 src/native-worker.js：原生编辑与生命周期。
- scripts/benchmark-pptx.js：四方法对照及快速入口。
- docs/codex_worklog.md：去标识化技术记录。

采用 MAJOR.MINOR.PATCH：兼容修复递增补丁号，新增经验证能力递增次版本号；完整工作流验收前不标为 1.0。版本说明见 [CHANGELOG](CHANGELOG.md)。
