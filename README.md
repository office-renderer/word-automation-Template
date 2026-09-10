# Word Automation Template

用于 `office-renderer` 组织内 Word 文档自动化处理的模板仓库。

核心流程：

```text
ChatGPT / AI 修改 DOCX
        ↓
DOCX 提交到 GitHub
        ↓
GitHub Actions
        ↓
组织级 Windows self-hosted runner
        ↓
Microsoft Word 原生渲染
        ↓
DOCX → PDF
        ↓
结构化版式 QA
        ↓
PDF / QA 结果上传到 GitHub
        ↓
ChatGPT 获取 PDF 做视觉检查
        ↓
发现问题后继续修改 DOCX
```

## 使用方式

1. 用本仓库创建新的项目仓库，并尽量保留在 `office-renderer` 组织内。
2. 将需要处理的 Word 文档放入 `docs/`，支持子目录。
3. 提交或更新 `.docx` 后，`Render Word to PDF` 会自动运行。
4. GitHub 会从所有满足以下标签的在线 Runner 中自动选择一台空闲机器：
   - `self-hosted`
   - `Windows`
   - `X64`
   - `word`
5. Runner 使用本机 Microsoft Word 原生导出 PDF。
6. 生成的 PDF 会：
   - 作为 workflow artifact 上传，便于 ChatGPT 获取并做视觉 QA；
   - 在分支未发生并发变化时提交回当前 Git 分支。
7. 结构化 QA 报告作为 artifact 上传，不写入正式仓库历史。

## 目录

```text
.github/workflows/
├─ render-word.yml        # 正式：DOCX → Word → PDF → QA → GitHub
└─ runner-health.yml      # 手动：检查 Runner / Git / PowerShell / Word

scripts/
├─ Export-WordToPdf.ps1   # Microsoft Word 原生 PDF 导出
├─ Inspect-WordLayout.ps1 # 单个 DOCX 的结构化版式检查
└─ Run-WordLayoutQa.ps1   # 批量执行 QA

docs/
└─ README.md              # Word 文档放置目录
```

## Runner 选择

工作流按标签选择 Runner，不绑定具体电脑。

因此如果有多台合格电脑同时在线，GitHub 会选择其中一台当前可用的机器；如果只有一台在线，就使用那一台；全部离线时任务会等待。

只有安装环境满足要求的电脑才应配置 `word` 标签。建议这些电脑尽量保持一致的：

- Microsoft Word / Office 版本；
- 常用字体；
- 文档模板；
- PowerShell 7；
- Git for Windows。

## 重要原则

- 正式版式以 Windows 上 Microsoft Word 的渲染结果为准。
- LibreOffice、Pandoc 等不作为正式 Word 版式的权威渲染器。
- 结构化 QA 不能代替 PDF 视觉检查。
- 修改 DOCX 后必须重新渲染，不能继续检查旧 PDF。
- Runner 只执行受信任仓库中的 workflow；不要让不可信 PR 在个人电脑上的 self-hosted runner 执行任意脚本。
