# Word Automation Template

用于 `office-renderer` 组织内 Word 文档自动化处理的模板仓库。

这个仓库只提供处理方法。**实际 Word 文档始终留在目标项目仓库的原位置，不需要搬到固定的 `docs/` 目录，也不需要搬到本模板仓库。**

核心流程：

```text
目标项目仓库中的任意位置
        │
        ├─ 项目申请书.docx
        │
        ↓
GitHub Actions
        ↓
组织级 Windows self-hosted runner
        ↓
Microsoft Word 原生渲染
        ↓
原位置生成：
项目申请书.pdf
        ↓
结构化版式 QA
        ↓
PDF / QA artifact
        ↓
ChatGPT 获取 PDF 做视觉检查
```

## 使用原则

- Word 文件在哪个目录，就在那个目录直接转换。
- PDF 与对应 DOCX **同目录、同文件名，仅扩展名不同**。
- 不要求建立 `docs/` 文件夹。
- 不把业务文档复制到 `word-automation-Template`。
- 模板仓库只负责提供 workflow、PowerShell 脚本和 QA 逻辑。

例如：

```text
项目仓库/
├─ 申报材料/
│  ├─ 项目申请书.docx
│  └─ 项目申请书.pdf
├─ 汇报材料/
│  ├─ 阶段报告.docx
│  └─ 阶段报告.pdf
└─ 背景资料/
   └─ ...
```

提交或更新仓库中任意 `.docx` 后，`Render Word to PDF` 会自动扫描仓库内受 Git 管理的 Word 文档，使用 Microsoft Word 原生渲染，并把生成的 PDF 推送回各自 DOCX 所在的原目录。

## Runner 选择

工作流使用：

```yaml
runs-on: [self-hosted, Windows, X64, word]
```

因此不绑定具体电脑。GitHub 会从所有具有这些标签且在线、空闲的 Runner 中自动选择一台。

## 模板内容

```text
.github/workflows/
├─ render-word.yml
└─ runner-health.yml

scripts/
├─ Export-WordToPdf.ps1
├─ Inspect-WordLayout.ps1
└─ Run-WordLayoutQa.ps1
```

其中：

- `render-word.yml`：任意位置 DOCX → 原位置 PDF → QA → artifact → PDF 提交回当前分支；
- `runner-health.yml`：手动检查 Runner、Git、PowerShell 和 Word 环境；
- `Export-WordToPdf.ps1`：Word 原生 PDF 导出，输出始终位于源 DOCX 同目录；
- `Inspect-WordLayout.ps1`：结构化版式检查；
- `Run-WordLayoutQa.ps1`：批量扫描仓库 Word 文档并汇总 QA。

## 重要原则

- 正式版式以 Windows 上 Microsoft Word 的渲染结果为准。
- LibreOffice、Pandoc 等不作为正式 Word 版式的权威渲染器。
- 结构化 QA 不能代替 PDF 视觉检查。
- 修改 DOCX 后必须重新渲染，不能继续检查旧 PDF。
- Runner 只执行受信任仓库中的 workflow。
