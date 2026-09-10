# docs

将需要由 Microsoft Word 原生渲染的 `.docx` 文件放在这里，可以使用子目录。

提交或更新 DOCX 后，GitHub Actions 会自动：

```text
DOCX
→ Windows self-hosted runner
→ Microsoft Word
→ PDF
→ 结构化版式 QA
→ GitHub artifact
```

生成的 PDF 默认与 DOCX 放在同一目录，并在没有并发分支变化时提交回当前分支。
