# 模板目录

把 docx 模板放在本目录，建议命名为 `template.docx`：

```
plugins/docx/skills/md2docx/assets/template.docx
```

模板的职责只有一个：**提供样式**。脚本以模板为底版加载，清空正文后按 markdown 追加内容，另存为新文件，不会修改模板本身。

模板中应包含的样式见上一级 `SKILL.md` 的「准备模板」一节。缺少的样式会被降级为正文样式，并在转换结束的摘要中列出。

正文骨架、封面文字、目录条目都不要写进模板——它们来自 markdown 与 frontmatter。
