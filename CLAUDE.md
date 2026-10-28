# Pts Claude Plugins

## 项目简介
Claude Code 插件仓库，团队成员可通过以下命令注册和安装插件。

### 注册市场
```bash
claude plugin marketplace add https://it-gitlab.topscomm.com/xxxt/znzz/ai/pts-claude-plugins.git --scope project
```

### 安装插件
```bash
claude plugin install tops-encoding@pts-claude-plugins --scope project
claude plugin install tops-git@pts-claude-plugins --scope project
claude plugin install tops-docx@pts-claude-plugins --scope project
```

## 插件列表

- **tops-encoding** — 文件编码处理插件，用于处理读写文件过程中的文件编码问题
- **tops-git** — GitLab 代码管理插件，提供 `commit`（规范提交代码）与 `push`（提交并推送代码）两个技能
- **tops-docx** — 文档撰写与转换插件，提供 `prd`、`hld`、`lld`、`um` 四个文档生成技能（输出 markdown）与 `md2docx` 转换技能（按 docx 模板样式生成 Word 文档）

## 开发说明

### 添加新插件
1. 创建 `plugins/{plugin-name}/` 目录
2. 添加 `plugins/{plugin-name}/.claude-plugin/plugin.json` 插件元信息
3. 在 `plugins/{plugin-name}/skills/` 下添加技能目录和 `SKILL.md`
4. 在根目录 `.claude-plugin/marketplace.json` 的 `plugins` 数组中注册插件入口

### 项目结构
```
.
├── .claude-plugin/
│   └── marketplace.json      # 插件市场注册入口（plugins 数组）
├── plugins/
│   └── {plugin-name}/
│       ├── .claude-plugin/
│       │   └── plugin.json      # 插件元信息
│       └── skills/
│           └── {skill-name}/
│               ├── SKILL.md     # 技能定义
│               ├── scripts/     # 可选：技能附带的脚本（如 tops-docx 的 md2docx.py）
│               └── assets/      # 可选：技能附带的模板等资源（如 tops-docx 模板）
├── CLAUDE.md                 # 项目说明（本文件）
└── README.md                 # 项目介绍
```