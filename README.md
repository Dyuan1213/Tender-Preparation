# Tender-Preparation

招标投标工具集（Claude Code 插件 / Marketplace），包含两个 Skill：

- **tender-parser** — 解析招标文件（PDF/DOCX），自动提取关键信息并生成结构化摘要报告。
- **tender-writer** — 基于招标文件的响应格式要求，逐章编制投标文件，并输出宋体、带页码的 Word 文档（支持表格、配图、标题大纲）。

> **不用 Claude Code？** 在 Codex、扣子(Coze)、ChatGPT 等其他平台或本地手动使用，请看 [在非 Claude Code 环境下使用](docs/在非ClaudeCode环境下使用.md)。

## 安装（推荐：作为插件）

在 Claude Code 中执行：

```
/plugin marketplace add Dyuan1213/Tender-Preparation
/plugin install tender-tools@tender-preparation
/reload-plugins
```

安装后两个 skill 即可用。说"帮我解析招标文件""帮我编标书"即可触发；插件内 skill 也可按 `/tender-tools:tender-writer` 这样的命名调用。

## 更新

仓库每次推送即为新版本（plugin.json 未固定 version，按 commit 计版本）。使用者更新：

```
/plugin marketplace update tender-preparation
/reload-plugins
```

或在 `/plugin → Marketplaces` 中对本 marketplace 开启 **Enable auto-update**，启动时自动更新。

## 依赖

- Python 3，依赖 `python-docx`：`pip install python-docx`
- 生成的 Word 使用**宋体**，查看端需安装中文字体（Windows 自带）。
- tender-parser 如需上传飞书，参见 `skills/tender-parser/.env.example`，自行配置飞书应用凭据（`.env` 不会被提交）。

## 目录结构

```
Tender-Preparation/
├── .claude-plugin/marketplace.json   # marketplace 目录清单
└── plugins/
    └── tender-tools/
        ├── .claude-plugin/plugin.json
        └── skills/
            ├── tender-parser/
            └── tender-writer/
```
