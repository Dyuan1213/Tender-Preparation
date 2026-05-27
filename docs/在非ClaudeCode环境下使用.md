# 在非 Claude Code 环境下使用本工具集

> 适用人群：没有安装 Claude Code，使用 Codex、扣子(Coze)、WorkBuddy、ChatGPT、文心/通义/Kimi 等其他平台，或只想在本地手动使用的用户。

## 先理解：这两个 skill 是什么

每个 skill 由两部分组成，**与平台无关、可单独使用**：

1. **`SKILL.md`** —— 一份写给 AI 的"操作说明书/角色设定"，本质就是一段**提示词**。
2. **`scripts/` 里的 Python 脚本** —— 纯 Python 工具，任何能跑 Python 的环境都能用：
   - `tender-parser/scripts/read_docx.py`：把 Word 招标文件提取成纯文本；
   - `tender-writer/scripts/create_tender_docx.py`：把整理好的内容生成**宋体、带页码、含表格/配图**的 Word 标书；
   - `tender-parser/scripts/feishu_upload.py`：（可选）把报告上传到飞书，需要你自己的飞书应用凭据。

> Claude Code 里的"自动触发、`/plugin` 安装"是 Claude 生态专属功能，其他平台没有；但上面的**提示词 + 脚本**可以照常用，只是需要你手动操作。

---

## 准备工作（一次性）

1. **安装 Python 3**（python.org 下载安装）。
2. **安装依赖**：
   ```bash
   pip install python-docx
   ```
3. **获取本工具集**：在 GitHub 仓库页面点 `Code → Download ZIP` 解压，或：
   ```bash
   git clone https://github.com/Dyuan1213/Tender-Preparation.git
   ```
   两个 skill 在 `Tender-Preparation/plugins/tender-tools/skills/` 下。
4. **生成的 Word 用宋体**：查看时电脑需装中文字体（Windows 自带宋体）。

---

## 用法一：解析招标文件（tender-parser）

**目标**：把一份招标文件（Word/PDF）变成一份结构化摘要报告（预算、资质、评分标准、截止日期、实质性要求等）。

### 步骤 1：把招标文件提取成文本
```bash
python plugins/tender-tools/skills/tender-parser/scripts/read_docx.py "你的招标文件.docx" --output extracted.txt
```
> PDF 文件：先另存/转换为 Word，或用 `pdfplumber` 等工具提取文本，得到一份纯文本即可。

### 步骤 2：在任意 AI 平台按指令解析
1. 打开 `tender-parser/SKILL.md`，**复制其全部内容**，粘贴到 AI 平台的**系统提示词 / 角色设定 / 知识库**中（扣子、Coze 智能体可放进"人设与回复逻辑"；ChatGPT 可直接作为第一条消息）。
2. 同时参考 `references/` 下的三个文件，可一并提供给 AI：
   - `extraction-guide.md`（提取规则）、`output-template.md`（报告模板）。
3. 把步骤 1 得到的 `extracted.txt` 文本贴给 AI，说"按上述说明解析这份招标文件"。
4. AI 即按模板输出**速览卡片 + 七大维度完整解析 + 投标准备建议**的 Markdown 报告。

### 步骤 3（可选）：上传到飞书
如需把报告同步到飞书，配置你自己的飞书应用凭据后运行：
```bash
set FEISHU_APP_ID=你的AppID
set FEISHU_APP_SECRET=你的AppSecret
python plugins/tender-tools/skills/tender-parser/scripts/feishu_upload.py report.md
```
> 凭据获取见 `tender-parser/references/feishu-setup.md`。不需要飞书就跳过这一步，报告在本地用即可。

---

## 用法二：编制投标标书（tender-writer）

**目标**：基于招标文件，逐章写出投标文件内容，并生成排版规范的 Word 文档。

### 步骤 1：用 SKILL.md 作为指令，让 AI 逐章写内容
1. 复制 `tender-writer/SKILL.md` 全部内容，作为 AI 的系统提示/角色设定。
2. 把招标文件（或用法一得到的解析报告）提供给 AI。
3. 按 SKILL.md 的流程，与 AI 逐章生成：先确认目录框架与各章字数，再逐章编写内容、逐章确认。

### 步骤 2：把内容整理成一个 JSON 文件
生成 Word 需要一个 `content.json`，格式如下（可让 AI 直接帮你按此格式输出）：
```json
{
  "title": "投标文件",
  "project_name": "XX项目",
  "sections": [
    { "level": 1, "heading": "第一部分 商务文件", "content": "正文……", "type": "fixed_format" },
    { "level": 2, "heading": "2.1 技术方案", "content": "一、概述\n\n正文段落……\n\n|列1|列2|\n|内容1|内容2|" }
  ]
}
```
**content 字段支持的写法：**
- 空行分段；首行自动缩进 2 字符、1.5 倍行距；
- `一、`/`（一）`/`1.` 等有序标题会**自动识别为大纲层级**（Word 导航/目录可用）；
- Markdown 表格：`|列1|列2|` 一行一行写；
- 需人工填写处用 `【请填写：xxx】`，会显示为**红色**提醒；
- 插入图片：单独一行写 `[[IMG:图片绝对路径|图注文字]]`（PNG/JPEG）。

### 步骤 3：生成 Word 文档
```bash
python plugins/tender-tools/skills/tender-writer/scripts/create_tender_docx.py --content-file content.json --output 投标文件.docx
```
得到的 Word：宋体、标题加粗四号、正文小四、1.5 倍行距、首行缩进 2 字符、页脚居中页码、自动封面。

---

## 各平台适配小贴士

- **扣子(Coze) / 文心智能体 / 通义等**：把 `SKILL.md` 内容放入智能体的"人设/提示词"，把 `references/` 内容放入"知识库"；脚本无法在平台内跑，需在本地电脑执行步骤里的 Python 命令。
- **ChatGPT / 带"代码解释器"的平台**：可把 Python 脚本和招标文件一起上传，让其在沙箱里运行脚本；SKILL.md 作为指令贴入对话。
- **任何 LLM**：最低限度——SKILL.md 当提示词 + 本地手动跑脚本，即可完成解析与出 Word。

## 不支持 / 注意事项

- 其他平台**没有自动触发和一键安装**，所有步骤需手动操作。
- 扫描件（图片型 PDF）需先 OCR 成文本。
- 飞书上传是可选功能，需你自己的飞书应用凭据；不影响解析与标书生成。
- 脚本只依赖 `python-docx`，不联网、不上传任何数据（飞书上传除外，且需你主动配置）。
