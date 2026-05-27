# 飞书集成配置指南

## 第一步：创建飞书自建应用

1. 打开 https://open.feishu.cn → 登录你的飞书账号
2. 点击右上角「开发者后台」→「创建应用」→「自建应用」
3. 填写应用名称（如"招标解析助手"）和描述，点击「确认创建」

## 第二步：开通权限

在应用管理页面 → 「权限管理」，搜索并添加以下权限：

| 权限 | 说明 |
|------|------|
| `drive:drive` | 读写云空间文件（上传DOCX）|
| `drive:drive:readonly` | 读取云空间 |
| `docx:document` | 创建/编辑飞书文档 |

开通后点击右上角「创建版本」→「直接发布」（如有审批流程需等待审核）。

## 第三步：获取凭证

在「凭证与基础信息」页面找到：
- **App ID**（格式：`cli_xxxxxxxxxxxxxxxx`）
- **App Secret**（点击「查看」获取）

## 第四步：配置凭证

在 skill 根目录（`F:\03-AI-claude\skills\tender-parser\`）创建 `.env` 文件：

```
FEISHU_APP_ID=cli_你的AppID
FEISHU_APP_SECRET=你的AppSecret
```

**也可以用环境变量（推荐）：**

```bash
# Windows PowerShell
$env:FEISHU_APP_ID="cli_你的AppID"
$env:FEISHU_APP_SECRET="你的AppSecret"
```

**安全提示：** `.env` 文件已加入 `.gitignore`，不会被提交到版本控制。

## 第五步：（可选）指定目标文件夹

如果希望文档保存到飞书云空间的特定文件夹（而不是根目录）：

1. 在飞书云空间中打开目标文件夹
2. URL 中 `folder/` 后面的那段就是 `folder_token`
   - 例如：`https://bytedance.feishu.cn/drive/folder/Gq4mfXXXXXXXX` → `folder_token = Gq4mfXXXXXXXX`
3. 在 `.env` 中添加：`FEISHU_FOLDER_TOKEN=Gq4mfXXXXXXXX`

## 使用方式

配置好后，解析完成时 skill 会自动执行上传。也可以手动调用：

```bash
python scripts/feishu_upload.py report.md
python scripts/feishu_upload.py report.md --folder-token Gq4mfXXXXXXXX
```

## 常见问题

**Q: 报错 "code: 99991663"**
A: token 无效或过期。检查 app_id/app_secret 是否正确，应用是否已发布。

**Q: 报错 "code: 230001"（权限不足）**
A: 确认已开通 `drive:drive` 和 `docx:document` 权限并已发布版本。

**Q: 文档创建成功但内容乱码**
A: 通常是格式转换问题，可以在飞书里手动调整样式，不影响内容完整性。

**Q: 我用的是飞书国际版（Lark）**
A: 将脚本中的 `BASE_URL` 改为 `https://open.larksuite.com/open-apis`，文档 URL 域名改为 `larksuite.com`。
