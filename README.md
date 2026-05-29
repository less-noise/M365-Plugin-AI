# M365 AI 插件

集成 **DeepSeek AI API** 的 Microsoft 365 桌面插件，支持 **Word、Excel、PowerPoint**。所有代码在本地运行，无需云服务器。

## 功能

| 应用 | 能力 |
|------|------|
| **Word** | 学术润色、格式整理、语气调整、生成摘要、续写扩展 |
| **Excel** | 数据清洗建议、自然语言问答、趋势分析、公式生成 |
| **PowerPoint** | 演讲大纲生成、幻灯片内容撰写、备注稿生成 |

AI 响应以流式输出，结果可直接写入文档或复制到剪贴板。

## 架构

```
你的电脑（本地）                          DeepSeek API（云端）
┌────────────────────────────┐            ┌──────────────┐
│ Office（Word/Excel/PPT）    │  fetch()   │  DeepSeek AI │
│  侧边栏（Office.js）        │ ──────────►│  推理服务     │
│                            │ ◄───────── │              │
│ http-server :3000          │  流式返回   │              │
│（本地 HTTPS，仅本机可访问）  │            └──────────────┘
└────────────────────────────┘
```

- **http-server** 在本地提供 HTTPS 静态文件服务，仅监听 `127.0.0.1`
- **Office.js** 负责读写文档内容（选区、单元格、幻灯片）
- **DeepSeek API** 处理所有 AI 推理，API Key 仅保存在你本地

## 环境要求

- **Windows 10 / 11**
- **Microsoft 365** 桌面版（Word、Excel、PowerPoint）
- **Node.js** LTS → [下载](https://nodejs.org)
- **DeepSeek API Key** → [申请](https://platform.deepseek.com/api_keys)

## 安装（仅需一次）

### 1. 克隆项目

```bash
git clone https://github.com/你的用户名/office-ai-addin.git
cd office-ai-addin
```

### 2. 运行安装

双击 **install.bat**，脚本会自动：

- 检查 Node.js 环境
- 安装 npm 依赖
- 生成并信任本地 HTTPS 证书（如有安全弹窗，请点"是"）
- 注册插件清单到 Office

### 3. 填写 API Key

编辑 `config/config.json`：

```json
{
  "apiKey": "sk-你的DeepSeek密钥",
  "models": {
    "word": "deepseek-v4-flash",
    "excel": "deepseek-v4-flash",
    "ppt": "deepseek-v4-pro"
  },
  "language": "zh-CN"
}
```

## 使用

1. 双击 **start.bat**
2. 选择要打开的应用：`W`（Word）/ `E`（Excel）/ `P`（PPT）/ `A`（全部）
3. Office 自动启动，侧边栏自动弹出
4. **不要关闭服务器窗口**，使用完毕后关掉即可

### Word

在文档中选中文字 → 点击「获取选中文字」→ 选择功能（学术润色/格式整理/语气调整/生成摘要/续写扩展）→ 点击「开始处理」→ 结果流式输出 → 点击「替换原文」写回文档。

### Excel

选中数据区域 → 选择功能（数据清洗/自然语言提问/数据趋势描述/生成公式）→ 输入问题或指令 → 点击「开始处理」。

### PowerPoint

选择功能（生成演讲大纲/幻灯片文字内容/备注稿）→ 输入主题或内容 → 点击「开始生成」。

## 目录结构

```
office-ai-addin/
├── install.bat              一次性安装
├── start.bat                启动服务 + Office
├── stop.bat                 停止服务
├── package.json
├── config/
│   ├── config.json          你的 API Key（不提交）
│   └── config.example.json  配置模板
├── manifest/
│   ├── manifest-word.xml
│   ├── manifest-excel.xml
│   └── manifest-ppt.xml
├── shared/
│   ├── api.js               DeepSeek API 调用（流式）
│   ├── config-loader.js     配置加载
│   └── style.css            共用 UI
├── word/
│   ├── taskpane.html
│   └── taskpane.js
├── excel/
│   ├── taskpane.html
│   └── taskpane.js
├── ppt/
│   ├── taskpane.html
│   └── taskpane.js
└── scripts/
    └── generate-certs.js    证书生成（回退方案）
```

## 技术栈

- Office Add-in 框架（Office.js / Task Pane）
- 原生 HTML + CSS + JavaScript（无框架）
- http-server 本地 HTTPS 服务
- office-addin-dev-certs 证书管理
- office-addin-debugging 插件注入
- DeepSeek API（Anthropic 兼容端点，SSE 流式）

## 常见问题

**打开 Office 找不到插件？**

确认 `start.bat` 窗口在运行，然后重启 Office。本插件通过调试工具注入，`start.bat` 会自动处理。

**提示"Failed to fetch"？**

确认 `start.bat` 窗口未关闭，端口 3000 未被防火墙阻止。

**浏览器提示证书不安全？**

以管理员身份重新运行 `install.bat`。

**WPS 抢了 .docx 文件关联？**

右键任意 `.docx` 文件 → 打开方式 → 选择 Microsoft Word → 始终使用此应用。

## 安全

- API Key 仅存储在本地 `config/config.json` 中，已加入 `.gitignore`
- 所有 AI 请求从你电脑直连 DeepSeek，无数据中转
- 本地 HTTPS 服务仅监听 `127.0.0.1`，外网无法访问

## 许可证

MIT
