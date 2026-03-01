# file-mcp

基于 Node.js + TypeScript 的 MCP 服务，提供 AI 对以下文件格式的读写能力：

- `DOCX`（读取纯文本、写入 Word）
- `PDF`（读取文本、写入 PDF）
- `Excel`（读取全部 Sheet 为 JSON、写入 `.xlsx`）

实现参考了 MCP 官方 TypeScript SDK 的 stdio server 用法。

## 1) 安装与构建

```bash
npm install
npm run build
```

开发调试：

```bash
npm run dev
```

生产运行：

```bash
npm start
```

## 2) 可用 MCP Tools

- `read_docx`
  - 入参：`file_path: string`
  - 返回：DOCX 提取出的纯文本
- `write_docx`
  - 入参：`file_path: string`, `content: string`
  - 说明：支持 `#`/`##`/`###` 作为标题
- `read_pdf`
  - 入参：`file_path: string`
  - 返回：PDF 提取文本
- `write_pdf`
  - 入参：`file_path: string`, `content: string`
  - 说明：按行写入，自动分页
- `read_excel`
  - 入参：`file_path: string`
  - 返回：`{ sheets, data }`，其中 `data` 为每个 sheet 的 JSON 数组
- `write_excel`
  - 入参：`file_path: string`, `data: object[]`, `sheet_name?: string`
  - 说明：将 JSON 数组写入 `.xlsx`

## 3) Cursor MCP 配置示例

你可以在 Cursor 的 MCP 配置中添加如下 server（根据你的本机路径修改）：

```json
{
  "mcpServers": {
    "file-mcp": {
      "command": "node",
      "args": [
        "/Users/mac/Documents/WorkCodeSpace/File-MCP/dist/index.js"
      ]
    }
  }
}
```

如果你希望开发模式直接运行，也可以改成：

```json
{
  "mcpServers": {
    "file-mcp-dev": {
      "command": "npx",
      "args": [
        "tsx",
        "/Users/mac/Documents/WorkCodeSpace/File-MCP/src/index.ts"
      ]
    }
  }
}
```

## 4) 注意事项

- 相对路径会默认基于“当前工作区路径”解析（优先读取 `MCP_DEFAULT_ROOT` / `CURSOR_WORKSPACE_PATH` 等环境变量，最后回退到进程 `cwd`）。
- 如需强制指定默认目录，可在 MCP 配置里加 `env.MCP_DEFAULT_ROOT`。
- `write_pdf` 是文本写入，不保留原始 PDF 排版。
- `read_docx` 与 `read_pdf` 以文本抽取为主，不保证 100% 还原复杂格式（表格、图文混排等）。

示例（固定默认根目录）：

```json
{
  "mcpServers": {
    "file-mcp": {
      "command": "node",
      "args": [
        "/Users/mac/Documents/WorkCodeSpace/File-MCP/dist/index.js"
      ],
      "env": {
        "MCP_DEFAULT_ROOT": "/Users/mac/Documents/你的项目目录"
      }
    }
  }
}
```

## 5) 参考

- MCP TypeScript SDK：<https://github.com/modelcontextprotocol/typescript-sdk>
- npm SDK（v1）：<https://www.npmjs.com/package/@modelcontextprotocol/sdk?activeTab=readme>
