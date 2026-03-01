# file-mcp

AI 可调用的 MCP 文件服务，支持读写：
- `DOCX`
- `PDF`
- `Excel (.xlsx/.xls)`

## 安装

```bash
npm install
npm run build
```

## 运行

```bash
npm run dev
# 或
npm start
```

## Tools

- `read_docx(file_path)`
- `write_docx(file_path, content)`
- `read_pdf(file_path)`
- `write_pdf(file_path, content)`
- `read_excel(file_path)`
- `write_excel(file_path, data, sheet_name?)`

## Cursor MCP 配置

```json
{
  "mcpServers": {
    "file-mcp": {
      "command": "node",
      "args": ["/Users/mac/Documents/WorkCodeSpace/File-MCP/dist/index.js"],
      "env": {
        "MCP_DEFAULT_ROOT": "/Users/mac/Documents/你的项目目录"
      }
    }
  }
}
```

## 路径规则

- 绝对路径：直接使用
- 相对路径：默认按当前工作区解析
- 可用 `MCP_DEFAULT_ROOT` 强制指定默认根目录

## 参考

- https://github.com/modelcontextprotocol/typescript-sdk
- https://www.npmjs.com/package/@modelcontextprotocol/sdk?activeTab=readme
