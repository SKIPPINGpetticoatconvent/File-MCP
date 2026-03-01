#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  readDocx,
  writeDocx,
  readPdf,
  writePdf,
  readExcel,
  writeExcel,
} from "./handlers.js";

const server = new McpServer({
  name: "file-mcp",
  version: "1.0.0",
});

// ----- DOCX -----
server.tool(
  "read_docx",
  {
    file_path: z.string().describe("DOCX 文件的路径（绝对或相对）"),
  },
  async ({ file_path }) => {
    try {
      const text = await readDocx(file_path);
      return { content: [{ type: "text" as const, text }] };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return {
        content: [{ type: "text" as const, text: `读取 DOCX 失败: ${message}` }],
        isError: true,
      };
    }
  }
);

server.tool(
  "write_docx",
  {
    file_path: z.string().describe("要写入的 DOCX 文件路径"),
    content: z.string().describe("要写入的纯文本内容，支持 # ## ### 作为标题"),
  },
  async ({ file_path, content }) => {
    try {
      await writeDocx(file_path, content);
      return {
        content: [{ type: "text" as const, text: `已成功写入: ${file_path}` }],
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return {
        content: [{ type: "text" as const, text: `写入 DOCX 失败: ${message}` }],
        isError: true,
      };
    }
  }
);

// ----- PDF -----
server.tool(
  "read_pdf",
  {
    file_path: z.string().describe("PDF 文件的路径（绝对或相对）"),
  },
  async ({ file_path }) => {
    try {
      const text = await readPdf(file_path);
      return { content: [{ type: "text" as const, text }] };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return {
        content: [{ type: "text" as const, text: `读取 PDF 失败: ${message}` }],
        isError: true,
      };
    }
  }
);

server.tool(
  "write_pdf",
  {
    file_path: z.string().describe("要写入的 PDF 文件路径"),
    content: z.string().describe("要写入的纯文本内容，每行一段"),
  },
  async ({ file_path, content }) => {
    try {
      await writePdf(file_path, content);
      return {
        content: [{ type: "text" as const, text: `已成功写入: ${file_path}` }],
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return {
        content: [{ type: "text" as const, text: `写入 PDF 失败: ${message}` }],
        isError: true,
      };
    }
  }
);

// ----- Excel -----
server.tool(
  "read_excel",
  {
    file_path: z.string().describe("Excel 文件路径（.xlsx/.xls）"),
  },
  async ({ file_path }) => {
    try {
      const { sheets, data } = await readExcel(file_path);
      const out = JSON.stringify({ sheets, data }, null, 2);
      return { content: [{ type: "text" as const, text: out }] };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return {
        content: [{ type: "text" as const, text: `读取 Excel 失败: ${message}` }],
        isError: true,
      };
    }
  }
);

server.tool(
  "write_excel",
  {
    file_path: z.string().describe("要写入的 Excel 文件路径（.xlsx）"),
    data: z
      .array(z.record(z.unknown()))
      .describe("要写入的数据，数组每项为一行对象"),
    sheet_name: z.string().optional().describe("工作表名称，默认 Sheet1"),
  },
  async ({ file_path, data, sheet_name }) => {
    try {
      await writeExcel(file_path, data, sheet_name ?? "Sheet1");
      return {
        content: [{ type: "text" as const, text: `已成功写入: ${file_path}` }],
      };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return {
        content: [{ type: "text" as const, text: `写入 Excel 失败: ${message}` }],
        isError: true,
      };
    }
  }
);

const transport = new StdioServerTransport();
await server.connect(transport);
