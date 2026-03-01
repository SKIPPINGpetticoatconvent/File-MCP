import * as fs from "node:fs/promises";
import * as nodeFs from "node:fs";
import * as path from "node:path";
import { createRequire } from "node:module";
import mammoth from "mammoth";
import { Document, Packer, Paragraph, TextRun, HeadingLevel, } from "docx";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import * as XLSX from "xlsx";
const require = createRequire(import.meta.url);
// pdf-parse 为 CommonJS，在 ESM 中通过 createRequire 引入
const pdfParse = require("pdf-parse");
function isExistingDirectory(dirPath) {
    if (!dirPath || !path.isAbsolute(dirPath)) {
        return false;
    }
    try {
        return nodeFs.statSync(dirPath).isDirectory();
    }
    catch {
        return false;
    }
}
function detectDefaultRoot() {
    const prioritizedCandidates = [
        process.env.MCP_DEFAULT_ROOT,
        // Cursor/VSCode 常见工作区环境变量（若客户端注入）
        process.env.CURSOR_WORKSPACE_PATH,
        process.env.CURSOR_PROJECT_PATH,
        process.env.WORKSPACE_PATH,
        process.env.VSCODE_CWD,
        process.env.PROJECT_ROOT,
    ].filter((v) => Boolean(v));
    for (const candidate of prioritizedCandidates) {
        if (isExistingDirectory(candidate)) {
            return candidate;
        }
    }
    // 再从环境变量里自动探测可能的工作区路径，尽量不要求用户手动配置
    const dynamicCandidates = Object.entries(process.env)
        .filter(([key, value]) => {
        if (!value)
            return false;
        const normalizedKey = key.toUpperCase();
        const maybeWorkspace = normalizedKey.includes("WORKSPACE") ||
            normalizedKey.includes("PROJECT") ||
            normalizedKey.includes("ROOT") ||
            normalizedKey.includes("CURSOR");
        return maybeWorkspace && path.isAbsolute(value);
    })
        .map(([, value]) => value);
    for (const candidate of dynamicCandidates) {
        if (isExistingDirectory(candidate)) {
            return candidate;
        }
    }
    return process.cwd();
}
function resolveFilePath(inputPath) {
    if (path.isAbsolute(inputPath)) {
        return inputPath;
    }
    // 默认自动使用当前打开的工作区路径，其次回退到进程 cwd
    const baseDir = detectDefaultRoot();
    return path.resolve(baseDir, inputPath);
}
/**
 * 读取 DOCX 为纯文本（基于 HTML 转 Markdown 风格）
 */
export async function readDocx(filePath) {
    const resolved = resolveFilePath(filePath);
    const buf = await fs.readFile(resolved);
    const result = await mammoth.extractRawText({ buffer: buf });
    return result.value;
}
/**
 * 将纯文本写入 DOCX
 */
export async function writeDocx(filePath, content) {
    const resolved = resolveFilePath(filePath);
    const lines = content.split(/\r?\n/).filter((l) => l.trim() !== "");
    const children = lines.map((line) => {
        const isHeading = /^#+\s/.test(line);
        const level = line.match(/^(#+)\s/)?.[1].length ?? 0;
        const text = line.replace(/^#+\s*/, "").trim();
        if (isHeading && level <= 3) {
            const headingLevel = level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3;
            return new Paragraph({
                text,
                heading: headingLevel,
            });
        }
        return new Paragraph({
            children: [new TextRun(text)],
        });
    });
    const doc = new Document({
        sections: [
            {
                properties: {},
                children: children.length > 0 ? children : [new Paragraph({ children: [new TextRun("")] })],
            },
        ],
    });
    const blob = await Packer.toBuffer(doc);
    await fs.writeFile(resolved, blob);
}
/**
 * 读取 PDF 文本
 */
export async function readPdf(filePath) {
    const resolved = resolveFilePath(filePath);
    const buf = await fs.readFile(resolved);
    const data = await pdfParse(buf);
    return typeof data.text === "string" ? data.text : String(data.text ?? "");
}
/**
 * 将文本写入新 PDF（每段一段落）
 */
export async function writePdf(filePath, content) {
    const resolved = resolveFilePath(filePath);
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const lines = content.split(/\r?\n/).filter((l) => l.trim() !== "");
    let y = 750;
    const lineHeight = 14;
    const margin = 50;
    const pageWidth = 550;
    if (lines.length > 0) {
        doc.addPage();
    }
    for (const line of lines) {
        if (y < 50) {
            doc.addPage();
            y = 750;
        }
        const page = doc.getPage(doc.getPageCount() - 1);
        page.drawText(line.slice(0, 200), {
            x: margin,
            y,
            size: 12,
            font,
            color: rgb(0, 0, 0),
            maxWidth: pageWidth,
        });
        y -= lineHeight;
    }
    if (lines.length === 0) {
        doc.addPage();
    }
    const pdfBytes = await doc.save();
    await fs.writeFile(resolved, pdfBytes);
}
/**
 * 读取 Excel：返回首个 sheet 的 JSON 数组及 sheet 名列表
 */
export async function readExcel(filePath) {
    const resolved = resolveFilePath(filePath);
    const buf = await fs.readFile(resolved);
    const wb = XLSX.read(buf, { type: "buffer" });
    const sheets = wb.SheetNames;
    const data = {};
    for (const name of sheets) {
        const sheet = wb.Sheets[name];
        data[name] = XLSX.utils.sheet_to_json(sheet);
    }
    return { sheets, data };
}
/**
 * 将 JSON 数组写入 Excel
 */
export async function writeExcel(filePath, data, sheetName = "Sheet1") {
    const resolved = resolveFilePath(filePath);
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
    await fs.writeFile(resolved, buf);
}
