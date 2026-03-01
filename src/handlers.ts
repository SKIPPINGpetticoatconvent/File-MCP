import * as fs from "node:fs/promises";
import * as nodeFs from "node:fs";
import * as path from "node:path";
import { createRequire } from "node:module";
import mammoth from "mammoth";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
} from "docx";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import * as XLSX from "xlsx";
import JSZip from "jszip";

const require = createRequire(import.meta.url);
// pdf-parse 为 CommonJS，在 ESM 中通过 createRequire 引入
const pdfParse = require("pdf-parse");
// pptxgenjs 为 CommonJS，在 ESM 中通过 createRequire 引入
const PptxGenJS = require("pptxgenjs");

function isExistingDirectory(dirPath: string): boolean {
  if (!dirPath || !path.isAbsolute(dirPath)) {
    return false;
  }
  try {
    return nodeFs.statSync(dirPath).isDirectory();
  } catch {
    return false;
  }
}

function detectDefaultRoot(): string {
  const prioritizedCandidates = [
    process.env.MCP_DEFAULT_ROOT,
    // Cursor/VSCode 常见工作区环境变量（若客户端注入）
    process.env.CURSOR_WORKSPACE_PATH,
    process.env.CURSOR_PROJECT_PATH,
    process.env.WORKSPACE_PATH,
    process.env.VSCODE_CWD,
    process.env.PROJECT_ROOT,
  ].filter((v): v is string => Boolean(v));

  for (const candidate of prioritizedCandidates) {
    if (isExistingDirectory(candidate)) {
      return candidate;
    }
  }

  // 再从环境变量里自动探测可能的工作区路径，尽量不要求用户手动配置
  const dynamicCandidates = Object.entries(process.env)
    .filter(([key, value]) => {
      if (!value) return false;
      const normalizedKey = key.toUpperCase();
      const maybeWorkspace =
        normalizedKey.includes("WORKSPACE") ||
        normalizedKey.includes("PROJECT") ||
        normalizedKey.includes("ROOT") ||
        normalizedKey.includes("CURSOR");
      return maybeWorkspace && path.isAbsolute(value);
    })
    .map(([, value]) => value as string);

  for (const candidate of dynamicCandidates) {
    if (isExistingDirectory(candidate)) {
      return candidate;
    }
  }

  return process.cwd();
}

function resolveFilePath(inputPath: string): string {
  if (path.isAbsolute(inputPath)) {
    return inputPath;
  }

  // 默认自动使用当前打开的工作区路径，其次回退到进程 cwd
  const baseDir = detectDefaultRoot();

  return path.resolve(baseDir, inputPath);
}

function ensurePptxExt(filePath: string): void {
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== ".pptx") {
    throw new Error(`不支持的 PPT 扩展名: ${ext || "(无扩展名)"}，仅支持 .pptx`);
  }
}

function decodeXmlEntities(input: string): string {
  return input
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

const COMMON_TEXT_EXTENSIONS = new Set([
  ".txt",
  ".md",
  ".markdown",
  ".html",
  ".htm",
  ".xml",
  ".yaml",
  ".yml",
]);

function ensureExt(filePath: string, extensions: Set<string>, typeName: string): void {
  const ext = path.extname(filePath).toLowerCase();
  if (!extensions.has(ext)) {
    throw new Error(`不支持的 ${typeName} 扩展名: ${ext || "(无扩展名)"}`);
  }
}

export async function readCommonTextFile(filePath: string): Promise<string> {
  ensureExt(filePath, COMMON_TEXT_EXTENSIONS, "文本文件");
  const resolved = resolveFilePath(filePath);
  return fs.readFile(resolved, "utf8");
}

export async function writeCommonTextFile(filePath: string, content: string): Promise<void> {
  ensureExt(filePath, COMMON_TEXT_EXTENSIONS, "文本文件");
  const resolved = resolveFilePath(filePath);
  await fs.writeFile(resolved, content, "utf8");
}

export async function readJsonFile(filePath: string): Promise<unknown> {
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== ".json") {
    throw new Error(`不支持的 JSON 扩展名: ${ext || "(无扩展名)"}`);
  }
  const resolved = resolveFilePath(filePath);
  const text = await fs.readFile(resolved, "utf8");
  return JSON.parse(text);
}

export async function writeJsonFile(filePath: string, data: unknown): Promise<void> {
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== ".json") {
    throw new Error(`不支持的 JSON 扩展名: ${ext || "(无扩展名)"}`);
  }
  const resolved = resolveFilePath(filePath);
  await fs.writeFile(resolved, `${JSON.stringify(data, null, 2)}\n`, "utf8");
}

export async function readCsvFile(filePath: string): Promise<Record<string, unknown>[]> {
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== ".csv") {
    throw new Error(`不支持的 CSV 扩展名: ${ext || "(无扩展名)"}`);
  }
  const resolved = resolveFilePath(filePath);
  const csvText = await fs.readFile(resolved, "utf8");
  const wb = XLSX.read(csvText, { type: "string" });
  const firstSheetName = wb.SheetNames[0];
  if (!firstSheetName) {
    return [];
  }
  const sheet = wb.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet);
}

export async function writeCsvFile(filePath: string, data: Record<string, unknown>[]): Promise<void> {
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== ".csv") {
    throw new Error(`不支持的 CSV 扩展名: ${ext || "(无扩展名)"}`);
  }
  const resolved = resolveFilePath(filePath);
  const ws = XLSX.utils.json_to_sheet(data);
  const csvText = XLSX.utils.sheet_to_csv(ws);
  await fs.writeFile(resolved, csvText, "utf8");
}

/**
 * 读取 PPTX：提取每页文本
 */
export async function readPptx(filePath: string): Promise<string> {
  ensurePptxExt(filePath);
  const resolved = resolveFilePath(filePath);
  const buf = await fs.readFile(resolved);
  const zip = await JSZip.loadAsync(buf);

  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((a, b) => {
      const aNum = Number(a.match(/slide(\d+)\.xml/i)?.[1] ?? "0");
      const bNum = Number(b.match(/slide(\d+)\.xml/i)?.[1] ?? "0");
      return aNum - bNum;
    });

  const pages: string[] = [];
  for (const slidePath of slideFiles) {
    const xml = await zip.file(slidePath)?.async("string");
    if (!xml) {
      continue;
    }
    const texts = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decodeXmlEntities(m[1]).trim());
    const cleaned = texts.filter((t) => t.length > 0).join("\n");
    pages.push(cleaned);
  }

  if (pages.length === 0) {
    return "";
  }

  return pages
    .map((pageText, idx) => `# Slide ${idx + 1}\n${pageText}`)
    .join("\n\n");
}

/**
 * 写入 PPTX：用 --- 分割多页
 */
export async function writePptx(filePath: string, content: string): Promise<void> {
  ensurePptxExt(filePath);
  const resolved = resolveFilePath(filePath);
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";

  const rawSlides = content
    .split(/\r?\n-{3,}\r?\n/g)
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
  const slides = rawSlides.length > 0 ? rawSlides : [content.trim()];

  for (const slideText of slides) {
    const slide = pptx.addSlide();
    slide.addText(slideText || " ", {
      x: 0.6,
      y: 0.6,
      w: 12.0,
      h: 6.0,
      fontSize: 20,
      color: "1F2937",
      valign: "top",
      fit: "shrink",
      breakLine: false,
    });
  }

  await pptx.writeFile({ fileName: resolved });
}

/**
 * 读取 DOCX 为纯文本（基于 HTML 转 Markdown 风格）
 */
export async function readDocx(filePath: string): Promise<string> {
  const resolved = resolveFilePath(filePath);
  const buf = await fs.readFile(resolved);
  const result = await mammoth.extractRawText({ buffer: buf });
  return result.value;
}

/**
 * 将纯文本写入 DOCX
 */
export async function writeDocx(filePath: string, content: string): Promise<void> {
  const resolved = resolveFilePath(filePath);
  const lines = content.split(/\r?\n/).filter((l) => l.trim() !== "");
  const children = lines.map((line) => {
    const isHeading = /^#+\s/.test(line);
    const level = line.match(/^(#+)\s/)?.[1].length ?? 0;
    const text = line.replace(/^#+\s*/, "").trim();
    if (isHeading && level <= 3) {
      const headingLevel =
        level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3;
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
export async function readPdf(filePath: string): Promise<string> {
  const resolved = resolveFilePath(filePath);
  const buf = await fs.readFile(resolved);
  const data = await pdfParse(buf);
  return typeof data.text === "string" ? data.text : String(data.text ?? "");
}

/**
 * 将文本写入新 PDF（每段一段落）
 */
export async function writePdf(filePath: string, content: string): Promise<void> {
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
export async function readExcel(filePath: string): Promise<{ sheets: string[]; data: Record<string, unknown[]> }> {
  const resolved = resolveFilePath(filePath);
  const buf = await fs.readFile(resolved);
  const wb = XLSX.read(buf, { type: "buffer" });
  const sheets = wb.SheetNames;
  const data: Record<string, unknown[]> = {};
  for (const name of sheets) {
    const sheet = wb.Sheets[name];
    data[name] = XLSX.utils.sheet_to_json(sheet);
  }
  return { sheets, data };
}

/**
 * 将 JSON 数组写入 Excel
 */
export async function writeExcel(
  filePath: string,
  data: unknown[],
  sheetName: string = "Sheet1"
): Promise<void> {
  const resolved = resolveFilePath(filePath);
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  await fs.writeFile(resolved, buf);
}
