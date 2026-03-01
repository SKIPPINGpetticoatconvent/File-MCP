/**
 * 读取 DOCX 为纯文本（基于 HTML 转 Markdown 风格）
 */
export declare function readDocx(filePath: string): Promise<string>;
/**
 * 将纯文本写入 DOCX
 */
export declare function writeDocx(filePath: string, content: string): Promise<void>;
/**
 * 读取 PDF 文本
 */
export declare function readPdf(filePath: string): Promise<string>;
/**
 * 将文本写入新 PDF（每段一段落）
 */
export declare function writePdf(filePath: string, content: string): Promise<void>;
/**
 * 读取 Excel：返回首个 sheet 的 JSON 数组及 sheet 名列表
 */
export declare function readExcel(filePath: string): Promise<{
    sheets: string[];
    data: Record<string, unknown[]>;
}>;
/**
 * 将 JSON 数组写入 Excel
 */
export declare function writeExcel(filePath: string, data: unknown[], sheetName?: string): Promise<void>;
