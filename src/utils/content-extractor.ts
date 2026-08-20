import { extractText, getDocumentProxy } from "unpdf";
import mammoth from "mammoth";

const SUPPORTED_TEXT_EXTENSIONS = new Set([
  ".txt", ".csv", ".json", ".xml", ".html", ".htm", ".md",
  ".yaml", ".yml", ".log", ".ini", ".cfg", ".conf", ".env",
  ".ts", ".js", ".py", ".sh", ".sql", ".css", ".scss",
]);

export interface ExtractResult {
  text: string;
  totalPages?: number;
  pagesReturned?: string;
}

export function getFileExtension(nameOrPath: string): string {
  const dot = nameOrPath.lastIndexOf(".");
  return dot >= 0 ? nameOrPath.slice(dot).toLowerCase() : "";
}

export function isTextFile(nameOrPath: string): boolean {
  return SUPPORTED_TEXT_EXTENSIONS.has(getFileExtension(nameOrPath));
}

export function isSupportedBinary(nameOrPath: string): boolean {
  const ext = getFileExtension(nameOrPath);
  return ext === ".pdf" || ext === ".docx";
}

export async function extractFromPdf(
  data: Uint8Array,
  startPage = 1,
  maxPages?: number
): Promise<ExtractResult> {
  const pdf = await getDocumentProxy(data);
  const totalPages = pdf.numPages;
  const end = maxPages ? Math.min(startPage + maxPages - 1, totalPages) : totalPages;

  const result = await extractText(pdf, { mergePages: false });
  await pdf.cleanup();

  const pages = (result.text as string[]).slice(startPage - 1, end);
  const text = pages.join("\n\n---\n\n");

  if (text.trim().length === 0) {
    return {
      text: `[PDF sem camada de texto — ${totalPages} página(s). Provável digitalização; exige OCR para leitura.]`,
      totalPages,
      pagesReturned: "0",
    };
  }

  return {
    text,
    totalPages,
    pagesReturned: startPage === 1 && end === totalPages
      ? `1-${totalPages}`
      : `${startPage}-${end}`,
  };
}

export async function extractFromDocx(data: Uint8Array): Promise<ExtractResult> {
  const result = await mammoth.extractRawText({ buffer: Buffer.from(data) });
  if (result.value.trim().length === 0) {
    return { text: "[DOCX sem conteúdo de texto extraível.]" };
  }
  return { text: result.value };
}
