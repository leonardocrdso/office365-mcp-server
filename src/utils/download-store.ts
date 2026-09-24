import { randomUUID } from "node:crypto";
import { mkdir, readdir, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { resolveStoragePaths } from "../auth/storage.js";
import { DOWNLOAD_MAX_BYTES, DOWNLOAD_RETENTION_MS, DOWNLOAD_SWEEP_INTERVAL_MS } from "../constants.js";
import { formatSize } from "./format.js";

export interface StoredDownload {
  readonly path: string;
  readonly fileName: string;
  readonly displayName: string;
  readonly sizeBytes: number;
}

const FALLBACK_FILE_NAME = "arquivo";
const FORBIDDEN_FILE_NAME_CHARS = /[\u0000-\u001f<>:"|?*]/g;

export function resolveDownloadRoot(): string {
  const { baseDir } = resolveStoragePaths();
  return baseDir ? join(baseDir, "downloads") : join(tmpdir(), "office365-mcp-downloads");
}

export function sanitizeFileName(rawName: string): string {
  const lastSegment = basename(rawName.replace(/\\/g, "/"));
  const cleaned = lastSegment.replace(FORBIDDEN_FILE_NAME_CHARS, "_").trim();
  const isUsable = cleaned.length > 0 && cleaned !== "." && cleaned !== "..";
  return isUsable ? cleaned : FALLBACK_FILE_NAME;
}

export function assertDownloadable(fileName: string, sizeBytes: number): void {
  if (sizeBytes <= DOWNLOAD_MAX_BYTES) return;
  throw new Error(
    `'${fileName}' tem ${formatSize(sizeBytes)}, acima do limite de ${formatSize(DOWNLOAD_MAX_BYTES)} por arquivo.`
  );
}

async function removeDirectoryIfExpired(directory: string, now: number): Promise<void> {
  const directoryStat = await stat(directory).catch(() => undefined);
  if (!directoryStat || now - directoryStat.mtimeMs <= DOWNLOAD_RETENTION_MS) return;
  await rm(directory, { recursive: true, force: true });
}

export async function sweepExpiredDownloads(now: number = Date.now()): Promise<void> {
  const root = resolveDownloadRoot();
  const entries = await readdir(root, { withFileTypes: true }).catch(() => []);
  const directories = entries.filter((entry) => entry.isDirectory());
  await Promise.all(directories.map((entry) => removeDirectoryIfExpired(join(root, entry.name), now)));
}

export function startDownloadSweeper(): void {
  const runSweep = (): void => {
    sweepExpiredDownloads().catch((error: unknown) => {
      console.error(JSON.stringify({ event: "download_sweep_failed", error: String(error) }));
    });
  };
  runSweep();
  setInterval(runSweep, DOWNLOAD_SWEEP_INTERVAL_MS).unref();
}

export async function storeDownload(rawName: string, content: Uint8Array): Promise<StoredDownload> {
  const fileName = sanitizeFileName(rawName);
  assertDownloadable(fileName, content.byteLength);
  await sweepExpiredDownloads();
  const directory = join(resolveDownloadRoot(), randomUUID());
  await mkdir(directory, { recursive: true });
  const path = join(directory, fileName);
  await writeFile(path, content);
  return { path, fileName, displayName: fileName, sizeBytes: content.byteLength };
}
