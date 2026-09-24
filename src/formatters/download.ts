import type { StoredDownload } from "../utils/download-store.js";
import { formatSize } from "../utils/format.js";

export function formatStoredDownloads(downloads: readonly StoredDownload[]): string {
  const lines = [downloads.length === 1 ? "## Arquivo baixado" : `## ${downloads.length} arquivos baixados`, ""];
  for (const download of downloads) {
    lines.push(`- **${download.displayName}** (${formatSize(download.sizeBytes)})`);
    lines.push(`  Caminho local: ${download.path}`);
  }
  return lines.join("\n");
}
