import type { GraphDriveItem, GraphSharingLink } from "../types/graph.js";
import { formatDateBR } from "../utils/date.js";
import { formatSize } from "../utils/format.js";

export function formatDriveItems(items: GraphDriveItem[]): string {
  if (items.length === 0) return "Pasta vazia.";

  const formatted = items.map((item) => {
    const isFolder = !!item.folder;
    const icon = isFolder ? "[pasta]" : "[arquivo]";
    const size = item.size ? ` (${formatSize(item.size)})` : "";
    const modified = formatDateBR(item.lastModifiedDateTime);
    const childCount = item.folder?.childCount != null ? ` — ${item.folder.childCount} itens` : "";
    return `- ${icon} **${item.name}**${size}${childCount}\n  Modificado: ${modified}\n  ID: ${item.id}`;
  });

  return `## OneDrive (${formatted.length} itens)\n\n${formatted.join("\n\n")}`;
}

export function formatFileSearchResults(
  query: string,
  items: GraphDriveItem[]
): string {
  if (items.length === 0) return `Nenhum arquivo encontrado para "${query}".`;

  const formatted = items.map((item) => {
    const size = item.size ? ` (${formatSize(item.size)})` : "";
    const path = item.parentReference?.path ?? "";
    return `- **${item.name}**${size}\n  Caminho: ${path}\n  URL: ${item.webUrl}\n  ID: ${item.id}`;
  });

  return `## Resultados "${query}" (${formatted.length})\n\n${formatted.join("\n\n")}`;
}

export function formatUploadResult(item: GraphDriveItem): string {
  return `Arquivo enviado com sucesso!\n\n**${item.name}** (${formatSize(item.size ?? 0)})\nURL: ${item.webUrl}\nID: ${item.id}`;
}

export function formatShareLink(result: GraphSharingLink): string {
  const link = result.link;
  return `Link criado!\n\n**URL:** ${link?.webUrl}\n**Tipo:** ${link?.type}\n**Escopo:** ${link?.scope}`;
}
