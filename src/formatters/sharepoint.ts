import type {
  GraphSite,
  GraphDrive,
  GraphDriveItem,
  GraphSearchResponse,
} from "../types/graph.js";
import { formatSize } from "../utils/format.js";

export function formatSiteList(sites: GraphSite[]): string {
  if (sites.length === 0) return "Nenhum site encontrado.";

  const formatted = sites.map(
    (s) =>
      `- **${s.displayName}** (${s.name})\n  ${s.description ?? ""}\n  URL: ${s.webUrl}\n  ID: ${s.id}`
  );

  return `## Sites SharePoint (${formatted.length})\n\n${formatted.join("\n\n")}`;
}

export function formatSiteDetail(site: GraphSite): string {
  return [
    `## ${site.displayName}`,
    "",
    `**Nome:** ${site.name}`,
    `**Descrição:** ${site.description ?? "N/A"}`,
    `**URL:** ${site.webUrl}`,
    `**ID:** ${site.id}`,
  ].join("\n");
}

export function formatDocumentLibraries(drives: GraphDrive[]): string {
  if (drives.length === 0) return "Nenhuma biblioteca encontrada.";

  const formatted = drives.map(
    (d) =>
      `- **${d.name}** (${d.driveType})\n  ${d.description ?? ""}\n  URL: ${d.webUrl}\n  ID: ${d.id}`
  );

  return `## Bibliotecas de Documentos (${formatted.length})\n\n${formatted.join("\n\n")}`;
}

export function formatLibraryItems(items: GraphDriveItem[]): string {
  if (items.length === 0) return "Nenhum item encontrado.";

  const formatted = items.map((item) => {
    const isFolder = !!item.folder;
    const icon = isFolder ? "[pasta]" : "[arquivo]";
    const size = item.size ? ` (${formatSize(item.size)})` : "";
    return `- ${icon} **${item.name}**${size}\n  URL: ${item.webUrl}\n  ID: ${item.id}`;
  });

  return `## Itens (${formatted.length})\n\n${formatted.join("\n\n")}`;
}

export function formatSearchResults(
  query: string,
  results: GraphSearchResponse[]
): string {
  if (!results?.length || !results[0]?.hitsContainers?.length) {
    return `Nenhum resultado encontrado para "${query}".`;
  }

  const hits = results[0].hitsContainers[0].hits ?? [];
  const formatted = hits.map((hit) => {
    const resource = hit.resource;
    return `- **${resource?.name ?? "Sem nome"}** (${hit.hitId})\n  Resumo: ${hit.summary ?? "N/A"}\n  URL: ${resource?.webUrl ?? "N/A"}`;
  });

  return `## Resultados SharePoint "${query}" (${formatted.length})\n\n${formatted.join("\n\n")}`;
}
