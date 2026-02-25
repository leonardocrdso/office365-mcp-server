import type {
  GraphSite,
  GraphDrive,
  GraphDriveItem,
  GraphSearchResponse,
  GraphSearchHit,
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

function getResourceTypeName(odataType?: string): string {
  if (!odataType) return "item";
  if (odataType.includes("driveItem")) return "driveItem";
  if (odataType.includes("site")) return "site";
  if (odataType.includes("listItem")) return "listItem";
  return "item";
}

function stripHtmlHighlights(text: string): string {
  return text.replace(/<c0>/g, "").replace(/<\/c0>/g, "").replace(/<ddd\/>/g, "…");
}

function formatHit(hit: GraphSearchHit): string {
  const resource = hit.resource;
  const typeName = getResourceTypeName(resource?.["@odata.type"]);
  const name = resource?.name ?? resource?.displayName ?? "Sem nome";
  const summary = hit.summary ? stripHtmlHighlights(hit.summary) : "N/A";
  const url = resource?.webUrl ?? "N/A";

  const lines: string[] = [
    `- **${name}** [${typeName}]`,
    `  Resumo: ${summary}`,
    `  URL: ${url}`,
  ];

  if (typeName === "driveItem") {
    const driveId = resource?.parentReference?.driveId;
    const itemId = hit.hitId;
    if (driveId) {
      lines.push(`  **driveId:** ${driveId}`);
      lines.push(`  **itemId:** ${itemId}`);
      lines.push(`  _Use driveId + itemId com read-shared-file-content para ler este arquivo._`);
    }
  } else if (typeName === "site") {
    const siteId = resource?.id;
    if (siteId) {
      lines.push(`  **siteId:** ${siteId}`);
      lines.push(`  _Use siteId com get-site ou list-document-libraries._`);
    }
  } else if (typeName === "listItem") {
    const siteId = resource?.parentReference?.siteId;
    if (siteId) {
      lines.push(`  **siteId:** ${siteId}`);
    }
  }

  return lines.join("\n");
}

export function formatSearchResults(
  query: string,
  results: GraphSearchResponse[]
): string {
  if (!results?.length || !results[0]?.hitsContainers?.length) {
    return `Nenhum resultado encontrado para "${query}".`;
  }

  const hits = results[0].hitsContainers[0].hits ?? [];
  const formatted = hits.map(formatHit);

  return `## Resultados SharePoint "${query}" (${formatted.length})\n\n${formatted.join("\n\n")}`;
}
