import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import * as sharepoint from "../services/sharepoint.ts";
import { safeTool } from "../utils/errors.ts";

export function registerSharePointTools(server: McpServer) {
  server.tool(
    "list-sites",
    "Lista sites do SharePoint. Opcionalmente filtra por nome.",
    {
      query: z.string().optional().describe("Termo de busca para filtrar sites"),
    },
    safeTool(async (params) => {
      const sites = await sharepoint.listSites(params.query);
      const formatted = sites.map(
        (s: any) =>
          `- **${s.displayName}** (${s.name})\n  ${s.description ?? ""}\n  URL: ${s.webUrl}\n  ID: ${s.id}`
      );

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? `## Sites SharePoint (${formatted.length})\n\n${formatted.join("\n\n")}`
              : "Nenhum site encontrado.",
          },
        ],
      };
    })
  );

  server.tool(
    "get-site",
    "Mostra detalhes de um site SharePoint específico.",
    {
      siteId: z.string().describe("ID do site SharePoint"),
    },
    safeTool(async (params) => {
      const site = await sharepoint.getSite(params.siteId);
      return {
        content: [
          {
            type: "text" as const,
            text: [
              `## ${site.displayName}`,
              "",
              `**Nome:** ${site.name}`,
              `**Descrição:** ${site.description ?? "N/A"}`,
              `**URL:** ${site.webUrl}`,
              `**ID:** ${site.id}`,
            ].join("\n"),
          },
        ],
      };
    })
  );

  server.tool(
    "list-document-libraries",
    "Lista bibliotecas de documentos de um site SharePoint.",
    {
      siteId: z.string().describe("ID do site SharePoint"),
    },
    safeTool(async (params) => {
      const drives = await sharepoint.listDocumentLibraries(params.siteId);
      const formatted = drives.map(
        (d: any) =>
          `- **${d.name}** (${d.driveType})\n  ${d.description ?? ""}\n  URL: ${d.webUrl}\n  ID: ${d.id}`
      );

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? `## Bibliotecas de Documentos (${formatted.length})\n\n${formatted.join("\n\n")}`
              : "Nenhuma biblioteca encontrada.",
          },
        ],
      };
    })
  );

  server.tool(
    "list-library-items",
    "Lista itens de uma biblioteca de documentos do SharePoint.",
    {
      driveId: z.string().describe("ID da biblioteca (drive)"),
      itemId: z.string().optional().describe("ID do item/pasta para listar conteúdo"),
      top: z.number().optional().describe("Número máximo de itens (padrão: 20)"),
    },
    safeTool(async (params) => {
      const items = await sharepoint.listLibraryItems(params);
      const formatted = items.map((item: any) => {
        const isFolder = !!item.folder;
        const icon = isFolder ? "[pasta]" : "[arquivo]";
        const size = item.size ? ` (${item.size} bytes)` : "";
        return `- ${icon} **${item.name}**${size}\n  URL: ${item.webUrl}\n  ID: ${item.id}`;
      });

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? `## Itens (${formatted.length})\n\n${formatted.join("\n\n")}`
              : "Nenhum item encontrado.",
          },
        ],
      };
    })
  );

  server.tool(
    "search-sharepoint",
    "Busca conteúdo no SharePoint (documentos, listas, sites).",
    {
      query: z.string().describe("Texto para buscar"),
    },
    safeTool(async (params) => {
      const results = await sharepoint.searchSharePoint(params.query);

      if (!results?.length || !results[0]?.hitsContainers?.length) {
        return {
          content: [
            {
              type: "text" as const,
              text: `Nenhum resultado encontrado para "${params.query}".`,
            },
          ],
        };
      }

      const hits = results[0].hitsContainers[0].hits ?? [];
      const formatted = hits.map((hit: any) => {
        const resource = hit.resource;
        return `- **${resource?.name ?? "Sem nome"}** (${hit.hitId})\n  Resumo: ${hit.summary ?? "N/A"}\n  URL: ${resource?.webUrl ?? "N/A"}`;
      });

      return {
        content: [
          {
            type: "text" as const,
            text: `## Resultados SharePoint "${params.query}" (${formatted.length})\n\n${formatted.join("\n\n")}`,
          },
        ],
      };
    })
  );
}
