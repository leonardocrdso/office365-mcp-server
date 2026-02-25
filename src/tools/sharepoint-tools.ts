import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { SharePointService } from "../services/sharepoint.js";
import { safeTool } from "../utils/safe-tool.js";
import {
  formatSiteList,
  formatSiteDetail,
  formatDocumentLibraries,
  formatLibraryItems,
  formatSearchResults,
} from "../formatters/sharepoint.js";

export function registerSharePointTools(server: McpServer, sharepoint: SharePointService) {
  server.tool(
    "list-sites",
    "Lista sites do SharePoint. Opcionalmente filtra por nome.",
    {
      query: z.string().optional().describe("Termo de busca para filtrar sites"),
    },
    safeTool(async (params) => {
      const sites = await sharepoint.listSites(params.query);
      return {
        content: [{ type: "text" as const, text: formatSiteList(sites) }],
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
        content: [{ type: "text" as const, text: formatSiteDetail(site) }],
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
      return {
        content: [{ type: "text" as const, text: formatDocumentLibraries(drives) }],
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
      return {
        content: [{ type: "text" as const, text: formatLibraryItems(items) }],
      };
    })
  );

  server.tool(
    "search-sharepoint",
    "Busca conteúdo no SharePoint (documentos, listas, sites). Para arquivos (driveItem), retorna driveId e itemId que podem ser usados com read-shared-file-content.",
    {
      query: z.string().describe("Texto para buscar"),
    },
    safeTool(async (params) => {
      const results = await sharepoint.searchSharePoint(params.query);
      return {
        content: [{ type: "text" as const, text: formatSearchResults(params.query, results) }],
      };
    })
  );
}
