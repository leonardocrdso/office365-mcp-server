import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { OneDriveService } from "../services/onedrive.js";
import { safeTool } from "../utils/errors.js";
import {
  formatDriveItems,
  formatFileSearchResults,
  formatUploadResult,
  formatShareLink,
} from "../formatters/onedrive.js";

export function registerOneDriveTools(server: McpServer, onedrive: OneDriveService) {
  server.tool(
    "list-drive-files",
    "Lista arquivos e pastas do OneDrive. Por padrão lista a raiz.",
    {
      itemId: z.string().optional().describe("ID do item/pasta para listar conteúdo"),
      path: z.string().optional().describe("Caminho da pasta (ex: 'Documents/Projetos')"),
      top: z.number().optional().describe("Número máximo de itens (padrão: 20)"),
    },
    safeTool(async (params) => {
      const items = await onedrive.listFiles(params);
      return {
        content: [{ type: "text" as const, text: formatDriveItems(items) }],
      };
    })
  );

  server.tool(
    "read-file-content",
    "Lê o conteúdo de um arquivo de texto do OneDrive.",
    {
      itemId: z.string().describe("ID do arquivo no OneDrive"),
    },
    safeTool(async (params) => {
      const content = await onedrive.readFileContent(params.itemId);
      return {
        content: [
          {
            type: "text" as const,
            text: typeof content === "string" ? content : JSON.stringify(content, null, 2),
          },
        ],
      };
    })
  );

  server.tool(
    "upload-file",
    "Faz upload de um arquivo para o OneDrive (até 4MB de conteúdo texto).",
    {
      path: z.string().describe("Caminho de destino no OneDrive (ex: 'Documents/relatorio.txt')"),
      content: z.string().describe("Conteúdo do arquivo"),
    },
    safeTool(async (params) => {
      const result = await onedrive.uploadFile(params);
      return {
        content: [{ type: "text" as const, text: formatUploadResult(result) }],
      };
    })
  );

  server.tool(
    "search-files",
    "Busca arquivos no OneDrive por texto.",
    {
      query: z.string().describe("Texto para buscar"),
      top: z.number().optional().describe("Número máximo de resultados (padrão: 10)"),
    },
    safeTool(async (params) => {
      const items = await onedrive.searchFiles(params.query, params.top);
      return {
        content: [{ type: "text" as const, text: formatFileSearchResults(params.query, items) }],
      };
    })
  );

  server.tool(
    "share-file",
    "Cria um link de compartilhamento para um arquivo do OneDrive.",
    {
      itemId: z.string().describe("ID do arquivo"),
      type: z.enum(["view", "edit"]).optional().describe("Tipo de acesso (padrão: view)"),
      scope: z.enum(["anonymous", "organization"]).optional().describe("Escopo (padrão: organization)"),
    },
    safeTool(async (params) => {
      const result = await onedrive.shareFile(params);
      return {
        content: [{ type: "text" as const, text: formatShareLink(result) }],
      };
    })
  );
}
