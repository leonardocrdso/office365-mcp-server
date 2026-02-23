import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import * as onedrive from "../services/onedrive.ts";
import { safeTool } from "../utils/errors.ts";

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 * 1024 * 1024) return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  return `${(bytes / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

export function registerOneDriveTools(server: McpServer) {
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
      const formatted = items.map((item: any) => {
        const isFolder = !!item.folder;
        const icon = isFolder ? "[pasta]" : "[arquivo]";
        const size = item.size ? ` (${formatSize(item.size)})` : "";
        const modified = new Date(item.lastModifiedDateTime).toLocaleString("pt-BR");
        const childCount = item.folder?.childCount != null ? ` — ${item.folder.childCount} itens` : "";
        return `- ${icon} **${item.name}**${size}${childCount}\n  Modificado: ${modified}\n  ID: ${item.id}`;
      });

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? `## OneDrive (${formatted.length} itens)\n\n${formatted.join("\n\n")}`
              : "Pasta vazia.",
          },
        ],
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
        content: [
          {
            type: "text" as const,
            text: `Arquivo enviado com sucesso!\n\n**${result.name}** (${formatSize(result.size)})\nURL: ${result.webUrl}\nID: ${result.id}`,
          },
        ],
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
      const formatted = items.map((item: any) => {
        const size = item.size ? ` (${formatSize(item.size)})` : "";
        const path = item.parentReference?.path ?? "";
        return `- **${item.name}**${size}\n  Caminho: ${path}\n  URL: ${item.webUrl}\n  ID: ${item.id}`;
      });

      return {
        content: [
          {
            type: "text" as const,
            text: formatted.length > 0
              ? `## Resultados "${params.query}" (${formatted.length})\n\n${formatted.join("\n\n")}`
              : `Nenhum arquivo encontrado para "${params.query}".`,
          },
        ],
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
      const link = result.link;
      return {
        content: [
          {
            type: "text" as const,
            text: `Link criado!\n\n**URL:** ${link?.webUrl}\n**Tipo:** ${link?.type}\n**Escopo:** ${link?.scope}`,
          },
        ],
      };
    })
  );
}
