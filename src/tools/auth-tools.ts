import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { MsalClient } from "../auth/msal-client.js";
import { saveConfig } from "../auth/config-manager.js";
import { resolveStoragePaths } from "../auth/storage.js";
import { safeTool } from "../utils/safe-tool.js";

export function registerAuthTools(server: McpServer, auth: MsalClient) {
  server.tool(
    "configure",
    "Configura as credenciais do Azure (Client ID e Tenant ID) para autenticação com Microsoft 365. Salva em arquivo persistente no storage configurado.",
    {
      clientId: z.string().describe("Azure App Registration Client ID"),
      tenantId: z.string().optional().default("common").describe("Azure Tenant ID (padrão: 'common')"),
    },
    safeTool(async (params) => {
      const savedPath = await saveConfig(params.clientId, params.tenantId);
      auth.resetPca();
      return {
        content: [
          {
            type: "text" as const,
            text: [
              "## Configuração salva com sucesso",
              "",
              `**Client ID:** \`${params.clientId}\``,
              `**Tenant ID:** \`${params.tenantId}\``,
              "",
              `As credenciais foram salvas em \`${savedPath}\`.`,
              "Agora use a tool **'login'** para autenticar com sua conta Microsoft.",
            ].join("\n"),
          },
        ],
      };
    })
  );

  server.tool(
    "login",
    "Inicia autenticação com Microsoft 365 via Device Code Flow. Retorna um código e URL para o usuário autenticar no navegador.",
    {},
    safeTool(async () => {
      const result = await auth.login();
      return {
        content: [
          {
            type: "text" as const,
            text: [
              "## Autenticação Microsoft 365",
              "",
              `**Código:** \`${result.userCode}\``,
              `**URL:** ${result.verificationUri}`,
              "",
              "Abra a URL acima no navegador, insira o código e faça login com sua conta Microsoft.",
              "Após completar o login no navegador, o token será salvo automaticamente.",
              "",
              "Use a tool 'auth-status' para verificar se a autenticação foi concluída.",
            ].join("\n"),
          },
        ],
      };
    })
  );

  server.tool(
    "auth-status",
    "Verifica o status da autenticação e mostra informações do usuário logado.",
    {},
    safeTool(async () => {
      const info = await auth.getAccountInfo();

      if (!info.isLoggedIn) {
        return {
          content: [
            {
              type: "text" as const,
              text: "Não autenticado. Use a tool 'login' para iniciar a autenticação.",
            },
          ],
        };
      }

      const account = info.account!;
      const { baseDir, configPath } = resolveStoragePaths();
      const storageInfo = baseDir
        ? `\`${baseDir}\` (isolado via OFFICE365_MCP_HOME)`
        : `\`${configPath}\` (storage padrão)`;
      return {
        content: [
          {
            type: "text" as const,
            text: [
              "## Status: Autenticado",
              "",
              `**Nome:** ${account.name ?? "N/A"}`,
              `**Email:** ${account.username}`,
              `**Tenant:** ${account.tenantId}`,
              `**Storage:** ${storageInfo}`,
            ].join("\n"),
          },
        ],
      };
    })
  );

  server.tool(
    "logout",
    "Faz logout e remove tokens salvos.",
    {},
    safeTool(async () => {
      await auth.logout();
      return {
        content: [
          {
            type: "text" as const,
            text: "Logout realizado com sucesso. Tokens removidos.",
          },
        ],
      };
    })
  );
}
