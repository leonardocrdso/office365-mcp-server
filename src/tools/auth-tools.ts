import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { msalClient } from "../auth/msal-client.js";
import { safeTool } from "../utils/errors.js";

export function registerAuthTools(server: McpServer) {
  server.tool(
    "login",
    "Inicia autenticação com Microsoft 365 via Device Code Flow. Retorna um código e URL para o usuário autenticar no navegador.",
    {},
    safeTool(async () => {
      const result = await msalClient.login();
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
      const info = await msalClient.getAccountInfo();

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
      await msalClient.logout();
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
