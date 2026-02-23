import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

export class GraphApiError extends Error {
  constructor(
    message: string,
    public statusCode: number,
    public graphCode?: string
  ) {
    super(message);
    this.name = "GraphApiError";
  }
}

function formatGraphError(error: unknown): { message: string; isAuthError: boolean } {
  if (error instanceof GraphApiError) {
    switch (error.statusCode) {
      case 401:
        return { message: "Token expirado ou inválido. Use a tool 'login' para re-autenticar.", isAuthError: true };
      case 403:
        return { message: `Sem permissão: ${error.message}. Verifique as permissões do app no Azure AD.`, isAuthError: false };
      case 404:
        return { message: `Recurso não encontrado: ${error.message}`, isAuthError: false };
      case 429:
        return { message: "Rate limit atingido. Tente novamente em alguns segundos.", isAuthError: false };
      default:
        return { message: `Erro Graph API (${error.statusCode}): ${error.message}`, isAuthError: false };
    }
  }

  if (error instanceof Error) {
    if (error.message.includes("No account found") || error.message.includes("No token")) {
      return { message: "Usuário não autenticado. Use a tool 'login' primeiro.", isAuthError: true };
    }
    return { message: error.message, isAuthError: false };
  }

  return { message: String(error), isAuthError: false };
}

type ToolHandler = (...args: any[]) => Promise<{ content: { type: "text"; text: string }[]; isError?: boolean }>;

export function safeTool(handler: ToolHandler): ToolHandler {
  return async (...args: any[]) => {
    try {
      return await handler(...args);
    } catch (error) {
      const { message } = formatGraphError(error);
      return {
        content: [{ type: "text" as const, text: message }],
        isError: true,
      };
    }
  };
}

export async function graphFetch(
  accessToken: string,
  url: string,
  options: RequestInit = {}
): Promise<any> {
  const baseUrl = url.startsWith("https://") ? url : `https://graph.microsoft.com/v1.0${url}`;

  const response = await fetch(baseUrl, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
      ...options.headers,
    },
  });

  if (!response.ok) {
    let errorMessage = response.statusText;
    let graphCode: string | undefined;
    try {
      const errorBody = await response.json();
      errorMessage = errorBody?.error?.message ?? errorMessage;
      graphCode = errorBody?.error?.code;
    } catch {}
    throw new GraphApiError(errorMessage, response.status, graphCode);
  }

  if (response.status === 204) return null;

  const contentType = response.headers.get("content-type") ?? "";
  if (contentType.includes("application/json")) {
    return response.json();
  }
  return response.text();
}
