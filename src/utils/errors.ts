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

export function formatGraphError(error: unknown): string {
  if (error instanceof GraphApiError) {
    switch (error.statusCode) {
      case 401:
        return "Token expirado ou inválido. Use a tool 'login' para re-autenticar.";
      case 403:
        return `Sem permissão: ${error.message}. Verifique as permissões do app no Azure AD.`;
      case 404:
        return `Recurso não encontrado: ${error.message}`;
      case 429:
        return "Rate limit atingido. Tente novamente em alguns segundos.";
      default:
        return `Erro Graph API (${error.statusCode}): ${error.message}`;
    }
  }

  if (error instanceof Error) {
    if (error.message.includes("No account found") || error.message.includes("No token")) {
      return "Usuário não autenticado. Use a tool 'login' primeiro.";
    }
    return error.message;
  }

  return String(error);
}

type ToolResult = { content: { type: "text"; text: string }[]; isError?: boolean };
type ToolHandler<T extends unknown[] = unknown[]> = (...args: T) => Promise<ToolResult>;

export function safeTool<T extends unknown[]>(handler: ToolHandler<T>): ToolHandler<T> {
  return (async (...args: T) => {
    try {
      return await handler(...args);
    } catch (error) {
      const message = formatGraphError(error);
      return {
        content: [{ type: "text" as const, text: message }],
        isError: true,
      };
    }
  }) as ToolHandler<T>;
}

export async function graphFetch<T = unknown>(
  accessToken: string,
  url: string,
  options: RequestInit = {}
): Promise<T> {
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
    } catch (parseError) {
      console.error("Failed to parse Graph API error response:", parseError);
    }
    throw new GraphApiError(errorMessage, response.status, graphCode);
  }

  if (response.status === 204) return null as T;

  const contentType = response.headers.get("content-type") ?? "";
  if (contentType.includes("application/json")) {
    return response.json() as Promise<T>;
  }
  return response.text() as T;
}
