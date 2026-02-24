export const ErrorCode = {
  // Auth
  AUTH_NOT_CONFIGURED: "AUTH_NOT_CONFIGURED",
  AUTH_NOT_AUTHENTICATED: "AUTH_NOT_AUTHENTICATED",
  AUTH_TOKEN_EXPIRED: "AUTH_TOKEN_EXPIRED",
  AUTH_TOKEN_REFRESH_FAILED: "AUTH_TOKEN_REFRESH_FAILED",

  // Permissions
  PERMISSION_DENIED: "PERMISSION_DENIED",
  INSUFFICIENT_SCOPES: "INSUFFICIENT_SCOPES",

  // Resource
  RESOURCE_NOT_FOUND: "RESOURCE_NOT_FOUND",
  RESOURCE_CONFLICT: "RESOURCE_CONFLICT",
  RESOURCE_GONE: "RESOURCE_GONE",

  // Request
  BAD_REQUEST: "BAD_REQUEST",
  VALIDATION_ERROR: "VALIDATION_ERROR",
  PAYLOAD_TOO_LARGE: "PAYLOAD_TOO_LARGE",
  RATE_LIMITED: "RATE_LIMITED",

  // Server
  GRAPH_SERVER_ERROR: "GRAPH_SERVER_ERROR",
  GRAPH_UNAVAILABLE: "GRAPH_UNAVAILABLE",
  GRAPH_TIMEOUT: "GRAPH_TIMEOUT",

  // Network
  NETWORK_ERROR: "NETWORK_ERROR",

  // Generic
  UNKNOWN_ERROR: "UNKNOWN_ERROR",
} as const;

export type ErrorCode = (typeof ErrorCode)[keyof typeof ErrorCode];

export class GraphApiError extends Error {
  public readonly code: ErrorCode;

  constructor(
    message: string,
    public statusCode: number,
    public graphCode?: string
  ) {
    super(message);
    this.name = "GraphApiError";
    this.code = resolveErrorCode(statusCode, graphCode);
  }
}

export class AuthError extends Error {
  constructor(
    message: string,
    public readonly code: ErrorCode
  ) {
    super(message);
    this.name = "AuthError";
  }
}

function resolveErrorCode(statusCode: number, graphCode?: string): ErrorCode {
  if (graphCode) {
    const mapped = GRAPH_CODE_MAP[graphCode];
    if (mapped) return mapped;
  }

  switch (statusCode) {
    case 400: return ErrorCode.BAD_REQUEST;
    case 401: return ErrorCode.AUTH_TOKEN_EXPIRED;
    case 403: return ErrorCode.PERMISSION_DENIED;
    case 404: return ErrorCode.RESOURCE_NOT_FOUND;
    case 409: return ErrorCode.RESOURCE_CONFLICT;
    case 410: return ErrorCode.RESOURCE_GONE;
    case 413: return ErrorCode.PAYLOAD_TOO_LARGE;
    case 429: return ErrorCode.RATE_LIMITED;
    case 500: return ErrorCode.GRAPH_SERVER_ERROR;
    case 502: return ErrorCode.GRAPH_UNAVAILABLE;
    case 503: return ErrorCode.GRAPH_UNAVAILABLE;
    case 504: return ErrorCode.GRAPH_TIMEOUT;
    default: return ErrorCode.UNKNOWN_ERROR;
  }
}

const GRAPH_CODE_MAP: Record<string, ErrorCode> = {
  InvalidAuthenticationToken: ErrorCode.AUTH_TOKEN_EXPIRED,
  AuthenticationError: ErrorCode.AUTH_TOKEN_EXPIRED,
  AccessDenied: ErrorCode.PERMISSION_DENIED,
  ErrorAccessDenied: ErrorCode.PERMISSION_DENIED,
  Authorization_RequestDenied: ErrorCode.INSUFFICIENT_SCOPES,
  ErrorItemNotFound: ErrorCode.RESOURCE_NOT_FOUND,
  ResourceNotFound: ErrorCode.RESOURCE_NOT_FOUND,
  ErrorInvalidIdMalformed: ErrorCode.BAD_REQUEST,
  BadRequest: ErrorCode.BAD_REQUEST,
  ErrorInvalidRequest: ErrorCode.BAD_REQUEST,
  RequestBodyRead: ErrorCode.VALIDATION_ERROR,
  ValidationError: ErrorCode.VALIDATION_ERROR,
  ErrorMessageSizeExceeded: ErrorCode.PAYLOAD_TOO_LARGE,
  MaximumFileSize: ErrorCode.PAYLOAD_TOO_LARGE,
  activityLimitReached: ErrorCode.RATE_LIMITED,
  ErrorMailboxNotEnabledForRESTAPI: ErrorCode.PERMISSION_DENIED,
  ErrorMailboxMoveInProgress: ErrorCode.GRAPH_UNAVAILABLE,
  ErrorSendAsDenied: ErrorCode.PERMISSION_DENIED,
};

const ERROR_MESSAGES: Record<ErrorCode, string> = {
  [ErrorCode.AUTH_NOT_CONFIGURED]:
    "AZURE_CLIENT_ID não configurado. Defina a variável de ambiente AZURE_CLIENT_ID com o ID do seu app registration no Azure AD.",
  [ErrorCode.AUTH_NOT_AUTHENTICATED]:
    "Usuário não autenticado. Use a tool 'login' para iniciar a autenticação.",
  [ErrorCode.AUTH_TOKEN_EXPIRED]:
    "Token expirado ou inválido. Use a tool 'login' para re-autenticar.",
  [ErrorCode.AUTH_TOKEN_REFRESH_FAILED]:
    "Falha ao renovar token silenciosamente. Use a tool 'login' para re-autenticar.",
  [ErrorCode.PERMISSION_DENIED]:
    "Sem permissão para esta operação. Verifique se as permissões (API permissions) estão configuradas e consentidas no Azure AD.",
  [ErrorCode.INSUFFICIENT_SCOPES]:
    "Permissões insuficientes. O app precisa de permissões adicionais no Azure AD. Faça logout e login novamente após adicionar as permissões.",
  [ErrorCode.RESOURCE_NOT_FOUND]:
    "Recurso não encontrado. Verifique se o ID fornecido está correto e se o recurso ainda existe.",
  [ErrorCode.RESOURCE_CONFLICT]:
    "Conflito: o recurso foi modificado por outro processo. Tente novamente.",
  [ErrorCode.RESOURCE_GONE]:
    "O recurso foi removido permanentemente e não está mais disponível.",
  [ErrorCode.BAD_REQUEST]:
    "Requisição inválida. Verifique os parâmetros enviados.",
  [ErrorCode.VALIDATION_ERROR]:
    "Erro de validação nos dados enviados. Verifique o formato dos campos.",
  [ErrorCode.PAYLOAD_TOO_LARGE]:
    "Conteúdo muito grande. O limite para upload via Graph API é 4MB.",
  [ErrorCode.RATE_LIMITED]:
    "Rate limit atingido. A Microsoft Graph API está limitando requisições. Aguarde alguns segundos e tente novamente.",
  [ErrorCode.GRAPH_SERVER_ERROR]:
    "Erro interno da Microsoft Graph API. Tente novamente em alguns instantes.",
  [ErrorCode.GRAPH_UNAVAILABLE]:
    "Microsoft Graph API temporariamente indisponível. Tente novamente em alguns instantes.",
  [ErrorCode.GRAPH_TIMEOUT]:
    "Timeout na Microsoft Graph API. A operação pode ter sido concluída no servidor. Verifique antes de tentar novamente.",
  [ErrorCode.NETWORK_ERROR]:
    "Erro de rede ao conectar com a Microsoft Graph API. Verifique sua conexão com a internet.",
  [ErrorCode.UNKNOWN_ERROR]:
    "Erro inesperado. Veja os detalhes abaixo.",
};

export function formatGraphError(error: unknown): string {
  if (error instanceof GraphApiError) {
    const baseMessage = ERROR_MESSAGES[error.code];
    const detail = error.message !== error.code ? error.message : "";
    const parts = [`[${error.code}] ${baseMessage}`];
    if (detail) parts.push(`Detalhe: ${detail}`);
    if (error.graphCode) parts.push(`Graph code: ${error.graphCode}`);
    return parts.join("\n");
  }

  if (error instanceof AuthError) {
    return `[${error.code}] ${ERROR_MESSAGES[error.code]}`;
  }

  if (error instanceof TypeError && error.message.includes("fetch")) {
    return `[${ErrorCode.NETWORK_ERROR}] ${ERROR_MESSAGES[ErrorCode.NETWORK_ERROR]}\nDetalhe: ${error.message}`;
  }

  if (error instanceof Error) {
    return `[${ErrorCode.UNKNOWN_ERROR}] ${error.message}`;
  }

  return `[${ErrorCode.UNKNOWN_ERROR}] ${String(error)}`;
}
