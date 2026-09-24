import {
  GRAPH_ATTEMPT_TIMEOUT_MS,
  GRAPH_MAX_ATTEMPTS,
  GRAPH_MIN_ATTEMPT_BUDGET_MS,
  GRAPH_RETRY_BASE_DELAY_MS,
  GRAPH_RETRY_MAX_DELAY_MS,
  GRAPH_TOTAL_BUDGET_MS,
} from "../constants.js";
import { GraphApiError } from "./errors.js";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const HTTP_TOO_MANY_REQUESTS = 429;
const HTTP_GATEWAY_TIMEOUT = 504;
const RETRIABLE_STATUS_CODES: ReadonlySet<number> = new Set([429, 500, 502, 503, 504]);
const IDEMPOTENT_METHODS: ReadonlySet<string> = new Set(["GET", "HEAD"]);

function toAbsoluteUrl(url: string): string {
  return url.startsWith("https://") ? url : `${GRAPH_BASE_URL}${url}`;
}

function isTimeoutAbort(error: unknown): boolean {
  return error instanceof Error && (error.name === "TimeoutError" || error.name === "AbortError");
}

function canRetry(method: string, statusCode?: number): boolean {
  if (statusCode === HTTP_TOO_MANY_REQUESTS) return true;
  return IDEMPOTENT_METHODS.has(method.toUpperCase());
}

function backoffDelayMs(attempt: number): number {
  const exponential = GRAPH_RETRY_BASE_DELAY_MS * 2 ** (attempt - 1);
  const capped = Math.min(exponential, GRAPH_RETRY_MAX_DELAY_MS);
  return capped / 2 + Math.random() * (capped / 2);
}

function retryAfterMs(response: Response): number | undefined {
  const header = response.headers.get("retry-after");
  if (!header) return undefined;
  const seconds = Number(header);
  if (!Number.isFinite(seconds) || seconds < 0) return undefined;
  return Math.min(seconds * 1_000, GRAPH_RETRY_MAX_DELAY_MS);
}

function clampAttemptTimeout(remainingMs: number): number {
  return Math.max(GRAPH_MIN_ATTEMPT_BUDGET_MS, Math.min(GRAPH_ATTEMPT_TIMEOUT_MS, remainingMs));
}

function hasSpentBudget(attempt: number, deadline: number): boolean {
  return attempt >= GRAPH_MAX_ATTEMPTS || deadline - Date.now() < GRAPH_MIN_ATTEMPT_BUDGET_MS;
}

function sleep(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

function logRetry(method: string, url: string, attempt: number, delay: number, reason: string): void {
  const { pathname } = new URL(url);
  console.error(
    JSON.stringify({
      event: "graph_request_retry",
      method,
      path: pathname,
      attempt,
      maxAttempts: GRAPH_MAX_ATTEMPTS,
      delayMs: Math.round(delay),
      reason,
    })
  );
}

async function toGraphApiError(response: Response): Promise<GraphApiError> {
  let errorMessage = response.statusText;
  let graphCode: string | undefined;
  try {
    const errorBody = await response.json();
    errorMessage = errorBody?.error?.message ?? errorMessage;
    graphCode = errorBody?.error?.code;
  } catch (parseError) {
    console.error("Failed to parse Graph API error response:", parseError);
  }
  return new GraphApiError(errorMessage, response.status, graphCode);
}

async function fetchOnce(
  accessToken: string,
  url: string,
  options: RequestInit,
  timeoutMs: number
): Promise<Response> {
  try {
    return await fetch(url, {
      ...options,
      signal: AbortSignal.timeout(timeoutMs),
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        ...options.headers,
      },
    });
  } catch (networkError) {
    if (isTimeoutAbort(networkError)) {
      throw new GraphApiError(
        `A requisição ficou ${Math.round(timeoutMs / 1_000)}s sem resposta e foi cancelada pelo cliente.`,
        HTTP_GATEWAY_TIMEOUT
      );
    }
    throw new TypeError(
      `fetch failed: não foi possível conectar a ${new URL(url).hostname}. ${networkError instanceof Error ? networkError.message : String(networkError)}`
    );
  }
}

async function executeRequest(
  accessToken: string,
  url: string,
  options: RequestInit = {}
): Promise<Response> {
  const absoluteUrl = toAbsoluteUrl(url);
  const method = options.method ?? "GET";
  const deadline = Date.now() + GRAPH_TOTAL_BUDGET_MS;

  for (let attempt = 1; ; attempt += 1) {
    const attemptTimeout = clampAttemptTimeout(deadline - Date.now());

    let response: Response;
    try {
      response = await fetchOnce(accessToken, absoluteUrl, options, attemptTimeout);
    } catch (requestError) {
      const isLastAttempt = hasSpentBudget(attempt, deadline);
      const statusCode = requestError instanceof GraphApiError ? requestError.statusCode : undefined;
      if (isLastAttempt || !canRetry(method, statusCode)) throw requestError;
      const delay = backoffDelayMs(attempt);
      logRetry(method, absoluteUrl, attempt, delay, statusCode ? `http_${statusCode}` : "network_error");
      await sleep(delay);
      continue;
    }

    if (response.ok) return response;

    const isRetriable = RETRIABLE_STATUS_CODES.has(response.status) && canRetry(method, response.status);
    if (hasSpentBudget(attempt, deadline) || !isRetriable) throw await toGraphApiError(response);

    const delay = retryAfterMs(response) ?? backoffDelayMs(attempt);
    logRetry(method, absoluteUrl, attempt, delay, `http_${response.status}`);
    await sleep(delay);
  }
}

export async function graphFetch<T>(
  accessToken: string,
  url: string,
  options: RequestInit = {}
): Promise<T> {
  const response = await executeRequest(accessToken, url, options);

  const contentType = response.headers.get("content-type") ?? "";
  if (contentType.includes("application/json")) {
    return response.json() as Promise<T>;
  }
  return response.text() as unknown as T;
}

export async function graphFetchBinary(
  accessToken: string,
  url: string,
  options: RequestInit = {}
): Promise<Uint8Array> {
  const response = await executeRequest(accessToken, url, options);
  return new Uint8Array(await response.arrayBuffer());
}

export async function graphFetchVoid(
  accessToken: string,
  url: string,
  options: RequestInit = {}
): Promise<void> {
  await executeRequest(accessToken, url, options);
}
