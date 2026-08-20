import { GraphApiError } from "./errors.js";

async function executeRequest(
  accessToken: string,
  url: string,
  options: RequestInit = {}
): Promise<Response> {
  const baseUrl = url.startsWith("https://") ? url : `https://graph.microsoft.com/v1.0${url}`;

  let response: Response;
  try {
    response = await fetch(baseUrl, {
      ...options,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        ...options.headers,
      },
    });
  } catch (networkError) {
    throw new TypeError(
      `fetch failed: não foi possível conectar a ${new URL(baseUrl).hostname}. ${networkError instanceof Error ? networkError.message : String(networkError)}`
    );
  }

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

  return response;
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
