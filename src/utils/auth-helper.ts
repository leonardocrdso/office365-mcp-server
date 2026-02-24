import type { AuthProvider } from "../types/auth.js";

export function createGetToken(auth: AuthProvider, scopes: readonly string[]): () => Promise<string> {
  return () => auth.getAccessToken([...scopes]);
}
