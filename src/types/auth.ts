export interface AuthProvider {
  getAccessToken(scopes: string[]): Promise<string>;
}
