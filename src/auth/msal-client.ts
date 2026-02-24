import {
  PublicClientApplication,
  type Configuration,
  type DeviceCodeRequest,
  type AccountInfo,
  type SilentFlowRequest,
} from "@azure/msal-node";
import { readFile, writeFile, unlink } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";
import type { AuthProvider } from "../types/auth.js";
import { SCOPES } from "../constants.js";
import { AuthError, ErrorCode } from "../utils/errors.js";

const TOKEN_CACHE_PATH = join(homedir(), ".office365-mcp-tokens.json");

const ALL_SCOPES = ["User.Read", ...Object.values(SCOPES).flat()];

export class MsalClient implements AuthProvider {
  private pca: PublicClientApplication | null = null;
  private deviceCodeCallback: ((message: string) => void) | null = null;
  private pendingLogin: Promise<void> | null = null;

  setDeviceCodeCallback(callback: (message: string) => void) {
    this.deviceCodeCallback = callback;
  }

  private getConfig(): Configuration {
    const clientId = process.env.AZURE_CLIENT_ID;
    if (!clientId) {
      throw new AuthError(
        "AZURE_CLIENT_ID não configurado.",
        ErrorCode.AUTH_NOT_CONFIGURED
      );
    }

    const tenantId = process.env.AZURE_TENANT_ID || "common";

    return {
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
    };
  }

  private async getPca(): Promise<PublicClientApplication> {
    if (!this.pca) {
      this.pca = new PublicClientApplication(this.getConfig());

      try {
        const cacheData = await readFile(TOKEN_CACHE_PATH, "utf-8");
        this.pca.getTokenCache().deserialize(cacheData);
      } catch (error) {
        console.error("Failed to load token cache:", error);
      }
    }
    return this.pca;
  }

  private async saveCache(): Promise<void> {
    if (!this.pca) return;
    const cacheData = this.pca.getTokenCache().serialize();
    await writeFile(TOKEN_CACHE_PATH, cacheData, { mode: 0o600 });
  }

  async login(): Promise<{ userCode: string; verificationUri: string; message: string }> {
    const pca = await this.getPca();

    return new Promise((resolve, reject) => {
      const request: DeviceCodeRequest = {
        scopes: ALL_SCOPES,
        deviceCodeCallback: (response) => {
          if (this.deviceCodeCallback) {
            this.deviceCodeCallback(response.message);
          }
          resolve({
            userCode: response.userCode,
            verificationUri: response.verificationUri,
            message: response.message,
          });
        },
      };

      this.pendingLogin = pca
        .acquireTokenByDeviceCode(request)
        .then(async (result) => {
          if (result) {
            console.error("[msal] Token acquired, saving cache...");
            await this.saveCache();
            console.error("[msal] Cache saved successfully");
          } else {
            console.error("[msal] acquireTokenByDeviceCode returned null");
          }
        })
        .catch((error) => {
          console.error("[msal] Device code flow error:", error.message);
        })
        .finally(() => {
          this.pendingLogin = null;
        });
    });
  }

  async getAccessToken(scopes?: string[]): Promise<string> {
    const pca = await this.getPca();
    const accounts = await pca.getTokenCache().getAllAccounts();

    if (accounts.length === 0) {
      throw new AuthError(
        "Nenhuma conta encontrada.",
        ErrorCode.AUTH_NOT_AUTHENTICATED
      );
    }

    const account = accounts[0] as AccountInfo;
    const silentRequest: SilentFlowRequest = {
      account,
      scopes: scopes ?? ALL_SCOPES,
    };

    try {
      const result = await pca.acquireTokenSilent(silentRequest);
      await this.saveCache();
      return result.accessToken;
    } catch (error) {
      console.error("Silent token acquisition failed:", error);
      throw new AuthError(
        "Falha ao renovar token silenciosamente.",
        ErrorCode.AUTH_TOKEN_REFRESH_FAILED
      );
    }
  }

  async getAccountInfo(): Promise<{ isLoggedIn: boolean; account?: AccountInfo }> {
    try {
      const pca = await this.getPca();
      const accounts = await pca.getTokenCache().getAllAccounts();
      if (accounts.length === 0) {
        return { isLoggedIn: false };
      }
      return { isLoggedIn: true, account: accounts[0] as AccountInfo };
    } catch (error) {
      console.error("Get account info failed:", error);
      return { isLoggedIn: false };
    }
  }

  async logout(): Promise<void> {
    if (this.pca) {
      const accounts = await this.pca.getTokenCache().getAllAccounts();
      for (const account of accounts) {
        await this.pca.getTokenCache().removeAccount(account as AccountInfo);
      }
    }
    try {
      await unlink(TOKEN_CACHE_PATH);
    } catch (error) {
      console.error("Failed to remove token cache file:", error);
    }
    this.pca = null;
  }
}

export const msalClient = new MsalClient();
