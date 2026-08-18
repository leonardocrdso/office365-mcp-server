import {
  PublicClientApplication,
  type Configuration,
  type DeviceCodeRequest,
  type AccountInfo,
  type SilentFlowRequest,
} from "@azure/msal-node";
import type { AuthProvider } from "../types/auth.js";
import { SCOPES } from "../constants.js";
import { AuthError, ErrorCode } from "../utils/errors.js";
import { loadConfig } from "./config-manager.js";
import { loadTokenCache, saveTokenCache, removeTokenCache } from "./token-cache-manager.js";

const ALL_SCOPES = ["User.Read", ...Object.values(SCOPES).flat()];

export class MsalClient implements AuthProvider {
  private pca: PublicClientApplication | null = null;
  private pcaPromise: Promise<PublicClientApplication> | null = null;
  private fileConfig: { clientId?: string; tenantId?: string } | null = null;
  private deviceCodeCallback: ((message: string) => void) | null = null;
  private pendingLogin: Promise<void> | null = null;
  private lastCacheSnapshot: string | null = null;

  setDeviceCodeCallback(callback: (message: string) => void) {
    this.deviceCodeCallback = callback;
  }

  resetPca(): void {
    this.pca = null;
    this.pcaPromise = null;
    this.fileConfig = null;
    this.lastCacheSnapshot = null;
  }

  private getConfig(): Configuration {
    const clientId = process.env.AZURE_CLIENT_ID ?? this.fileConfig?.clientId;
    if (!clientId) {
      throw new AuthError(
        "AZURE_CLIENT_ID não configurado.",
        ErrorCode.AUTH_NOT_CONFIGURED
      );
    }

    const tenantId = process.env.AZURE_TENANT_ID ?? this.fileConfig?.tenantId ?? "common";

    return {
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
      },
    };
  }

  private async getPca(): Promise<PublicClientApplication> {
    if (this.pca) return this.pca;
    this.pcaPromise ??= this.initPca().catch((error) => {
      this.pcaPromise = null;
      throw error;
    });
    return this.pcaPromise;
  }

  private async initPca(): Promise<PublicClientApplication> {
    if (!this.fileConfig) {
      this.fileConfig = await loadConfig();
    }
    const pca = new PublicClientApplication(this.getConfig());
    try {
      const cacheData = await loadTokenCache();
      if (cacheData) {
        pca.getTokenCache().deserialize(cacheData);
        this.lastCacheSnapshot = pca.getTokenCache().serialize();
      }
    } catch (error) {
      console.error("Failed to load token cache:", error);
    }
    this.pca = pca;
    return pca;
  }

  private async saveCache(): Promise<void> {
    if (!this.pca) return;
    const cacheData = this.pca.getTokenCache().serialize();
    if (cacheData === this.lastCacheSnapshot) return;
    this.lastCacheSnapshot = cacheData;
    await saveTokenCache(cacheData);
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
    await removeTokenCache();
    this.pca = null;
    this.pcaPromise = null;
    this.lastCacheSnapshot = null;
  }
}

export const msalClient = new MsalClient();
