import path from "node:path";
import { PublicClientApplication, type AccountInfo } from "@azure/msal-node";
import {
  DataProtectionScope,
  PersistenceCachePlugin,
  PersistenceCreator,
} from "@azure/msal-node-extensions";
import type { GatewayConfig } from "./config.js";
import { COPILOT_DELEGATED_SCOPES } from "./types.js";

export class AuthenticationRequiredError extends Error {
  public constructor() {
    super("Sign in is required. Run npm run auth:login before starting the gateway.");
    this.name = "AuthenticationRequiredError";
  }
}

export interface AuthService {
  getAccessToken(): Promise<string>;
  loginWithDeviceCode(onMessage: (message: string) => void): Promise<{ username: string }>;
}

export async function createAuthService(config: GatewayConfig): Promise<AuthService> {
  const cachePath = path.join(config.tokenCacheDirectory, "msal-cache.json");
  const persistence = await PersistenceCreator.createPersistence({
    cachePath,
    dataProtectionScope: DataProtectionScope.CurrentUser,
    serviceName: "m365-copilot-openai-gateway",
    accountName: `${config.tenantId}.${config.clientId}`,
  });

  const application = new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
    },
    cache: { cachePlugin: new PersistenceCachePlugin(persistence) },
  });

  async function account(): Promise<AccountInfo> {
    const accounts = await application.getTokenCache().getAllAccounts();
    const firstAccount = accounts[0];
    if (!firstAccount) {
      throw new AuthenticationRequiredError();
    }
    return firstAccount;
  }

  return {
    async getAccessToken(): Promise<string> {
      const result = await application.acquireTokenSilent({
        account: await account(),
        scopes: [...COPILOT_DELEGATED_SCOPES],
      });
      if (!result?.accessToken) {
        throw new AuthenticationRequiredError();
      }
      return result.accessToken;
    },

    async loginWithDeviceCode(onMessage: (message: string) => void): Promise<{ username: string }> {
      const result = await application.acquireTokenByDeviceCode({
        scopes: [...COPILOT_DELEGATED_SCOPES],
        deviceCodeCallback(response) {
          onMessage(response.message);
        },
      });
      if (!result?.account?.username) {
        throw new Error("Microsoft Entra sign-in completed without a user account.");
      }
      return { username: result.account.username };
    },
  };
}
