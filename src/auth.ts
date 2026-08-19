import { mkdir } from "node:fs/promises";
import path from "node:path";
import {
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo,
} from "@azure/msal-node";
import {
  DataProtectionScope,
  PersistenceCachePlugin,
  PersistenceCreator,
} from "@azure/msal-node-extensions";
import type { GatewayConfig } from "./config.js";
import { COPILOT_DELEGATED_SCOPES } from "./types.js";

export class AuthenticationRequiredError extends Error {
  public constructor(message = "Sign in is required. Run npm run auth:login before starting the gateway.") {
    super(message);
    this.name = "AuthenticationRequiredError";
  }
}

export function isAuthenticationRequiredError(
  error: unknown,
): error is AuthenticationRequiredError | InteractionRequiredAuthError {
  return error instanceof AuthenticationRequiredError || error instanceof InteractionRequiredAuthError;
}

export interface AuthService {
  getAccessToken(): Promise<string>;
  loginInteractively(openBrowser: (url: string) => Promise<void>): Promise<{ username: string }>;
}

export async function ensureTokenCacheDirectory(directory: string): Promise<void> {
  await mkdir(directory, { recursive: true });
}

export function selectSingleAccount(accounts: AccountInfo[]): AccountInfo {
  const [account] = accounts;
  if (!account) {
    throw new AuthenticationRequiredError();
  }
  if (accounts.length > 1) {
    throw new AuthenticationRequiredError(
      "Multiple cached accounts were found. Run npm run auth:login to select an account.",
    );
  }
  return account;
}

export async function createAuthService(config: GatewayConfig): Promise<AuthService> {
  await ensureTokenCacheDirectory(config.tokenCacheDirectory);
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
    return selectSingleAccount(await application.getTokenCache().getAllAccounts());
  }

  return {
    async getAccessToken(): Promise<string> {
      try {
        const result = await application.acquireTokenSilent({
          account: await account(),
          scopes: [...COPILOT_DELEGATED_SCOPES],
        });
        if (!result?.accessToken) {
          throw new AuthenticationRequiredError();
        }
        return result.accessToken;
      } catch (error) {
        if (error instanceof InteractionRequiredAuthError) {
          throw new AuthenticationRequiredError();
        }
        throw error;
      }
    },

    async loginInteractively(openBrowser: (url: string) => Promise<void>): Promise<{ username: string }> {
      const result = await application.acquireTokenInteractive({
        scopes: [...COPILOT_DELEGATED_SCOPES],
        openBrowser,
        successTemplate:
          "Microsoft 365 sign-in completed. You can close this browser tab and return to the terminal.",
        errorTemplate:
          "Microsoft 365 sign-in failed. Return to the terminal for error details.",
      });
      if (!result?.account?.username) {
        throw new Error("Microsoft Entra sign-in completed without a user account.");
      }
      const cache = application.getTokenCache();
      for (const cachedAccount of await cache.getAllAccounts()) {
        if (cachedAccount.homeAccountId !== result.account.homeAccountId) {
          await cache.removeAccount(cachedAccount);
        }
      }
      return { username: result.account.username };
    },
  };
}
