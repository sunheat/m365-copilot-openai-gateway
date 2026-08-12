import { existsSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import type { AccountInfo } from "@azure/msal-node";
import { AuthenticationRequiredError, ensureTokenCacheDirectory, selectSingleAccount } from "../src/auth.js";

function account(homeAccountId: string): AccountInfo {
  return { homeAccountId } as AccountInfo;
}

describe("ensureTokenCacheDirectory", () => {
  it("creates a missing cache directory and its parents", async () => {
    const parent = await mkdtemp(path.join(tmpdir(), "m365-copilot-gateway-"));
    const cacheDirectory = path.join(parent, "nested", "msal-cache");

    try {
      await ensureTokenCacheDirectory(cacheDirectory);
      expect(existsSync(cacheDirectory)).toBe(true);
    } finally {
      await rm(parent, { recursive: true, force: true });
    }
  });
});

describe("selectSingleAccount", () => {
  it("rejects an ambiguous token cache instead of choosing an arbitrary account", () => {
    expect(() => selectSingleAccount([
      account("first"),
      account("second"),
    ])).toThrow(AuthenticationRequiredError);
  });
});
