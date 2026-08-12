import { describe, expect, it } from "vitest";
import { homedir } from "node:os";
import path from "node:path";
import { loadConfig } from "../src/config.js";

const validEnvironment = {
  M365_TENANT_ID: "123e4567-e89b-12d3-a456-426614174000",
  M365_CLIENT_ID: "123e4567-e89b-12d3-a456-426614174001",
};

describe("loadConfig", () => {
  it("uses the per-user cache directory when the override is empty", () => {
    const config = loadConfig({ ...validEnvironment, M365_TOKEN_CACHE_DIR: "" });

    expect(config.tokenCacheDirectory).toBe(
      path.join(homedir(), ".m365-copilot-openai-gateway", "msal-cache"),
    );
    expect(config.graphStreamStartTimeoutMs).toBe(30_000);
    expect(config.graphStreamIdleTimeoutMs).toBe(60_000);
    expect(config.graphStreamMaxEventBytes).toBe(2 * 1024 * 1024);
    expect(config.gatewaySseHeartbeatMs).toBe(15_000);
  });
});
