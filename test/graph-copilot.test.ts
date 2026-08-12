import { describe, expect, it } from "vitest";
import { createCopilotClient } from "../src/graph-copilot.js";
import type { GatewayConfig } from "../src/config.js";

const config: GatewayConfig = {
  tenantId: "123e4567-e89b-12d3-a456-426614174000",
  clientId: "123e4567-e89b-12d3-a456-426614174001",
  host: "127.0.0.1",
  port: 8787,
  timeZone: "Australia/Sydney",
  tokenCacheDirectory: "C:/test/cache",
  graphBaseUrl: "https://graph.microsoft.com/beta",
};

describe("createCopilotClient", () => {
  it("preserves Graph Retry-After metadata on throttling errors", async () => {
    const client = createCopilotClient(config, async () => new Response("throttled", {
      status: 429,
      headers: { "Retry-After": "30" },
    }));

    await expect(client.createConversation("test-token")).rejects.toMatchObject({
      statusCode: 429,
      retryAfter: "30",
    });
  });
});
