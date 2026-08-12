import { describe, expect, it } from "vitest";
import { InteractionRequiredAuthError } from "@azure/msal-node";
import type { AuthService } from "../src/auth.js";
import type { GatewayConfig } from "../src/config.js";
import { GraphCopilotError, type CopilotClient } from "../src/graph-copilot.js";
import { buildServer } from "../src/server.js";

const config: GatewayConfig = {
  tenantId: "00000000-0000-0000-0000-000000000001",
  clientId: "00000000-0000-0000-0000-000000000002",
  host: "127.0.0.1",
  port: 8787,
  timeZone: "Australia/Sydney",
  tokenCacheDirectory: "C:/test/cache",
  graphBaseUrl: "https://graph.microsoft.com/beta",
};

const auth: AuthService = {
  getAccessToken: async () => "test-token",
  loginWithDeviceCode: async () => ({ username: "test@example.com" }),
};

const copilot: CopilotClient = {
  createConversation: async () => ({ id: "conversation-123" }),
  chat: async () => ({ messages: [{ text: "Gateway test successful." }] }),
};

describe("gateway server", () => {
  it("returns a Copilot answer using the OpenAI chat completion shape", async () => {
    const app = buildServer({ config, auth, copilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any-client-model", messages: [{ role: "user", content: "Hello" }] },
    });
    expect(response.statusCode).toBe(200);
    expect(response.json().choices[0].message.content).toBe("Gateway test successful.");
    await app.close();
  });

  it("rejects phase-1 streaming requests explicitly", async () => {
    const app = buildServer({ config, auth, copilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any", stream: true, messages: [{ role: "user", content: "Hello" }] },
    });
    expect(response.statusCode).toBe(400);
    expect(response.json().error.code).toBe("streaming_not_supported");
    await app.close();
  });

  it("maps MSAL interaction-required failures to a login-required response", async () => {
    const authRequiringInteraction: AuthService = {
      getAccessToken: async () => {
        throw new InteractionRequiredAuthError("interaction_required", "test-correlation");
      },
      loginWithDeviceCode: async () => ({ username: "test@example.com" }),
    };
    const app = buildServer({ config, auth: authRequiringInteraction, copilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any", messages: [{ role: "user", content: "Hello" }] },
    });

    expect(response.statusCode).toBe(401);
    expect(response.json().error.code).toBe("m365_login_required");
    await app.close();
  });

  it("preserves Graph throttling as a rate-limit response with Retry-After", async () => {
    const throttledCopilot: CopilotClient = {
      createConversation: async () => {
        throw new GraphCopilotError(429, "throttled", "30");
      },
      chat: async () => ({ messages: [{ text: "unreachable" }] }),
    };
    const app = buildServer({ config, auth, copilot: throttledCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any", messages: [{ role: "user", content: "Hello" }] },
    });

    expect(response.statusCode).toBe(429);
    expect(response.headers["retry-after"]).toBe("30");
    expect(response.json().error.type).toBe("rate_limit_error");
    expect(response.json().error.code).toBe("rate_limit_exceeded");
    await app.close();
  });
});
