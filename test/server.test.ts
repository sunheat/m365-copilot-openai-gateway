import { EventEmitter } from "node:events";
import type { ServerResponse } from "node:http";
import { describe, expect, it } from "vitest";
import { InteractionRequiredAuthError } from "@azure/msal-node";
import type { AuthService } from "../src/auth.js";
import type { GatewayConfig } from "../src/config.js";
import { GraphCopilotError, type CopilotClient } from "../src/graph-copilot.js";
import { buildServer, writeSse } from "../src/server.js";

const config: GatewayConfig = {
  tenantId: "00000000-0000-0000-0000-000000000001",
  clientId: "00000000-0000-0000-0000-000000000002",
  host: "127.0.0.1",
  port: 8787,
  timeZone: "Australia/Sydney",
  tokenCacheDirectory: "C:/test/cache",
  graphBaseUrl: "https://graph.microsoft.com/beta",
  graphStreamStartTimeoutMs: 30_000,
  graphStreamIdleTimeoutMs: 60_000,
  graphStreamMaxEventBytes: 2 * 1024 * 1024,
  gatewaySseHeartbeatMs: 15_000,
};

const auth: AuthService = {
  getAccessToken: async () => "test-token",
  loginWithDeviceCode: async () => ({ username: "test@example.com" }),
};

const copilot: CopilotClient = {
  createConversation: async () => ({ id: "conversation-123" }),
  chat: async () => ({ messages: [{ text: "Gateway test successful." }] }),
  chatStream: async () => (async function* () {
    yield { copilotConversation: { messages: [{ text: "Gateway stream successful." }] } };
  })(),
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

  it("returns an OpenAI SSE stream for streaming requests", async () => {
    const app = buildServer({ config, auth, copilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        stream: true,
        stream_options: { include_usage: true },
        n: 1,
        messages: [{ role: "user", content: "Hello" }],
      },
    });
    expect(response.statusCode).toBe(200);
    expect(response.headers["content-type"]).toBe("text/event-stream; charset=utf-8");
    expect(response.body).toContain('"role":"assistant"');
    expect(response.body).toContain("Gateway stream successful.");
    expect(response.body.match(/data: \[DONE\]/g)).toHaveLength(1);
    await app.close();
  });

  it("delivers the first content chunk before the upstream stream completes", async () => {
    let releaseSecondSnapshot!: () => void;
    const secondSnapshot = new Promise<void>((resolve) => {
      releaseSecondSnapshot = resolve;
    });
    let upstreamFinished = false;
    const streamingCopilot: CopilotClient = {
      ...copilot,
      chatStream: async () => (async function* () {
        yield { copilotConversation: { messages: [{ text: "Hello" }] } };
        await secondSnapshot;
        yield { copilotConversation: { messages: [{ text: "Hello world" }] } };
        upstreamFinished = true;
      })(),
    };
    const app = buildServer({ config: { ...config, gatewaySseHeartbeatMs: 0 }, auth, copilot: streamingCopilot });
    await app.listen({ host: "127.0.0.1", port: 0 });
    const address = app.server.address();
    if (!address || typeof address === "string") throw new Error("Test server did not expose a TCP address.");

    const response = await fetch(`http://127.0.0.1:${address.port}/v1/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: "any", stream: true, messages: [{ role: "user", content: "Hello" }] }),
    });
    expect(response.status).toBe(200);
    expect(response.headers.get("content-type")).toBe("text/event-stream; charset=utf-8");

    const reader = response.body?.getReader();
    if (!reader) throw new Error("Streaming response did not expose a body.");
    const decoder = new TextDecoder();
    let output = "";
    while (!output.includes('"content":"Hello"')) {
      const { done, value } = await reader.read();
      if (done) throw new Error("Stream ended before the first content chunk.");
      output += decoder.decode(value, { stream: true });
    }
    expect(upstreamFinished).toBe(false);

    releaseSecondSnapshot();
    while (!output.includes("[DONE]")) {
      const { done, value } = await reader.read();
      if (done) break;
      output += decoder.decode(value, { stream: true });
    }
    expect(upstreamFinished).toBe(true);
    expect(output).toContain('"content":" world"');
    expect(output.match(/data: \[DONE\]/g)).toHaveLength(1);
    await app.close();
  });

  it("closes an established stream with a safe error and no DONE sentinel", async () => {
    const failedCopilot: CopilotClient = {
      ...copilot,
      chatStream: async () => (async function* () {
        yield { copilotConversation: { messages: [{ text: "Partial" }] } };
        throw new Error("upstream details must not leak");
      })(),
    };
    const app = buildServer({ config: { ...config, gatewaySseHeartbeatMs: 0 }, auth, copilot: failedCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any", stream: true, messages: [{ role: "user", content: "Hello" }] },
    });

    expect(response.statusCode).toBe(200);
    expect(response.body).toContain('"code":"stream_upstream_error"');
    expect(response.body).not.toContain("upstream details");
    expect(response.body).not.toContain("[DONE]");
    await app.close();
  });

  it("closes the upstream iterator after a projection error", async () => {
    let iteratorClosed = false;
    const divergentCopilot: CopilotClient = {
      ...copilot,
      chatStream: async () => (async function* () {
        try {
          yield { copilotConversation: { messages: [{ text: "Hello" }] } };
          yield { copilotConversation: { messages: [{ text: "Goodbye" }] } };
        } finally {
          iteratorClosed = true;
        }
      })(),
    };
    const app = buildServer({ config: { ...config, gatewaySseHeartbeatMs: 0 }, auth, copilot: divergentCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any", stream: true, messages: [{ role: "user", content: "Hello" }] },
    });

    expect(response.statusCode).toBe(200);
    expect(response.body).toContain('"code":"stream_protocol_error"');
    expect(response.body).not.toContain("[DONE]");
    expect(iteratorClosed).toBe(true);
    await app.close();
  });

  it("turns an upstream idle timeout into a safe incomplete-stream error", async () => {
    const idleCopilot: CopilotClient = {
      ...copilot,
      chatStream: async () => (async function* () {
        yield { copilotConversation: { messages: [{ text: "Partial" }] } };
        await new Promise((resolve) => setTimeout(resolve, 30));
        yield { copilotConversation: { messages: [{ text: "Partial answer" }] } };
      })(),
    };
    const app = buildServer({
      config: { ...config, graphStreamIdleTimeoutMs: 10, gatewaySseHeartbeatMs: 5 },
      auth,
      copilot: idleCopilot,
    });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any", stream: true, messages: [{ role: "user", content: "Hello" }] },
    });

    expect(response.statusCode).toBe(200);
    expect(response.body).toContain('"code":"stream_idle_timeout"');
    expect(response.body).not.toContain("[DONE]");
    await app.close();
  });

  it("resets the idle deadline when upstream byte activity arrives", async () => {
    const activeCopilot: CopilotClient = {
      ...copilot,
      chatStream: async (_token, _conversationId, _prompt, _signal, onActivity) => (async function* () {
        await new Promise((resolve) => setTimeout(resolve, 5));
        onActivity?.();
        await new Promise((resolve) => setTimeout(resolve, 8));
        yield { copilotConversation: { messages: [{ text: "Active" }] } };
      })(),
    };
    const app = buildServer({
      config: { ...config, graphStreamIdleTimeoutMs: 10, gatewaySseHeartbeatMs: 0 },
      auth,
      copilot: activeCopilot,
    });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any", stream: true, messages: [{ role: "user", content: "Hello" }] },
    });

    expect(response.statusCode).toBe(200);
    expect(response.body).toContain("Active");
    expect(response.body).toContain("[DONE]");
    await app.close();
  });

  it("aborts active streams before Fastify waits during shutdown", async () => {
    let upstreamAborted = false;
    const shutdownCopilot: CopilotClient = {
      ...copilot,
      chatStream: async (_token, _conversationId, _prompt, signal) => (async function* () {
        yield { copilotConversation: { messages: [{ text: "Partial" }] } };
        await new Promise<void>((resolve) => {
          if (signal.aborted) {
            upstreamAborted = true;
            resolve();
            return;
          }
          signal.addEventListener("abort", () => {
            upstreamAborted = true;
            resolve();
          }, { once: true });
        });
      })(),
    };
    const app = buildServer({ config: { ...config, gatewaySseHeartbeatMs: 0 }, auth, copilot: shutdownCopilot });
    await app.listen({ host: "127.0.0.1", port: 0 });
    const address = app.server.address();
    if (!address || typeof address === "string") throw new Error("Test server did not expose a TCP address.");

    const response = await fetch(`http://127.0.0.1:${address.port}/v1/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: "any", stream: true, messages: [{ role: "user", content: "Hello" }] }),
    });
    const reader = response.body?.getReader();
    if (!reader) throw new Error("Streaming response did not expose a body.");
    await reader.read();

    await app.close();
    expect(upstreamAborted).toBe(true);
  });

  it("applies the start deadline and abort signal while creating a conversation", async () => {
    let conversationAborted = false;
    const stalledCopilot: CopilotClient = {
      ...copilot,
      createConversation: async (_token, signal) => new Promise((resolve) => {
        if (!signal) throw new Error("Streaming conversation creation did not receive an AbortSignal.");
        if (signal.aborted) {
          conversationAborted = true;
          resolve({ id: "conversation-123" });
          return;
        }
        signal.addEventListener("abort", () => {
          conversationAborted = true;
          resolve({ id: "conversation-123" });
        }, { once: true });
      }),
    };
    const app = buildServer({
      config: { ...config, graphStreamStartTimeoutMs: 10, gatewaySseHeartbeatMs: 0 },
      auth,
      copilot: stalledCopilot,
    });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any", stream: true, messages: [{ role: "user", content: "Hello" }] },
    });

    expect(response.statusCode).toBe(502);
    expect(conversationAborted).toBe(true);
    await app.close();
  });

  it("releases a backpressure wait when its abort signal fires", async () => {
    class BackpressuredResponse extends EventEmitter {
      public destroyed = false;
      public writableEnded = false;

      public write(): boolean {
        return false;
      }
    }

    const response = new BackpressuredResponse() as unknown as ServerResponse;
    const controller = new AbortController();
    const pending = writeSse(response, "data: test\n\n", controller.signal);
    controller.abort();

    await expect(pending).rejects.toThrow("downstream client disconnected");
  });

  it("aborts the upstream stream when the downstream client disconnects", async () => {
    let upstreamAborted = false;
    const cancellableCopilot: CopilotClient = {
      ...copilot,
      chatStream: async (_token, _conversationId, _prompt, signal) => (async function* () {
        yield { copilotConversation: { messages: [{ text: "Partial" }] } };
        await new Promise<void>((resolve) => {
          signal.addEventListener("abort", () => {
            upstreamAborted = true;
            resolve();
          }, { once: true });
        });
      })(),
    };
    const app = buildServer({ config: { ...config, gatewaySseHeartbeatMs: 0 }, auth, copilot: cancellableCopilot });
    await app.listen({ host: "127.0.0.1", port: 0 });
    const address = app.server.address();
    if (!address || typeof address === "string") throw new Error("Test server did not expose a TCP address.");

    const clientAbort = new AbortController();
    const response = await fetch(`http://127.0.0.1:${address.port}/v1/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: "any", stream: true, messages: [{ role: "user", content: "Hello" }] }),
      signal: clientAbort.signal,
    });
    const reader = response.body?.getReader();
    if (!reader) throw new Error("Streaming response did not expose a body.");
    await reader.read();
    clientAbort.abort();
    await expect.poll(() => upstreamAborted, { timeout: 1_000 }).toBe(true);
    await app.close();
  });

  it("allows an empty tools list for a text-only request", async () => {
    const app = buildServer({ config, auth, copilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any", tools: [], messages: [{ role: "user", content: "Hello" }] },
    });
    expect(response.statusCode).toBe(200);
    await app.close();
  });

  it("rejects non-empty tool definitions explicitly", async () => {
    const app = buildServer({ config, auth, copilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any", tools: [{ type: "function" }], messages: [{ role: "user", content: "Hello" }] },
    });
    expect(response.statusCode).toBe(400);
    expect(response.json().error.code).toBe("tools_not_supported");
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
