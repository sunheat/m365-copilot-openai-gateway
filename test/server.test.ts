import { EventEmitter } from "node:events";
import type { ServerResponse } from "node:http";
import { describe, expect, it } from "vitest";
import { InteractionRequiredAuthError } from "@azure/msal-node";
import type { AuthService } from "../src/auth.js";
import type { GatewayConfig } from "../src/config.js";
import { GraphCopilotError, type CopilotClient } from "../src/graph-copilot.js";
import type { GatewayLogger } from "../src/logger.js";
import { buildServer, writeSse } from "../src/server.js";

const config: GatewayConfig = {
  tenantId: "00000000-0000-0000-0000-000000000001",
  clientId: "00000000-0000-0000-0000-000000000002",
  host: "127.0.0.1",
  port: 8787,
  logLevel: "silent",
  logFormat: "pretty",
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
  loginInteractively: async () => ({ username: "test@example.com" }),
};

const copilot: CopilotClient = {
  createConversation: async () => ({ id: "conversation-123" }),
  chat: async () => ({ messages: [{ text: "Gateway test successful." }] }),
  chatStream: async () => (async function* () {
    yield { copilotConversation: { messages: [{ text: "Gateway stream successful." }] } };
  })(),
};

const weatherTool = {
  type: "function" as const,
  function: {
    name: "get_weather",
    description: "Get the weather for a city.",
    parameters: {
      type: "object",
      properties: { city: { type: "string" } },
      required: ["city"],
      additionalProperties: false,
    },
  },
};

function protocolNonce(prompt: string): string {
  const match = prompt.match(/"protocol":"m365-copilot-openai-gateway\/tool-call\/v1","nonce":"([^"]+)"/);
  if (!match?.[1]) throw new Error("Tool protocol prompt did not contain a nonce.");
  return match[1];
}

describe("gateway server", () => {
  it("returns a Copilot answer using the OpenAI chat completion shape", async () => {
    const app = buildServer({ config, auth, copilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any-client-model", messages: [{ role: "user", content: "Hello" }] },
    });
    expect(response.statusCode).toBe(200);
    expect(response.headers["x-request-id"]).toBeTruthy();
    expect(response.json().choices[0].message.content).toBe("Gateway test successful.");
    await app.close();
  });

  it("does not log sync completion before response mapping succeeds", async () => {
    const events: string[] = [];
    const logger: GatewayLogger = {
      error: (event) => { events.push(`error:${event}`); },
      warn: (event) => { events.push(`warn:${event}`); },
      info: (event) => { events.push(`info:${event}`); },
      debug: (event) => { events.push(`debug:${event}`); },
      trace: (event) => { events.push(`trace:${event}`); },
    };
    const emptyResponseCopilot: CopilotClient = {
      ...copilot,
      chat: async () => ({ messages: [] }),
    };
    const app = buildServer({ config, auth, copilot: emptyResponseCopilot, logger });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: { model: "any-client-model", messages: [{ role: "user", content: "Hello" }] },
    });

    expect(response.statusCode).toBe(502);
    expect(events).toContain("error:request_failed");
    expect(events).not.toContain("info:request_completed");
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
    expect(response.headers["x-request-id"]).toBeTruthy();
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

  it("bounds buffered tool token acquisition with the request deadline", async () => {
    const stalledAuth: AuthService = {
      ...auth,
      getAccessToken: async () => new Promise<string>(() => undefined),
    };
    const app = buildServer({
      config: { ...config, graphStreamIdleTimeoutMs: 10 },
      auth: stalledAuth,
      copilot,
    });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        tools: [weatherTool],
        messages: [{ role: "user", content: "Use the tool." }],
      },
    });

    expect(response.statusCode).toBe(502);
    await app.close();
  });

  it("cancels buffered token acquisition before Fastify waits during shutdown", async () => {
    let tokenRequested = false;
    const stalledAuth: AuthService = {
      ...auth,
      getAccessToken: async () => {
        tokenRequested = true;
        return new Promise<string>(() => undefined);
      },
    };
    const app = buildServer({ config, auth: stalledAuth, copilot });
    await app.listen({ host: "127.0.0.1", port: 0 });
    const address = app.server.address();
    if (!address || typeof address === "string") throw new Error("Test server did not expose a TCP address.");

    const request = fetch(`http://127.0.0.1:${address.port}/v1/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "any",
        tools: [weatherTool],
        messages: [{ role: "user", content: "Use the tool." }],
      }),
    }).catch((error: unknown) => error);
    await expect.poll(() => tokenRequested, { timeout: 1_000 }).toBe(true);

    await app.close();
    await request;
  });

  it("aborts a buffered tool call when the downstream response closes", async () => {
    let toolChatStarted = false;
    let upstreamAborted = false;
    const cancellableCopilot: CopilotClient = {
      ...copilot,
      chat: async (_token, _conversationId, _prompt, signal) => new Promise((_resolve, reject) => {
        if (!signal) throw new Error("Buffered tool chat did not receive an AbortSignal.");
        toolChatStarted = true;
        const onAbort = (): void => {
          upstreamAborted = true;
          reject(new Error("aborted"));
        };
        if (signal.aborted) {
          onAbort();
          return;
        }
        signal.addEventListener("abort", onAbort, { once: true });
      }),
    };
    const app = buildServer({ config, auth, copilot: cancellableCopilot });
    await app.listen({ host: "127.0.0.1", port: 0 });
    const address = app.server.address();
    if (!address || typeof address === "string") throw new Error("Test server did not expose a TCP address.");

    const clientAbort = new AbortController();
    const request = fetch(`http://127.0.0.1:${address.port}/v1/chat/completions`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "any",
        tools: [weatherTool],
        messages: [{ role: "user", content: "Use the tool." }],
      }),
      signal: clientAbort.signal,
    });
    await expect.poll(() => toolChatStarted, { timeout: 1_000 }).toBe(true);
    clientAbort.abort();

    await expect(request).rejects.toThrow();
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

  it("maps a validated function decision to a non-streaming OpenAI tool call", async () => {
    const toolCopilot: CopilotClient = {
      ...copilot,
      chat: async (_token, _conversationId, prompt) => ({
        messages: [{
          text: JSON.stringify({
            type: "tool_call",
            nonce: protocolNonce(prompt),
            name: "get_weather",
            arguments: { city: "Sydney" },
          }),
        }],
      }),
    };
    const app = buildServer({ config, auth, copilot: toolCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        tools: [weatherTool],
        tool_choice: { type: "function", function: { name: "get_weather" } },
        messages: [{ role: "user", content: "What is the weather?" }],
      },
    });

    expect(response.statusCode).toBe(200);
    expect(response.json().choices[0].message.content).toBeNull();
    expect(response.json().choices[0].message.tool_calls[0]).toMatchObject({
      type: "function",
      function: { name: "get_weather", arguments: '{"city":"Sydney"}' },
    });
    expect(response.json().choices[0].finish_reason).toBe("tool_calls");
    await app.close();
  });

  it("maps a validated function decision to buffered OpenAI SSE chunks", async () => {
    const toolCopilot: CopilotClient = {
      ...copilot,
      chat: async (_token, _conversationId, prompt) => ({
        messages: [{
          text: JSON.stringify({
            type: "tool_call",
            nonce: protocolNonce(prompt),
            name: "get_weather",
            arguments: { city: "Sydney" },
          }),
        }],
      }),
    };
    const app = buildServer({ config, auth, copilot: toolCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        stream: true,
        tools: [weatherTool],
        messages: [{ role: "user", content: "What is the weather?" }],
      },
    });

    expect(response.statusCode).toBe(200);
    expect(response.headers["content-type"]).toBe("text/event-stream; charset=utf-8");
    expect(response.body).toContain('"tool_calls":[{"index":0');
    expect(response.body).toContain('"name":"get_weather"');
    expect(response.body).toContain('"finish_reason":"tool_calls"');
    expect(response.body.match(/data: \[DONE\]/g)).toHaveLength(1);
    await app.close();
  });

  it("returns a normal final answer when auto tool choice does not need a tool", async () => {
    const finalCopilot: CopilotClient = {
      ...copilot,
      chat: async (_token, _conversationId, prompt) => ({
        messages: [{
          text: JSON.stringify({
            type: "final",
            nonce: protocolNonce(prompt),
            content: "No tool needed.",
          }),
        }],
      }),
    };
    const app = buildServer({ config, auth, copilot: finalCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        tools: [weatherTool],
        tool_choice: "auto",
        messages: [{ role: "user", content: "Say hello." }],
      },
    });

    expect(response.statusCode).toBe(200);
    expect(response.json().choices[0].message.content).toBe("No tool needed.");
    expect(response.json().choices[0].finish_reason).toBe("stop");
    await app.close();
  });

  it("uses one bounded correction turn for an invalid required-tool response", async () => {
    let nonce = "";
    const prompts: string[] = [];
    const correctingCopilot: CopilotClient = {
      ...copilot,
      chat: async (_token, _conversationId, prompt) => {
        prompts.push(prompt);
        if (prompts.length === 1) {
          nonce = protocolNonce(prompt);
          return { messages: [{ text: JSON.stringify({ type: "final", nonce, content: "Not allowed." }) }] };
        }
        return {
          messages: [{
            text: JSON.stringify({
              type: "tool_call",
              nonce,
              name: "get_weather",
              arguments: { city: "Sydney" },
            }),
          }],
        };
      },
    };
    const app = buildServer({ config, auth, copilot: correctingCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        tools: [weatherTool],
        tool_choice: "required",
        messages: [{ role: "user", content: "Use the tool." }],
      },
    });

    expect(response.statusCode).toBe(200);
    expect(prompts).toHaveLength(2);
    expect(prompts[1]).toContain("Validation category: tool_required.");
    expect(response.json().choices[0].finish_reason).toBe("tool_calls");
    await app.close();
  });

  it("fails closed after one invalid correction response", async () => {
    let callCount = 0;
    const invalidCopilot: CopilotClient = {
      ...copilot,
      chat: async () => {
        callCount += 1;
        return { messages: [{ text: "not json" }] };
      },
    };
    const app = buildServer({ config, auth, copilot: invalidCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        tools: [weatherTool],
        messages: [{ role: "user", content: "Use the tool." }],
      },
    });

    expect(response.statusCode).toBe(502);
    expect(response.json().error.code).toBe("tool_protocol_error");
    expect(callCount).toBe(2);
    await app.close();
  });

  it("rejects invalid tool schemas before calling Microsoft Graph", async () => {
    let graphCalled = false;
    const trackingCopilot: CopilotClient = {
      ...copilot,
      createConversation: async () => {
        graphCalled = true;
        return { id: "conversation-123" };
      },
    };
    const app = buildServer({ config, auth, copilot: trackingCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        tools: [{
          type: "function",
          function: { name: "broken", parameters: { type: "not-a-json-schema-type" } },
        }],
        messages: [{ role: "user", content: "Hello" }],
      },
    });

    expect(response.statusCode).toBe(400);
    expect(response.json().error.code).toBe("invalid_tool_definition");
    expect(graphCalled).toBe(false);
    await app.close();
  });

  it("keeps the native text path when tool_choice is none", async () => {
    let graphPrompt = "";
    const textCopilot: CopilotClient = {
      ...copilot,
      chat: async (_token, _conversationId, prompt) => {
        graphPrompt = prompt;
        return { messages: [{ text: "Text only." }] };
      },
    };
    const app = buildServer({ config, auth, copilot: textCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        tools: [weatherTool],
        tool_choice: "none",
        messages: [{ role: "user", content: "Hello" }],
      },
    });

    expect(response.statusCode).toBe(200);
    expect(response.json().choices[0].message.content).toBe("Text only.");
    expect(graphPrompt).not.toContain("m365-copilot-openai-gateway/tool-call/v1");
    await app.close();
  });

  it("still validates tool definitions when tool_choice is none", async () => {
    let graphCalled = false;
    const trackingCopilot: CopilotClient = {
      ...copilot,
      createConversation: async () => {
        graphCalled = true;
        return { id: "conversation-123" };
      },
    };
    const app = buildServer({ config, auth, copilot: trackingCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        tool_choice: "none",
        tools: [{
          type: "function",
          function: { name: "broken", parameters: { type: "not-a-json-schema-type" } },
        }],
        messages: [{ role: "user", content: "Hello" }],
      },
    });

    expect(response.statusCode).toBe(400);
    expect(response.json().error.code).toBe("invalid_tool_definition");
    expect(graphCalled).toBe(false);
    await app.close();
  });

  it("rejects a required tool choice when no tools are supplied", async () => {
    const app = buildServer({ config, auth, copilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        tool_choice: "required",
        messages: [{ role: "user", content: "Hello" }],
      },
    });

    expect(response.statusCode).toBe(400);
    expect(response.json().error.code).toBe("invalid_tool_definition");
    await app.close();
  });

  it("preserves assistant tool calls and tool results in a follow-up turn", async () => {
    let graphPrompt = "";
    const followUpCopilot: CopilotClient = {
      ...copilot,
      chat: async (_token, _conversationId, prompt) => {
        graphPrompt = prompt;
        return {
          messages: [{
            text: JSON.stringify({ type: "final", nonce: protocolNonce(prompt), content: "It is sunny." }),
          }],
        };
      },
    };
    const app = buildServer({ config, auth, copilot: followUpCopilot });
    const response = await app.inject({
      method: "POST",
      url: "/v1/chat/completions",
      payload: {
        model: "any",
        tools: [weatherTool],
        messages: [
          { role: "user", content: "What is the weather?" },
          {
            role: "assistant",
            content: null,
            tool_calls: [{
              id: "call-1",
              type: "function",
              function: { name: "get_weather", arguments: '{"city":"Sydney"}' },
            }],
          },
          { role: "tool", tool_call_id: "call-1", content: '{"condition":"sunny"}' },
        ],
      },
    });

    expect(response.statusCode).toBe(200);
    expect(response.json().choices[0].message.content).toBe("It is sunny.");
    expect(graphPrompt).toContain('"tool_call_id":"call-1"');
    await app.close();
  });

  it("maps MSAL interaction-required failures to a login-required response", async () => {
    const authRequiringInteraction: AuthService = {
      getAccessToken: async () => {
        throw new InteractionRequiredAuthError("interaction_required", "test-correlation");
      },
      loginInteractively: async () => ({ username: "test@example.com" }),
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
