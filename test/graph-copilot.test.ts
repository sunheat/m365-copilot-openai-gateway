import { describe, expect, it } from "vitest";
import { createCopilotClient } from "../src/graph-copilot.js";
import type { GatewayConfig } from "../src/config.js";

const config: GatewayConfig = {
  tenantId: "123e4567-e89b-12d3-a456-426614174000",
  clientId: "123e4567-e89b-12d3-a456-426614174001",
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

  it("requests the native Graph SSE endpoint and validates its response", async () => {
    let requestUrl = "";
    let requestInit: RequestInit | undefined;
    const client = createCopilotClient(config, async (input, init) => {
      requestUrl = String(input);
      requestInit = init;
      return new Response(
        'data: {"copilotConversation":{"messages":[{"text":"Hello"}]}}\n\n',
        { status: 200, headers: { "Content-Type": "text/event-stream; charset=utf-8" } },
      );
    });

    let activityCount = 0;
    const stream = await client.chatStream(
      "test-token",
      "conversation/123",
      "prompt",
      new AbortController().signal,
      () => { activityCount += 1; },
    );
    const events = [];
    for await (const event of stream) events.push(event);

    expect(requestUrl).toBe("https://graph.microsoft.com/beta/copilot/conversations/conversation%2F123/chatOverStream");
    expect(requestInit?.method).toBe("POST");
    expect(new Headers(requestInit?.headers).get("accept")).toBe("text/event-stream");
    expect(events[0]?.copilotConversation.messages?.[0]?.text).toBe("Hello");
    expect(activityCount).toBe(1);
  });

  it("rejects a successful response with the wrong media type", async () => {
    const client = createCopilotClient(config, async () => new Response("not an event stream", {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }));

    await expect(client.chatStream("test-token", "conversation-123", "prompt", new AbortController().signal))
      .rejects.toMatchObject({ statusCode: 502 });
  });
});
