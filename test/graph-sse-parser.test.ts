import { describe, expect, it } from "vitest";
import { parseGraphSse, GraphSseParseError } from "../src/graph-sse-parser.js";

function streamFromChunks(chunks: Uint8Array[]): ReadableStream<Uint8Array> {
  return new ReadableStream({
    start(controller) {
      for (const chunk of chunks) controller.enqueue(chunk);
      controller.close();
    },
  });
}

function encodedChunks(text: string, splitPoints: number[]): Uint8Array[] {
  const bytes = new TextEncoder().encode(text);
  const points = [0, ...splitPoints, bytes.byteLength];
  return points.slice(0, -1).map((start, index) => bytes.slice(start, points[index + 1]));
}

describe("parseGraphSse", () => {
  it("handles split UTF-8 bytes and CRLF event framing", async () => {
    const payload = 'data: {"copilotConversation":{"messages":[{"text":"你好"}]}}\r\n\r\n';
    const bytes = new TextEncoder().encode(payload);
    const helloOffset = payload.indexOf("你好");
    const split = new TextEncoder().encode(payload.slice(0, helloOffset + 1)).byteLength;
    const events = [];
    for await (const event of parseGraphSse(streamFromChunks(encodedChunks(payload, [split, split + 1])))) {
      events.push(event);
    }

    expect(bytes.byteLength).toBeGreaterThan(split);
    expect(events).toHaveLength(1);
    expect(events[0]?.copilotConversation.messages?.[0]?.text).toBe("你好");
  });

  it("joins data lines and ignores comments, ids, and unknown fields", async () => {
    const body = [
      ": heartbeat",
      "id: upstream-1",
      "event: message",
      'data: {"copilotConversation":',
      'data: {"messages":[{"text":"Hello"}]}}',
      "",
      'data: {"copilotConversation":{"messages":[]}}',
      "",
      "",
    ].join("\n");
    const events = [];
    for await (const event of parseGraphSse(streamFromChunks([new TextEncoder().encode(body)]))) {
      events.push(event);
    }

    expect(events).toHaveLength(2);
    expect(events[0]?.copilotConversation.messages?.[0]?.text).toBe("Hello");
    expect(events[1]?.copilotConversation.messages).toEqual([]);
  });

  it("rejects malformed, incomplete, invalid, and oversized events", async () => {
    await expect(async () => {
      for await (const _event of parseGraphSse(streamFromChunks([new TextEncoder().encode("data: {bad}\n\n")]))) {
        // Consume the generator.
      }
    }).rejects.toBeInstanceOf(GraphSseParseError);

    await expect(async () => {
      for await (const _event of parseGraphSse(streamFromChunks([new TextEncoder().encode("data: {}\n\n")]))) {
        // Consume the generator.
      }
    }).rejects.toBeInstanceOf(GraphSseParseError);

    await expect(async () => {
      for await (const _event of parseGraphSse(streamFromChunks([new TextEncoder().encode("data: {\"copilotConversation\":{\"messages\":[{\"text\":3}]}}\n\n")]))) {
        // Consume the generator.
      }
    }).rejects.toBeInstanceOf(GraphSseParseError);

    await expect(async () => {
      for await (const _event of parseGraphSse(
        streamFromChunks([new TextEncoder().encode('data: {"copilotConversation":{"messages":[{"text":"large"}]}}\n\n')]),
        { maxEventBytes: 16 },
      )) {
        // Consume the generator.
      }
    }).rejects.toBeInstanceOf(GraphSseParseError);

    await expect(async () => {
      for await (const _event of parseGraphSse(streamFromChunks([new TextEncoder().encode("data: {}\n")]))) {
        // Consume the generator.
      }
    }).rejects.toBeInstanceOf(GraphSseParseError);
  });
});
