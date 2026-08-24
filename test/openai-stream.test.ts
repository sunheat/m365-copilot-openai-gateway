import { describe, expect, it } from "vitest";
import {
  bufferedTextChunks,
  bufferedToolCallChunks,
  createStreamProjector,
  serializeSseData,
  StreamProjectionError,
} from "../src/openai-stream.js";

describe("createStreamProjector", () => {
  it("emits stable role and metadata, then projects cumulative text suffixes", () => {
    const projector = createStreamProjector({ id: "conversation-123" }, 1_780_000_000);
    const role = projector.roleChunk();
    const first = projector.contentChunk({ messages: [{ text: "Hello" }] });
    const duplicate = projector.contentChunk({ messages: [{ text: "Hello" }] });
    const second = projector.contentChunk({ messages: [{ text: "Hello world" }] });
    const final = projector.finalChunk();

    expect(projector.roleChunk()).toBeUndefined();
    expect(role?.choices[0]?.delta).toEqual({ role: "assistant", content: "" });
    expect(first?.choices[0]?.delta).toEqual({ content: "Hello" });
    expect(duplicate).toBeUndefined();
    expect(second?.choices[0]?.delta).toEqual({ content: " world" });
    expect(final.choices[0]?.finish_reason).toBe("stop");
    expect(new Set([role?.id, first?.id, second?.id, final.id])).toEqual(new Set(["chatcmpl-conversation-123"]));
    expect(new Set([role?.created, first?.created, second?.created, final.created])).toEqual(new Set([1_780_000_000]));
  });

  it("preserves whitespace, Markdown, and an optional BOM", () => {
    const projector = createStreamProjector({ id: "conversation-123" });
    projector.roleChunk();
    const content = projector.contentChunk({ messages: [{ text: "\uFEFF  **Hello**\n\n" }] });
    expect(content?.choices[0]?.delta.content).toBe("  **Hello**\n\n");
  });

  it("ignores empty snapshots and fails closed on divergent text", () => {
    const projector = createStreamProjector({ id: "conversation-123" });
    projector.roleChunk();
    expect(projector.contentChunk({ messages: [] })).toBeUndefined();
    projector.contentChunk({ messages: [{ text: "Hello" }] });
    expect(() => projector.contentChunk({ messages: [{ text: "Goodbye" }] })).toThrow(StreamProjectionError);
  });

  it("requires content before completing and serializes one-line SSE frames", () => {
    const projector = createStreamProjector({ id: "conversation-123" });
    expect(() => projector.finalChunk()).toThrow(StreamProjectionError);
    expect(serializeSseData("[DONE]")).toBe("data: [DONE]\n\n");
    expect(serializeSseData({ error: { message: "safe", type: "api_error", code: "stream_error" } }))
      .toBe('data: {"error":{"message":"safe","type":"api_error","code":"stream_error"}}\n\n');
  });
});

describe("buffered compatibility chunks", () => {
  it("emits complete text using the normal stop sequence", () => {
    const chunks = bufferedTextChunks({ id: "conversation-123" }, "Final answer", 1_780_000_000);
    expect(chunks[0]?.choices[0]?.delta).toEqual({ role: "assistant", content: "" });
    expect(chunks[1]?.choices[0]?.delta).toEqual({ content: "Final answer" });
    expect(chunks[2]?.choices[0]?.finish_reason).toBe("stop");
  });

  it("emits one indexed function call and a tool_calls finish reason", () => {
    const chunks = bufferedToolCallChunks(
      { id: "conversation-123" },
      {
        id: "call-123",
        type: "function",
        function: { name: "get_weather", arguments: '{"city":"Sydney"}' },
      },
      1_780_000_000,
    );
    expect(chunks[1]?.choices[0]?.delta.tool_calls?.[0]).toMatchObject({
      index: 0,
      id: "call-123",
      function: { name: "get_weather", arguments: '{"city":"Sydney"}' },
    });
    expect(chunks[2]?.choices[0]?.finish_reason).toBe("tool_calls");
  });
});
