import { describe, expect, it } from "vitest";
import { GATEWAY_MODEL_ID, toOpenAICompletion, toOpenAIToolCompletion } from "../src/openai-mapper.js";

describe("toOpenAICompletion", () => {
  it("maps the final Graph message to an OpenAI-compatible response", () => {
    const response = toOpenAICompletion(
      { id: "conversation-123" },
      { messages: [{ text: "Earlier" }, { text: "Final answer" }] },
    );
    expect(response.id).toBe("chatcmpl-conversation-123");
    expect(response.model).toBe(GATEWAY_MODEL_ID);
    expect(response.choices[0]?.message.content).toBe("Final answer");
  });
});

describe("toOpenAIToolCompletion", () => {
  it("maps a function call with null content and a tool_calls finish reason", () => {
    const response = toOpenAIToolCompletion(
      { id: "conversation-123" },
      {
        id: "call-123",
        type: "function",
        function: { name: "get_weather", arguments: '{"city":"Sydney"}' },
      },
    );

    expect(response.choices[0]?.message.content).toBeNull();
    expect(response.choices[0]?.message.tool_calls?.[0]?.function.name).toBe("get_weather");
    expect(response.choices[0]?.finish_reason).toBe("tool_calls");
  });
});
