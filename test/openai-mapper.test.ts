import { describe, expect, it } from "vitest";
import { GATEWAY_MODEL_ID, toOpenAICompletion } from "../src/openai-mapper.js";

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
