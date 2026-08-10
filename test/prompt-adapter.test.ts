import { describe, expect, it } from "vitest";
import { flattenMessages } from "../src/prompt-adapter.js";

describe("flattenMessages", () => {
  it("preserves all supported OpenAI roles in a framed transcript", () => {
    const prompt = flattenMessages([
      { role: "system", content: "Be concise." },
      { role: "user", content: "Review this." },
      { role: "assistant", content: "Okay." },
      { role: "tool", content: "lint passed" },
    ]);
    expect(prompt).toContain("[TOOL RESULT]\nlint passed");
  });
});
