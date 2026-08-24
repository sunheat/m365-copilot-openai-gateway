import { describe, expect, it } from "vitest";
import {
  createToolProtocol,
  InvalidToolDefinitionError,
  ToolProtocolError,
} from "../src/tool-protocol.js";
import type { OpenAIFunctionTool } from "../src/types.js";

const tools: OpenAIFunctionTool[] = [{
  type: "function",
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
}];

describe("createToolProtocol", () => {
  it("accepts validated final answers and function calls", () => {
    const protocol = createToolProtocol(tools, "auto", "nonce-123");

    expect(protocol.parse('{"type":"final","nonce":"nonce-123","content":"No tool needed."}'))
      .toEqual({ kind: "final", content: "No tool needed." });
    expect(protocol.parse('{"type":"tool_call","nonce":"nonce-123","name":"get_weather","arguments":{"city":"Sydney"}}'))
      .toEqual({ kind: "tool_call", name: "get_weather", arguments: { city: "Sydney" } });
  });

  it("enforces required and named tool choices", () => {
    const required = createToolProtocol(tools, "required", "nonce-123");
    expect(() => required.parse('{"type":"final","nonce":"nonce-123","content":"No."}'))
      .toThrowError(ToolProtocolError);

    const secondTool: OpenAIFunctionTool = {
      type: "function",
      function: { name: "search_docs", parameters: { type: "object" } },
    };
    const named = createToolProtocol(
      [...tools, secondTool],
      { type: "function", function: { name: "search_docs" } },
      "nonce-123",
    );
    expect(() => named.parse('{"type":"tool_call","nonce":"nonce-123","name":"get_weather","arguments":{"city":"Sydney"}}'))
      .toThrowError(ToolProtocolError);
  });

  it("rejects wrappers, nonce mismatches, extra fields, and invalid arguments", () => {
    const protocol = createToolProtocol(tools, "auto", "nonce-123");
    const invalid = [
      '```json\n{"type":"final","nonce":"nonce-123","content":"No."}\n```',
      '{"type":"final","nonce":"wrong","content":"No."}',
      '{"type":"final","nonce":"nonce-123","content":"No.","extra":true}',
      '{"type":"tool_call","nonce":"nonce-123","name":"get_weather","arguments":{"city":42}}',
    ];

    for (const response of invalid) {
      expect(() => protocol.parse(response)).toThrowError(ToolProtocolError);
    }
  });

  it("enforces common JSON Schema string formats", () => {
    const formattedTool: OpenAIFunctionTool = {
      type: "function",
      function: {
        name: "send_email",
        parameters: {
          type: "object",
          properties: { address: { type: "string", format: "email" } },
          required: ["address"],
        },
      },
    };
    const protocol = createToolProtocol([formattedTool], "required", "nonce-123");

    expect(() => protocol.parse('{"type":"tool_call","nonce":"nonce-123","name":"send_email","arguments":{"address":"not-an-email"}}'))
      .toThrowError(ToolProtocolError);
    expect(protocol.parse('{"type":"tool_call","nonce":"nonce-123","name":"send_email","arguments":{"address":"test@example.com"}}'))
      .toMatchObject({ kind: "tool_call", name: "send_email" });
  });

  it("rejects invalid schemas, duplicate names, and missing named choices", () => {
    const invalidSchema: OpenAIFunctionTool = {
      type: "function",
      function: { name: "broken", parameters: { type: "not-a-json-schema-type" } },
    };
    expect(() => createToolProtocol([invalidSchema], "auto", "nonce"))
      .toThrowError(InvalidToolDefinitionError);
    expect(() => createToolProtocol([...tools, tools[0]!], "auto", "nonce"))
      .toThrowError(InvalidToolDefinitionError);
    expect(() => createToolProtocol(tools, { type: "function", function: { name: "missing" } }, "nonce"))
      .toThrowError(InvalidToolDefinitionError);
  });

  it("rejects unsupported JSON Schema dialects and keywords", () => {
    const unsupportedKeyword: OpenAIFunctionTool = {
      type: "function",
      function: {
        name: "modern_constraints",
        parameters: {
          type: "object",
          dependentRequired: { credit_card: ["billing_address"] },
        },
      },
    };
    const unsupportedDialect: OpenAIFunctionTool = {
      type: "function",
      function: {
        name: "modern_dialect",
        parameters: {
          $schema: "https://json-schema.org/draft/2020-12/schema",
          type: "object",
        },
      },
    };

    expect(() => createToolProtocol([unsupportedKeyword], "auto", "nonce"))
      .toThrowError(InvalidToolDefinitionError);
    expect(() => createToolProtocol([unsupportedDialect], "auto", "nonce"))
      .toThrowError(InvalidToolDefinitionError);
  });

  it("builds a bounded correction prompt without echoing model output", () => {
    const protocol = createToolProtocol(tools, "auto", "nonce-123");
    const prompt = protocol.correctionPrompt(new ToolProtocolError("invalid_arguments"));

    expect(prompt).toContain("Validation category: invalid_arguments.");
    expect(prompt).toContain("Do not execute a tool");
    expect(prompt).toContain('nonce "nonce-123"');
    expect(prompt).not.toContain("get_weather");
  });

  it("does not offer the final shape when a tool is required", () => {
    const protocol = createToolProtocol(tools, "required", "nonce-123");
    const prompt = protocol.correctionPrompt(new ToolProtocolError("tool_required"));

    expect(prompt).toContain("The only allowed shape");
    expect(prompt).not.toContain('"type":"final"');
  });
});
