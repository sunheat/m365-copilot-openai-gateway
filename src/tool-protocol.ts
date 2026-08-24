import { Ajv, type ValidateFunction } from "ajv";
import formatsModule from "ajv-formats";
import type { JsonObject, OpenAIFunctionTool, OpenAIToolChoice } from "./types.js";

const TOOL_NAME_PATTERN = /^[A-Za-z0-9_-]{1,64}$/;

export class InvalidToolDefinitionError extends Error {
  public constructor(message: string) {
    super(message);
    this.name = "InvalidToolDefinitionError";
  }
}

export class ToolProtocolError extends Error {
  public constructor(public readonly reason: string) {
    super("Microsoft 365 Copilot returned an invalid tool protocol response.");
    this.name = "ToolProtocolError";
  }
}

export type ToolDecision =
  | { kind: "final"; content: string }
  | { kind: "tool_call"; name: string; arguments: JsonObject };

export interface ToolProtocol {
  readonly nonce: string;
  readonly choice: Exclude<OpenAIToolChoice, "none">;
  parse(responseText: string): ToolDecision;
  correctionPrompt(error: ToolProtocolError): string;
}

function compileToolValidators(tools: OpenAIFunctionTool[]): Map<string, ValidateFunction> {
  if (tools.length === 0) {
    throw new InvalidToolDefinitionError("At least one function tool is required.");
  }

  const ajv = new Ajv({ allErrors: true, strict: false, strictSchema: true });
  formatsModule.default(ajv);
  const validators = new Map<string, ValidateFunction>();
  for (const tool of tools) {
    const name = tool.function.name;
    if (!TOOL_NAME_PATTERN.test(name)) {
      throw new InvalidToolDefinitionError("Function tool names must match ^[A-Za-z0-9_-]{1,64}$.");
    }
    if (validators.has(name)) {
      throw new InvalidToolDefinitionError("Function tool names must be unique.");
    }
    try {
      validators.set(name, ajv.compile(tool.function.parameters));
    } catch {
      throw new InvalidToolDefinitionError("Function tool parameters must be a valid JSON Schema.");
    }
  }
  return validators;
}

export function validateToolDefinitions(tools: OpenAIFunctionTool[]): void {
  compileToolValidators(tools);
}

function isRecord(value: unknown): value is JsonObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function hasExactKeys(value: JsonObject, expected: string[]): boolean {
  const actual = Object.keys(value).sort();
  const sortedExpected = [...expected].sort();
  return actual.length === sortedExpected.length && actual.every((key, index) => key === sortedExpected[index]);
}

function safeCategory(reason: string): string {
  const allowed = new Set([
    "invalid_json",
    "invalid_shape",
    "nonce_mismatch",
    "tool_required",
    "unknown_tool",
    "tool_not_allowed",
    "invalid_arguments",
  ]);
  return allowed.has(reason) ? reason : "invalid_shape";
}

export function createToolProtocol(
  tools: OpenAIFunctionTool[],
  choice: Exclude<OpenAIToolChoice, "none">,
  nonce: string,
): ToolProtocol {
  const validators = compileToolValidators(tools);

  const namedTool = typeof choice === "object" ? choice.function.name : undefined;
  if (namedTool && !validators.has(namedTool)) {
    throw new InvalidToolDefinitionError("tool_choice names a function that is not present in tools.");
  }

  return {
    nonce,
    choice,

    parse(responseText) {
      const trimmed = responseText.replace(/^\uFEFF/, "").trim();
      let value: unknown;
      try {
        value = JSON.parse(trimmed) as unknown;
      } catch {
        throw new ToolProtocolError("invalid_json");
      }
      if (!isRecord(value) || typeof value.type !== "string") {
        throw new ToolProtocolError("invalid_shape");
      }
      if (value.nonce !== nonce) {
        throw new ToolProtocolError("nonce_mismatch");
      }

      if (value.type === "final") {
        if (!hasExactKeys(value, ["type", "nonce", "content"]) || typeof value.content !== "string" || value.content.trim() === "") {
          throw new ToolProtocolError("invalid_shape");
        }
        if (choice === "required" || namedTool) {
          throw new ToolProtocolError("tool_required");
        }
        return { kind: "final", content: value.content };
      }

      if (value.type !== "tool_call" || !hasExactKeys(value, ["type", "nonce", "name", "arguments"])) {
        throw new ToolProtocolError("invalid_shape");
      }
      if (typeof value.name !== "string" || !validators.has(value.name)) {
        throw new ToolProtocolError("unknown_tool");
      }
      if (namedTool && value.name !== namedTool) {
        throw new ToolProtocolError("tool_not_allowed");
      }
      if (!isRecord(value.arguments)) {
        throw new ToolProtocolError("invalid_arguments");
      }
      const validate = validators.get(value.name);
      if (!validate?.(value.arguments)) {
        throw new ToolProtocolError("invalid_arguments");
      }
      return { kind: "tool_call", name: value.name, arguments: value.arguments };
    },

    correctionPrompt(error) {
      const toolRequired = choice === "required" || namedTool !== undefined;
      const allowedShape = toolRequired
        ? "The only allowed shape is {\"type\":\"tool_call\",\"nonce\":NONCE,\"name\":\"allowed tool\",\"arguments\":{}}."
        : "Allowed shapes are {\"type\":\"final\",\"nonce\":NONCE,\"content\":\"text\"} or {\"type\":\"tool_call\",\"nonce\":NONCE,\"name\":\"allowed tool\",\"arguments\":{}}.";
      return [
        "Your previous response failed the application tool protocol validation.",
        `Validation category: ${safeCategory(error.reason)}.`,
        "Do not execute a tool. Serialize the proposed client application action; the client alone decides whether to execute it.",
        `Return exactly one compact JSON object using nonce ${JSON.stringify(nonce)}.`,
        allowedShape,
        "Do not use Markdown fences, commentary, extra fields, or more than one tool call.",
        toolRequired ? "A tool call is required; a final answer is invalid." : "A final answer is allowed.",
        ...(namedTool ? [`The required tool name is ${JSON.stringify(namedTool)}.`] : []),
      ].join("\n");
    },
  };
}
