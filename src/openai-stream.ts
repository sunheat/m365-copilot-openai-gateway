import type {
  GraphChatResponse,
  GraphConversation,
  OpenAIChatCompletionChunk,
  OpenAIFunctionToolCall,
} from "./types.js";
import { GATEWAY_MODEL_ID } from "./openai-mapper.js";

export interface StreamProjectionState {
  completionId: string;
  created: number;
  emittedText: string;
  emittedRole: boolean;
  emittedContent: boolean;
}

export class StreamProjectionError extends Error {
  public constructor(message: string) {
    super(message);
    this.name = "StreamProjectionError";
  }
}

export interface StreamProjector {
  readonly state: StreamProjectionState;
  roleChunk(): OpenAIChatCompletionChunk | undefined;
  contentChunk(snapshot: GraphChatResponse): OpenAIChatCompletionChunk | undefined;
  finalChunk(): OpenAIChatCompletionChunk;
}

function chunk(
  state: StreamProjectionState,
  delta: OpenAIChatCompletionChunk["choices"][number]["delta"],
  finishReason: "stop" | "tool_calls" | null,
): OpenAIChatCompletionChunk {
  return {
    id: state.completionId,
    object: "chat.completion.chunk",
    created: state.created,
    model: GATEWAY_MODEL_ID,
    choices: [{ index: 0, delta, finish_reason: finishReason }],
  };
}

export function bufferedTextChunks(
  conversation: GraphConversation,
  content: string,
  created = Math.floor(Date.now() / 1_000),
): OpenAIChatCompletionChunk[] {
  const state: StreamProjectionState = {
    completionId: `chatcmpl-${conversation.id}`,
    created,
    emittedText: content,
    emittedRole: true,
    emittedContent: true,
  };
  return [
    chunk(state, { role: "assistant", content: "" }, null),
    chunk(state, { content }, null),
    chunk(state, {}, "stop"),
  ];
}

export function bufferedToolCallChunks(
  conversation: GraphConversation,
  toolCall: OpenAIFunctionToolCall,
  created = Math.floor(Date.now() / 1_000),
): OpenAIChatCompletionChunk[] {
  const state: StreamProjectionState = {
    completionId: `chatcmpl-${conversation.id}`,
    created,
    emittedText: "",
    emittedRole: true,
    emittedContent: false,
  };
  return [
    chunk(state, { role: "assistant", content: "" }, null),
    chunk(state, { tool_calls: [{ index: 0, ...toolCall }] }, null),
    chunk(state, {}, "tool_calls"),
  ];
}

function latestText(snapshot: GraphChatResponse): string | undefined {
  for (let index = (snapshot.messages?.length ?? 0) - 1; index >= 0; index -= 1) {
    const text = snapshot.messages?.[index]?.text;
    if (typeof text !== "string") continue;
    const normalized = text.startsWith("\uFEFF") ? text.slice(1) : text;
    if (normalized.length > 0) return normalized;
  }
  return undefined;
}

export function createStreamProjector(
  conversation: GraphConversation,
  created = Math.floor(Date.now() / 1_000),
): StreamProjector {
  const state: StreamProjectionState = {
    completionId: `chatcmpl-${conversation.id}`,
    created,
    emittedText: "",
    emittedRole: false,
    emittedContent: false,
  };

  return {
    state,

    roleChunk() {
      if (state.emittedRole) return undefined;
      state.emittedRole = true;
      return chunk(state, { role: "assistant", content: "" }, null);
    },

    contentChunk(snapshot) {
      const text = latestText(snapshot);
      if (text === undefined) return undefined;
      if (text === state.emittedText) return undefined;
      if (state.emittedText !== "" && !text.startsWith(state.emittedText)) {
        throw new StreamProjectionError("Microsoft Graph returned a divergent cumulative snapshot.");
      }

      const suffix = text.slice(state.emittedText.length);
      state.emittedText = text;
      if (suffix.length === 0) return undefined;
      state.emittedContent = true;
      return chunk(state, { content: suffix }, null);
    },

    finalChunk() {
      if (!state.emittedContent) {
        throw new StreamProjectionError("Microsoft Graph stream ended without assistant content.");
      }
      return chunk(state, {}, "stop");
    },
  };
}

export function serializeSseData(data: OpenAIChatCompletionChunk | { error: Record<string, string> } | "[DONE]"): string {
  const serialized = data === "[DONE]" ? data : JSON.stringify(data);
  return `data: ${serialized}\n\n`;
}
