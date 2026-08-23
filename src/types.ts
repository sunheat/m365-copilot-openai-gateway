export const COPILOT_DELEGATED_SCOPES = [
  "https://graph.microsoft.com/Sites.Read.All",
  "https://graph.microsoft.com/Mail.Read",
  "https://graph.microsoft.com/People.Read.All",
  "https://graph.microsoft.com/OnlineMeetingTranscript.Read.All",
  "https://graph.microsoft.com/Chat.Read",
  "https://graph.microsoft.com/ChannelMessage.Read.All",
  "https://graph.microsoft.com/ExternalItem.Read.All",
  "offline_access",
] as const;

export type JsonObject = Record<string, unknown>;

export interface OpenAIFunctionDefinition {
  name: string;
  description?: string;
  parameters: JsonObject;
  strict?: boolean;
}

export interface OpenAIFunctionTool {
  type: "function";
  function: OpenAIFunctionDefinition;
}

export interface OpenAIFunctionToolCall {
  id: string;
  type: "function";
  function: {
    name: string;
    arguments: string;
  };
}

export type OpenAIToolChoice =
  | "none"
  | "auto"
  | "required"
  | { type: "function"; function: { name: string } };

export type OpenAIChatMessage =
  | { role: "system" | "user"; content: string; name?: string | undefined }
  | {
    role: "assistant";
    content: string | null;
    name?: string | undefined;
    tool_calls?: OpenAIFunctionToolCall[] | undefined;
  }
  | { role: "tool"; content: string; tool_call_id: string };

export interface GraphConversation {
  id: string;
}

export interface GraphChatMessage {
  text?: string;
  createdDateTime?: string;
}

export interface GraphChatResponse {
  messages?: GraphChatMessage[];
}

export interface GraphChatStreamEvent {
  copilotConversation: GraphChatResponse;
}

export interface OpenAIChatCompletion {
  id: string;
  object: "chat.completion";
  created: number;
  model: string;
  choices: Array<{
    index: number;
    message: {
      role: "assistant";
      content: string | null;
      tool_calls?: OpenAIFunctionToolCall[];
    };
    finish_reason: "stop" | "tool_calls";
  }>;
}

export interface OpenAIChatCompletionChunk {
  id: string;
  object: "chat.completion.chunk";
  created: number;
  model: string;
  choices: Array<{
    index: 0;
    delta: {
      role?: "assistant";
      content?: string;
      tool_calls?: Array<OpenAIFunctionToolCall & { index: 0 }>;
    };
    finish_reason: "stop" | "tool_calls" | null;
  }>;
}
