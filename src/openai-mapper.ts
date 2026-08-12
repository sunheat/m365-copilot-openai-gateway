import type { GraphChatResponse, GraphConversation, OpenAIChatCompletion } from "./types.js";

export const GATEWAY_MODEL_ID = "m365-copilot-preview";

export function toOpenAICompletion(
  conversation: GraphConversation,
  graphResponse: GraphChatResponse,
): OpenAIChatCompletion {
  const content = graphResponse.messages?.at(-1)?.text?.trim();
  if (!content) {
    throw new Error("Microsoft Graph returned no assistant text.");
  }
  return {
    id: `chatcmpl-${conversation.id}`,
    object: "chat.completion",
    created: Math.floor(Date.now() / 1_000),
    model: GATEWAY_MODEL_ID,
    choices: [{ index: 0, message: { role: "assistant", content }, finish_reason: "stop" }],
  };
}
