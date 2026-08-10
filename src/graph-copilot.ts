import type { GatewayConfig } from "./config.js";
import type { GraphChatResponse, GraphConversation } from "./types.js";

export class GraphCopilotError extends Error {
  public constructor(
    public readonly statusCode: number,
    message: string,
  ) {
    super(message);
    this.name = "GraphCopilotError";
  }
}

export interface CopilotClient {
  createConversation(accessToken: string): Promise<GraphConversation>;
  chat(accessToken: string, conversationId: string, prompt: string): Promise<GraphChatResponse>;
}

export function createCopilotClient(config: GatewayConfig, fetcher: typeof fetch = fetch): CopilotClient {
  const baseUrl = `${config.graphBaseUrl}/copilot/conversations`;

  async function post<T>(accessToken: string, url: string, body: object): Promise<T> {
    const response = await fetcher(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const detail = await response.text();
      throw new GraphCopilotError(response.status, `Microsoft Graph request failed (${response.status}): ${detail.slice(0, 500)}`);
    }
    return (await response.json()) as T;
  }

  return {
    createConversation(accessToken) {
      return post<GraphConversation>(accessToken, baseUrl, {});
    },
    chat(accessToken, conversationId, prompt) {
      return post<GraphChatResponse>(accessToken, `${baseUrl}/${encodeURIComponent(conversationId)}/chat`, {
        message: { text: prompt },
        locationHint: { timeZone: config.timeZone },
        contextualResources: { webContext: { isWebEnabled: false } },
      });
    },
  };
}
