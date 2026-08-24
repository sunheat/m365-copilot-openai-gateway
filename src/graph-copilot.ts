import type { GatewayConfig } from "./config.js";
import { parseGraphSse } from "./graph-sse-parser.js";
import type { GraphChatResponse, GraphChatStreamEvent, GraphConversation } from "./types.js";

export class GraphCopilotError extends Error {
  public constructor(
    public readonly statusCode: number,
    message: string,
    public readonly retryAfter?: string,
  ) {
    super(message);
    this.name = "GraphCopilotError";
  }
}

export interface CopilotClient {
  createConversation(accessToken: string, signal?: AbortSignal): Promise<GraphConversation>;
  chat(accessToken: string, conversationId: string, prompt: string, signal?: AbortSignal): Promise<GraphChatResponse>;
  chatStream(
    accessToken: string,
    conversationId: string,
    prompt: string,
    signal: AbortSignal,
    onActivity?: () => void,
  ): Promise<AsyncIterable<GraphChatStreamEvent>>;
}

export function createCopilotClient(config: GatewayConfig, fetcher: typeof fetch = fetch): CopilotClient {
  const baseUrl = `${config.graphBaseUrl}/copilot/conversations`;

  async function post<T>(
    accessToken: string,
    url: string,
    body: object,
    options: { accept?: string; signal?: AbortSignal } = {},
  ): Promise<T> {
    const response = await fetcher(url, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
        ...(options.accept ? { Accept: options.accept } : {}),
      },
      body: JSON.stringify(body),
      ...(options.signal ? { signal: options.signal } : {}),
    });

    if (!response.ok) {
      const detail = await response.text();
      throw new GraphCopilotError(
        response.status,
        `Microsoft Graph request failed (${response.status}): ${detail.slice(0, 500)}`,
        response.headers.get("retry-after") ?? undefined,
      );
    }
    return (await response.json()) as T;
  }

  return {
    createConversation(accessToken, signal) {
      return post<GraphConversation>(accessToken, baseUrl, {}, signal ? { signal } : {});
    },
    chat(accessToken, conversationId, prompt, signal) {
      return post<GraphChatResponse>(accessToken, `${baseUrl}/${encodeURIComponent(conversationId)}/chat`, {
        message: { text: prompt },
        locationHint: { timeZone: config.timeZone },
        contextualResources: { webContext: { isWebEnabled: false } },
      }, signal ? { signal } : {});
    },
    async chatStream(accessToken, conversationId, prompt, signal, onActivity) {
      const response = await fetcher(`${baseUrl}/${encodeURIComponent(conversationId)}/chatOverStream`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken}`,
          "Content-Type": "application/json",
          Accept: "text/event-stream",
        },
        body: JSON.stringify({
          message: { text: prompt },
          locationHint: { timeZone: config.timeZone },
          contextualResources: { webContext: { isWebEnabled: false } },
        }),
        signal,
      });

      if (!response.ok) {
        const detail = await response.text();
        throw new GraphCopilotError(
          response.status,
          `Microsoft Graph request failed (${response.status}): ${detail.slice(0, 500)}`,
          response.headers.get("retry-after") ?? undefined,
        );
      }

      const mediaType = response.headers.get("content-type")?.split(";", 1)[0]?.trim().toLowerCase();
      if (mediaType !== "text/event-stream") {
        throw new GraphCopilotError(502, "Microsoft Graph streaming response was not text/event-stream.");
      }
      if (!response.body) {
        throw new GraphCopilotError(502, "Microsoft Graph streaming response did not include a body.");
      }

      return parseGraphSse(response.body, {
        maxEventBytes: config.graphStreamMaxEventBytes,
        ...(onActivity ? { onActivity } : {}),
      });
    },
  };
}
