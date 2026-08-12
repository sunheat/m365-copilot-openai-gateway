import type { GraphChatMessage, GraphChatStreamEvent } from "./types.js";

const DEFAULT_MAX_EVENT_BYTES = 2 * 1024 * 1024;

export class GraphSseParseError extends Error {
  public constructor(message: string) {
    super(message);
    this.name = "GraphSseParseError";
  }
}

interface GraphSseParserOptions {
  maxEventBytes?: number;
  onActivity?: () => void;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function validateEvent(value: unknown): GraphChatStreamEvent {
  if (!isRecord(value) || !isRecord(value.copilotConversation)) {
    throw new GraphSseParseError("Microsoft Graph SSE event has no valid copilotConversation snapshot.");
  }

  const rawMessages = value.copilotConversation.messages;
  if (!Array.isArray(rawMessages)) {
    throw new GraphSseParseError("Microsoft Graph SSE snapshot has no valid messages array.");
  }

  const messages: GraphChatMessage[] = rawMessages.map((message) => {
    if (!isRecord(message)) {
      throw new GraphSseParseError("Microsoft Graph SSE snapshot contains an invalid message.");
    }

    const text = message.text;
    if (text !== undefined && typeof text !== "string") {
      throw new GraphSseParseError("Microsoft Graph SSE message text is not a string.");
    }

    const createdDateTime = message.createdDateTime;
    if (createdDateTime !== undefined && typeof createdDateTime !== "string") {
      throw new GraphSseParseError("Microsoft Graph SSE message timestamp is not a string.");
    }

    return {
      ...(text !== undefined ? { text } : {}),
      ...(createdDateTime !== undefined ? { createdDateTime } : {}),
    };
  });

  return { copilotConversation: { messages } };
}

/** Parse Microsoft Graph's event-stream without assuming network chunk boundaries. */
export async function* parseGraphSse(
  body: ReadableStream<Uint8Array>,
  options: GraphSseParserOptions = {},
): AsyncIterable<GraphChatStreamEvent> {
  const maxEventBytes = options.maxEventBytes ?? DEFAULT_MAX_EVENT_BYTES;
  const reader = body.getReader();
  const decoder = new TextDecoder("utf-8", { fatal: true });
  const encoder = new TextEncoder();
  let pendingLine = "";
  let dataLines: string[] = [];
  let eventBytes = 0;

  const dispatchEvent = (): GraphChatStreamEvent | undefined => {
    if (dataLines.length === 0) {
      eventBytes = 0;
      return undefined;
    }

    const data = dataLines.join("\n");
    dataLines = [];
    eventBytes = 0;

    let parsed: unknown;
    try {
      parsed = JSON.parse(data) as unknown;
    } catch {
      throw new GraphSseParseError("Microsoft Graph SSE event contained malformed JSON.");
    }
    return validateEvent(parsed);
  };

  const processLine = (rawLine: string): GraphChatStreamEvent | undefined => {
    const line = rawLine.endsWith("\r") ? rawLine.slice(0, -1) : rawLine;
    eventBytes += encoder.encode(`${line}\n`).byteLength;
    if (eventBytes > maxEventBytes) {
      throw new GraphSseParseError("Microsoft Graph SSE event exceeded the configured size limit.");
    }

    if (line === "") {
      return dispatchEvent();
    }
    if (line.startsWith(":")) {
      return undefined;
    }

    const separator = line.indexOf(":");
    const field = separator === -1 ? line : line.slice(0, separator);
    let value = separator === -1 ? "" : line.slice(separator + 1);
    if (value.startsWith(" ")) {
      value = value.slice(1);
    }
    if (field === "data") {
      dataLines.push(value);
    }
    return undefined;
  };

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      if (value.byteLength > 0) options.onActivity?.();

      pendingLine += decoder.decode(value, { stream: true });

      while (true) {
        const newline = pendingLine.indexOf("\n");
        if (newline === -1) break;
        const line = pendingLine.slice(0, newline);
        pendingLine = pendingLine.slice(newline + 1);
        const event = processLine(line);
        if (event) yield event;
      }

      if (encoder.encode(pendingLine).byteLength + eventBytes > maxEventBytes) {
        throw new GraphSseParseError("Microsoft Graph SSE event exceeded the configured size limit.");
      }
    }

    pendingLine += decoder.decode();
    if (encoder.encode(pendingLine).byteLength + eventBytes > maxEventBytes) {
      throw new GraphSseParseError("Microsoft Graph SSE event exceeded the configured size limit.");
    }
    if (pendingLine.length > 0 || dataLines.length > 0) {
      throw new GraphSseParseError("Microsoft Graph SSE stream ended with an incomplete event.");
    }
  } catch (error) {
    if (error instanceof GraphSseParseError) throw error;
    if (error instanceof TypeError) {
      throw new GraphSseParseError("Microsoft Graph SSE stream contained invalid UTF-8.");
    }
    throw error;
  } finally {
    reader.releaseLock();
  }
}
