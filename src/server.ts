import "dotenv/config";
import crypto from "node:crypto";
import type { ServerResponse } from "node:http";
import { fileURLToPath } from "node:url";
import Fastify, { type FastifyInstance, type FastifyReply, type FastifyRequest } from "fastify";
import { z } from "zod";
import { createAuthService, isAuthenticationRequiredError, type AuthService } from "./auth.js";
import { loadConfig, type GatewayConfig } from "./config.js";
import { createCopilotClient, GraphCopilotError, type CopilotClient } from "./graph-copilot.js";
import { GraphSseParseError } from "./graph-sse-parser.js";
import { GATEWAY_MODEL_ID, toOpenAICompletion } from "./openai-mapper.js";
import {
  createStreamProjector,
  serializeSseData,
  StreamProjectionError,
} from "./openai-stream.js";
import { flattenMessages } from "./prompt-adapter.js";
import type { GraphChatStreamEvent } from "./types.js";

const chatRequestSchema = z.object({
  model: z.string().min(1),
  messages: z.array(z.object({
    role: z.enum(["system", "user", "assistant", "tool"]),
    content: z.string().min(1),
    tool_call_id: z.string().optional(),
  })).min(1),
  stream: z.boolean().optional(),
  stream_options: z.object({ include_usage: z.boolean().optional() }).optional(),
  n: z.literal(1).optional(),
  tools: z.array(z.unknown()).optional(),
});

export interface GatewayDependencies {
  config: GatewayConfig;
  auth: AuthService;
  copilot: CopilotClient;
}

function apiKeyIsValid(request: FastifyRequest, apiKey: string): boolean {
  const authorization = request.headers.authorization;
  if (!authorization?.startsWith("Bearer ")) return false;
  const expected = Buffer.from(apiKey);
  const received = Buffer.from(authorization.slice("Bearer ".length));
  return expected.length === received.length && crypto.timingSafeEqual(expected, received);
}

class StreamStartTimeoutError extends Error {
  public constructor() {
    super("Microsoft Graph streaming response did not start before the configured timeout.");
    this.name = "StreamStartTimeoutError";
  }
}

class StreamIdleTimeoutError extends Error {
  public constructor() {
    super("Microsoft Graph streaming response exceeded the configured idle timeout.");
    this.name = "StreamIdleTimeoutError";
  }
}

class StreamClientDisconnectedError extends Error {
  public constructor() {
    super("The downstream client disconnected.");
    this.name = "StreamClientDisconnectedError";
  }
}

function sendGatewayError(reply: FastifyReply, error: unknown): FastifyReply {
  if (isAuthenticationRequiredError(error)) {
    return reply.code(401).send({
      error: { message: error.message, type: "authentication_error", code: "m365_login_required" },
    });
  }
  if (error instanceof GraphCopilotError) {
    if (error.statusCode === 429) {
      if (error.retryAfter !== undefined) {
        reply.header("retry-after", error.retryAfter);
      }
      return reply.code(429).send({
        error: { message: "Microsoft Graph Copilot is rate limiting requests.", type: "rate_limit_error", code: "rate_limit_exceeded" },
      });
    }
    return reply.code(502).send({
      error: { message: "Microsoft Graph Copilot request failed.", type: "api_error", code: `graph_${error.statusCode}` },
    });
  }
  return reply.code(502).send({
    error: { message: "Gateway could not complete the Copilot request.", type: "api_error", code: "gateway_upstream_error" },
  });
}

function waitForDrain(response: ServerResponse): Promise<void> {
  return new Promise((resolve, reject) => {
    const cleanup = (): void => {
      response.removeListener("drain", onDrain);
      response.removeListener("close", onClose);
      response.removeListener("error", onError);
    };
    const onDrain = (): void => {
      cleanup();
      resolve();
    };
    const onClose = (): void => {
      cleanup();
      reject(new StreamClientDisconnectedError());
    };
    const onError = (): void => {
      cleanup();
      reject(new StreamClientDisconnectedError());
    };
    response.once("drain", onDrain);
    response.once("close", onClose);
    response.once("error", onError);
  });
}

async function writeSse(response: ServerResponse, data: string): Promise<void> {
  if (response.destroyed || response.writableEnded) {
    throw new StreamClientDisconnectedError();
  }
  if (!response.write(data)) {
    await waitForDrain(response);
  }
}

type PendingStreamResult =
  | { kind: "result"; result: IteratorResult<GraphChatStreamEvent> }
  | { kind: "error"; error: unknown }
  | { kind: "heartbeat" }
  | { kind: "timeout" };

async function nextStreamEvent(
  iterator: AsyncIterator<GraphChatStreamEvent>,
  config: GatewayConfig,
  writeHeartbeat: () => Promise<void>,
): Promise<IteratorResult<GraphChatStreamEvent>> {
  const pending: Promise<PendingStreamResult> = iterator.next().then(
    (result) => ({ kind: "result", result }),
    (error: unknown) => ({ kind: "error", error }),
  );
  let idleTimer: ReturnType<typeof setTimeout> | undefined;
  const idle = new Promise<PendingStreamResult>((resolve) => {
    idleTimer = setTimeout(() => resolve({ kind: "timeout" }), config.graphStreamIdleTimeoutMs);
  });

  try {
    while (true) {
      let heartbeatTimer: ReturnType<typeof setTimeout> | undefined;
      const heartbeat = config.gatewaySseHeartbeatMs > 0
        ? new Promise<PendingStreamResult>((resolve) => {
          heartbeatTimer = setTimeout(() => resolve({ kind: "heartbeat" }), config.gatewaySseHeartbeatMs);
        })
        : new Promise<PendingStreamResult>(() => undefined);

      const outcome = await Promise.race([pending, heartbeat, idle]);
      if (heartbeatTimer !== undefined) clearTimeout(heartbeatTimer);

      if (outcome.kind === "heartbeat") {
        await writeHeartbeat();
        continue;
      }
      if (outcome.kind === "timeout") {
        throw new StreamIdleTimeoutError();
      }
      if (outcome.kind === "error") {
        throw outcome.error;
      }
      return outcome.result;
    }
  } finally {
    if (idleTimer !== undefined) clearTimeout(idleTimer);
  }
}

function streamErrorCode(error: unknown): string {
  if (error instanceof StreamIdleTimeoutError) return "stream_idle_timeout";
  if (error instanceof StreamProjectionError || error instanceof GraphSseParseError) return "stream_protocol_error";
  return "stream_upstream_error";
}

async function streamChatCompletion(
  dependencies: GatewayDependencies,
  input: z.infer<typeof chatRequestSchema>,
  request: FastifyRequest,
  reply: FastifyReply,
  activeStreams: Set<AbortController>,
): Promise<FastifyReply | void> {
  const controller = new AbortController();
  activeStreams.add(controller);
  let clientDisconnected = false;
  let downstreamStarted = false;
  let completed = false;
  let startTimer: ReturnType<typeof setTimeout> | undefined;
  let rejectStart: ((error: unknown) => void) | undefined;
  let iterator: AsyncIterator<GraphChatStreamEvent> | undefined;

  const onClientDisconnect = (): void => {
    clientDisconnected = true;
    controller.abort();
    rejectStart?.(new StreamClientDisconnectedError());
  };
  const cleanup = (): void => {
    request.raw.removeListener("aborted", onClientDisconnect);
    reply.raw.removeListener("close", onClientDisconnect);
    if (startTimer !== undefined) clearTimeout(startTimer);
    activeStreams.delete(controller);
  };

  request.raw.once("aborted", onClientDisconnect);
  reply.raw.once("close", onClientDisconnect);

  try {
    const token = await dependencies.auth.getAccessToken();
    if (clientDisconnected) return;
    const conversation = await dependencies.copilot.createConversation(token);
    if (clientDisconnected) return;
    const startTimeout = new Promise<never>((_, reject) => {
      startTimer = setTimeout(() => {
        controller.abort();
        reject(new StreamStartTimeoutError());
      }, dependencies.config.graphStreamStartTimeoutMs);
    });
    const disconnected = new Promise<never>((_, reject) => {
      rejectStart = reject;
    });
    const graphStream = await Promise.race([
      dependencies.copilot.chatStream(token, conversation.id, flattenMessages(input.messages), controller.signal),
      startTimeout,
      disconnected,
    ]);
    if (startTimer !== undefined) clearTimeout(startTimer);
    if (clientDisconnected) return;

    reply.hijack();
    reply.raw.writeHead(200, {
      "Content-Type": "text/event-stream; charset=utf-8",
      "Cache-Control": "no-cache, no-transform",
      Connection: "keep-alive",
      "X-Accel-Buffering": "no",
    });
    downstreamStarted = true;

    const projector = createStreamProjector(conversation);
    const roleChunk = projector.roleChunk();
    if (roleChunk) {
      await writeSse(reply.raw, serializeSseData(roleChunk));
    }

    iterator = graphStream[Symbol.asyncIterator]();
    while (true) {
      const result = await nextStreamEvent(iterator, dependencies.config, async () => {
        await writeSse(reply.raw, ": keep-alive\n\n");
      });
      if (result.done) break;
      const contentChunk = projector.contentChunk(result.value.copilotConversation);
      if (contentChunk) {
        await writeSse(reply.raw, serializeSseData(contentChunk));
      }
    }

    await writeSse(reply.raw, serializeSseData(projector.finalChunk()));
    await writeSse(reply.raw, serializeSseData("[DONE]"));
    completed = true;
    cleanup();
    reply.raw.end();
    return reply;
  } catch (error) {
    if (clientDisconnected || error instanceof StreamClientDisconnectedError) {
      return;
    }
    if (!downstreamStarted) {
      return sendGatewayError(reply, error);
    }

    controller.abort();
    try {
      await writeSse(reply.raw, serializeSseData({
        error: {
          message: "Microsoft Graph Copilot streaming failed before completion.",
          type: "api_error",
          code: streamErrorCode(error),
        },
      }));
    } catch {
      // The downstream connection may have closed while the error was being written.
    }
    completed = true;
    cleanup();
    reply.raw.end();
    return reply;
  } finally {
    cleanup();
    if (!completed) {
      controller.abort();
      if (iterator?.return) {
        try {
          await iterator.return();
        } catch {
          // Cleanup must not replace the original request or stream error.
        }
      }
    }
  }
}

export function buildServer(dependencies: GatewayDependencies): FastifyInstance {
  const app = Fastify({ logger: false });
  const activeStreams = new Set<AbortController>();

  app.addHook("preClose", async () => {
    for (const controller of activeStreams) controller.abort();
  });

  app.addHook("onRequest", async (request, reply) => {
    if (dependencies.config.apiKey && !apiKeyIsValid(request, dependencies.config.apiKey)) {
      await reply.code(401).send({
        error: { message: "Invalid or missing gateway API key.", type: "authentication_error", code: "invalid_api_key" },
      });
    }
  });

  app.get("/health", async () => ({ status: "ok", service: "m365-copilot-openai-gateway" }));
  app.get("/v1/models", async () => ({
    object: "list",
    data: [{ id: GATEWAY_MODEL_ID, object: "model", created: 0, owned_by: "m365-copilot-openai-gateway" }],
  }));

  app.post("/v1/chat/completions", async (request, reply) => {
    const parsed = chatRequestSchema.safeParse(request.body);
    if (!parsed.success) {
      return reply.code(400).send({
        error: { message: "Invalid chat completion request.", type: "invalid_request_error", code: "invalid_request" },
      });
    }

    const input = parsed.data;
    if (input.tools && input.tools.length > 0) {
      return reply.code(400).send({
        error: {
          message: "OpenAI tool calling is planned but is not implemented in phase 1.",
          type: "invalid_request_error",
          code: "tools_not_supported",
        },
      });
    }

    if (input.stream === true) {
      return streamChatCompletion(dependencies, input, request, reply, activeStreams);
    }

    try {
      const token = await dependencies.auth.getAccessToken();
      const conversation = await dependencies.copilot.createConversation(token);
      const graphResponse = await dependencies.copilot.chat(token, conversation.id, flattenMessages(input.messages));
      return reply.send(toOpenAICompletion(conversation, graphResponse));
    } catch (error) {
      return sendGatewayError(reply, error);
    }
  });

  return app;
}

async function start(): Promise<void> {
  const config = loadConfig();
  const auth = await createAuthService(config);
  const copilot = createCopilotClient(config);
  const app = buildServer({ config, auth, copilot });
  await app.listen({ host: config.host, port: config.port });
  console.log(`Gateway listening at http://${config.host}:${config.port}`);
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  void start();
}
