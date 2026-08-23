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
import { GATEWAY_MODEL_ID, toOpenAICompletion, toOpenAIToolCompletion } from "./openai-mapper.js";
import {
  bufferedTextChunks,
  bufferedToolCallChunks,
  createStreamProjector,
  serializeSseData,
  StreamProjectionError,
} from "./openai-stream.js";
import { flattenMessages, flattenMessagesWithTools } from "./prompt-adapter.js";
import { createGatewayLogger, type GatewayLogger, type LogFields } from "./logger.js";
import {
  createToolProtocol,
  InvalidToolDefinitionError,
  ToolProtocolError,
  validateToolDefinitions,
  type ToolDecision,
  type ToolProtocol,
} from "./tool-protocol.js";
import type {
  GraphChatResponse,
  GraphChatStreamEvent,
  GraphConversation,
  OpenAIChatCompletion,
  OpenAIChatCompletionChunk,
  OpenAIFunctionTool,
  OpenAIFunctionToolCall,
  OpenAIToolChoice,
} from "./types.js";

const functionToolCallSchema = z.object({
  id: z.string().min(1),
  type: z.literal("function"),
  function: z.object({
    name: z.string().min(1),
    arguments: z.string(),
  }),
});

const chatMessageSchema = z.discriminatedUnion("role", [
  z.object({
    role: z.enum(["system", "user"]),
    content: z.string().min(1),
    name: z.string().min(1).optional(),
  }),
  z.object({
    role: z.literal("assistant"),
    content: z.string().nullable().optional().default(null),
    name: z.string().min(1).optional(),
    tool_calls: z.array(functionToolCallSchema).max(1).optional(),
  }).superRefine((message, context) => {
    if ((message.content === null || message.content === "") && !message.tool_calls?.length) {
      context.addIssue({ code: "custom", message: "Assistant messages require content or one tool call." });
    }
  }),
  z.object({
    role: z.literal("tool"),
    content: z.string(),
    tool_call_id: z.string().min(1),
  }),
]);

const functionToolSchema = z.object({
  type: z.literal("function"),
  function: z.object({
    name: z.string().min(1),
    description: z.string().min(1).optional(),
    parameters: z.record(z.string(), z.unknown()).default({ type: "object", properties: {} }),
    strict: z.boolean().optional(),
  }),
});

const toolChoiceSchema = z.union([
  z.enum(["none", "auto", "required"]),
  z.object({
    type: z.literal("function"),
    function: z.object({ name: z.string().min(1) }),
  }),
]);

const chatRequestSchema = z.object({
  model: z.string().min(1),
  messages: z.array(chatMessageSchema).min(1),
  stream: z.boolean().optional(),
  stream_options: z.object({ include_usage: z.boolean().optional() }).optional(),
  n: z.literal(1).optional(),
  tools: z.array(functionToolSchema).max(64).optional(),
  tool_choice: toolChoiceSchema.optional(),
  parallel_tool_calls: z.boolean().optional(),
});

export interface GatewayDependencies {
  config: GatewayConfig;
  auth: AuthService;
  copilot: CopilotClient;
  logger?: GatewayLogger;
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
  if (error instanceof InvalidToolDefinitionError) {
    return reply.code(400).send({
      error: { message: error.message, type: "invalid_request_error", code: "invalid_tool_definition" },
    });
  }
  if (error instanceof ToolProtocolError) {
    return reply.code(502).send({
      error: {
        message: "Microsoft 365 Copilot did not produce a valid tool call response.",
        type: "api_error",
        code: "tool_protocol_error",
      },
    });
  }
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

function waitForDrain(response: ServerResponse, signal: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    const cleanup = (): void => {
      response.removeListener("drain", onDrain);
      response.removeListener("close", onClose);
      response.removeListener("error", onError);
      signal.removeEventListener("abort", onAbort);
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
    const onAbort = (): void => {
      cleanup();
      reject(new StreamClientDisconnectedError());
    };
    if (signal.aborted) {
      onAbort();
      return;
    }
    response.once("drain", onDrain);
    response.once("close", onClose);
    response.once("error", onError);
    signal.addEventListener("abort", onAbort, { once: true });
  });
}

export async function writeSse(response: ServerResponse, data: string, signal: AbortSignal): Promise<void> {
  if (signal.aborted || response.destroyed || response.writableEnded) {
    throw new StreamClientDisconnectedError();
  }
  if (!response.write(data)) {
    await waitForDrain(response, signal);
  }
}

type PendingStreamResult =
  | { kind: "result"; result: IteratorResult<GraphChatStreamEvent> }
  | { kind: "error"; error: unknown }
  | { kind: "heartbeat" }
  | { kind: "activity" }
  | { kind: "timeout" };

interface IdleDeadline {
  start(): void;
  activity(): void;
  promise(): Promise<Pick<PendingStreamResult, "kind"> & ({ kind: "activity" } | { kind: "timeout" })>;
  stop(): void;
}

function createIdleDeadline(timeoutMs: number): IdleDeadline {
  let active = false;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let resolveCurrent: ((result: { kind: "activity" } | { kind: "timeout" }) => void) | undefined;
  let current = new Promise<{ kind: "activity" } | { kind: "timeout" }>(() => undefined);

  const reset = (): void => {
    if (!active) return;
    if (timer !== undefined) clearTimeout(timer);
    resolveCurrent?.({ kind: "activity" });
    current = new Promise((resolve) => {
      resolveCurrent = resolve;
      timer = setTimeout(() => resolve({ kind: "timeout" }), timeoutMs);
    });
  };

  return {
    start() {
      active = true;
      reset();
    },
    activity: reset,
    promise() {
      return current;
    },
    stop() {
      active = false;
      if (timer !== undefined) clearTimeout(timer);
      timer = undefined;
      resolveCurrent = undefined;
    },
  };
}

async function nextStreamEvent(
  iterator: AsyncIterator<GraphChatStreamEvent>,
  config: GatewayConfig,
  writeHeartbeat: () => Promise<void>,
  idleDeadline: IdleDeadline,
): Promise<IteratorResult<GraphChatStreamEvent>> {
  idleDeadline.start();
  const pending: Promise<PendingStreamResult> = iterator.next().then(
    (result) => ({ kind: "result", result }),
    (error: unknown) => ({ kind: "error", error }),
  );
  try {
    while (true) {
      let heartbeatTimer: ReturnType<typeof setTimeout> | undefined;
      const heartbeat = config.gatewaySseHeartbeatMs > 0
        ? new Promise<PendingStreamResult>((resolve) => {
          heartbeatTimer = setTimeout(() => resolve({ kind: "heartbeat" }), config.gatewaySseHeartbeatMs);
        })
        : new Promise<PendingStreamResult>(() => undefined);

      const outcome = await Promise.race([pending, heartbeat, idleDeadline.promise()]);
      if (heartbeatTimer !== undefined) clearTimeout(heartbeatTimer);

      if (outcome.kind === "heartbeat") {
        await writeHeartbeat();
        continue;
      }
      if (outcome.kind === "activity") {
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
    idleDeadline.stop();
  }
}

function streamErrorCode(error: unknown): string {
  if (error instanceof StreamIdleTimeoutError) return "stream_idle_timeout";
  if (error instanceof StreamProjectionError || error instanceof GraphSseParseError) return "stream_protocol_error";
  return "stream_upstream_error";
}

function gatewayErrorCode(error: unknown): string {
  if (error instanceof InvalidToolDefinitionError) return "invalid_tool_definition";
  if (error instanceof ToolProtocolError) return "tool_protocol_error";
  if (isAuthenticationRequiredError(error)) return "m365_login_required";
  if (error instanceof GraphCopilotError) {
    return error.statusCode === 429 ? "rate_limit_exceeded" : `graph_${error.statusCode}`;
  }
  if (error instanceof StreamIdleTimeoutError) return "stream_idle_timeout";
  if (error instanceof StreamProjectionError || error instanceof GraphSseParseError) return "stream_protocol_error";
  return "gateway_upstream_error";
}

function errorLogFields(error: unknown): LogFields {
  return {
    code: gatewayErrorCode(error),
    error_type: error instanceof Error ? error.name : typeof error,
    ...(error instanceof GraphCopilotError ? {
      upstream_status: error.statusCode,
      ...(error.retryAfter !== undefined ? { retry_after: error.retryAfter } : {}),
    } : {}),
    ...(error instanceof ToolProtocolError ? { protocol_reason: error.reason } : {}),
  };
}

function elapsedMs(startedAt: number): number {
  return Date.now() - startedAt;
}

function requestId(request: FastifyRequest): string {
  return String(request.id);
}

function latestGraphText(response: GraphChatResponse): string {
  for (let index = (response.messages?.length ?? 0) - 1; index >= 0; index -= 1) {
    const text = response.messages?.[index]?.text;
    if (typeof text === "string" && text.trim() !== "") return text;
  }
  throw new ToolProtocolError("invalid_shape");
}

function normalizeTools(tools: z.infer<typeof functionToolSchema>[]): OpenAIFunctionTool[] {
  return tools.map((tool) => ({
    type: "function",
    function: {
      name: tool.function.name,
      ...(tool.function.description !== undefined ? { description: tool.function.description } : {}),
      parameters: tool.function.parameters,
      ...(tool.function.strict !== undefined ? { strict: tool.function.strict } : {}),
    },
  }));
}

function createFunctionCall(decision: Extract<ToolDecision, { kind: "tool_call" }>): OpenAIFunctionToolCall {
  return {
    id: `call_${crypto.randomUUID().replaceAll("-", "")}`,
    type: "function",
    function: {
      name: decision.name,
      arguments: JSON.stringify(decision.arguments),
    },
  };
}

async function resolveToolDecision(
  dependencies: GatewayDependencies,
  accessToken: string,
  conversation: GraphConversation,
  initialPrompt: string,
  protocol: ToolProtocol,
  signal: AbortSignal,
  logger: GatewayLogger,
  requestIdentifier: string,
): Promise<{ decision: ToolDecision; correctionAttempted: boolean }> {
  const firstResponse = await dependencies.copilot.chat(accessToken, conversation.id, initialPrompt, signal);
  try {
    return { decision: protocol.parse(latestGraphText(firstResponse)), correctionAttempted: false };
  } catch (error) {
    if (!(error instanceof ToolProtocolError)) throw error;
    logger.warn("tool_protocol_retry", {
      request_id: requestIdentifier,
      protocol_reason: error.reason,
    });
    const corrected = await dependencies.copilot.chat(
      accessToken,
      conversation.id,
      protocol.correctionPrompt(error),
      signal,
    );
    return { decision: protocol.parse(latestGraphText(corrected)), correctionAttempted: true };
  }
}

async function toolChatCompletion(
  dependencies: GatewayDependencies,
  input: z.infer<typeof chatRequestSchema>,
  tools: OpenAIFunctionTool[],
  choice: Exclude<OpenAIToolChoice, "none">,
  request: FastifyRequest,
  reply: FastifyReply,
  activeStreams: Set<() => void>,
  logger: GatewayLogger,
  startedAt: number,
): Promise<FastifyReply | void> {
  const id = requestId(request);
  const controller = new AbortController();
  const abort = (): void => controller.abort();
  const timeout = setTimeout(abort, dependencies.config.graphStreamIdleTimeoutMs);
  let downstreamStarted = false;
  activeStreams.add(abort);
  request.raw.once("aborted", abort);

  try {
    const protocol = createToolProtocol(tools, choice, crypto.randomUUID());
    const token = await dependencies.auth.getAccessToken();
    logger.debug("auth_succeeded", { request_id: id, duration_ms: elapsedMs(startedAt) });
    const conversation = await dependencies.copilot.createConversation(token, controller.signal);
    logger.debug("graph_conversation_created", { request_id: id, duration_ms: elapsedMs(startedAt) });
    const prompt = flattenMessagesWithTools(input.messages, tools, choice, protocol.nonce);
    const result = await resolveToolDecision(
      dependencies,
      token,
      conversation,
      prompt,
      protocol,
      controller.signal,
      logger,
      id,
    );
    clearTimeout(timeout);

    const decision = result.decision;
    const outputChars = decision.kind === "final"
      ? decision.content.length
      : JSON.stringify(decision.arguments).length;
    let completion: OpenAIChatCompletion;
    let chunks: OpenAIChatCompletionChunk[];
    if (decision.kind === "tool_call") {
      const toolCall = createFunctionCall(decision);
      completion = toOpenAIToolCompletion(conversation, toolCall);
      chunks = bufferedToolCallChunks(conversation, toolCall);
    } else {
      completion = toOpenAICompletion(conversation, { messages: [{ text: decision.content }] });
      chunks = bufferedTextChunks(conversation, decision.content);
    }
    if (input.stream !== true) {
      logger.info("request_completed", {
        request_id: id,
        mode: "sync_tool",
        status: 200,
        duration_ms: elapsedMs(startedAt),
        outcome: decision.kind,
        correction_attempted: result.correctionAttempted,
        output_chars: outputChars,
      });
      return reply.send(completion);
    }

    reply.hijack();
    reply.raw.writeHead(200, {
      "Content-Type": "text/event-stream; charset=utf-8",
      "Cache-Control": "no-cache, no-transform",
      Connection: "keep-alive",
      "X-Accel-Buffering": "no",
      "X-Request-Id": id,
    });
    downstreamStarted = true;
    for (const chunk of chunks) {
      await writeSse(reply.raw, serializeSseData(chunk), controller.signal);
    }
    await writeSse(reply.raw, serializeSseData("[DONE]"), controller.signal);
    reply.raw.end();
    logger.info("request_completed", {
      request_id: id,
      mode: "stream_tool",
      status: 200,
      duration_ms: elapsedMs(startedAt),
      outcome: decision.kind,
      correction_attempted: result.correctionAttempted,
      output_chars: outputChars,
    });
    return reply;
  } catch (error) {
    logger.error("request_failed", {
      request_id: id,
      mode: input.stream === true ? "stream_tool" : "sync_tool",
      duration_ms: elapsedMs(startedAt),
      ...errorLogFields(error),
    });
    if (downstreamStarted) {
      if (!reply.raw.destroyed) reply.raw.destroy();
      return;
    }
    return sendGatewayError(reply, error);
  } finally {
    clearTimeout(timeout);
    request.raw.removeListener("aborted", abort);
    activeStreams.delete(abort);
  }
}

async function streamChatCompletion(
  dependencies: GatewayDependencies,
  input: z.infer<typeof chatRequestSchema>,
  request: FastifyRequest,
  reply: FastifyReply,
  activeStreams: Set<() => void>,
  logger: GatewayLogger,
  startedAt: number,
): Promise<FastifyReply | void> {
  const id = requestId(request);
  const controller = new AbortController();
  const downstreamController = new AbortController();
  const idleDeadline = createIdleDeadline(dependencies.config.graphStreamIdleTimeoutMs);
  let clientDisconnected = false;
  let downstreamStarted = false;
  let completed = false;
  let startTimer: ReturnType<typeof setTimeout> | undefined;
  let rejectStart: ((error: unknown) => void) | undefined;
  let iterator: AsyncIterator<GraphChatStreamEvent> | undefined;
  let snapshotCount = 0;
  let contentDeltaCount = 0;
  let outputChars = 0;
  let firstContentMs: number | undefined;
  const abortForCancellation = (): void => {
    controller.abort();
    downstreamController.abort();
    rejectStart?.(new StreamClientDisconnectedError());
    if (!reply.raw.destroyed) reply.raw.destroy();
  };
  activeStreams.add(abortForCancellation);

  const onClientDisconnect = (): void => {
    clientDisconnected = true;
    abortForCancellation();
  };
  const cleanup = (): void => {
    request.raw.removeListener("aborted", onClientDisconnect);
    reply.raw.removeListener("close", onClientDisconnect);
    if (startTimer !== undefined) clearTimeout(startTimer);
    idleDeadline.stop();
    activeStreams.delete(abortForCancellation);
  };

  request.raw.once("aborted", onClientDisconnect);
  reply.raw.once("close", onClientDisconnect);

  try {
    const startTimeout = new Promise<never>((_, reject) => {
      startTimer = setTimeout(() => {
        controller.abort();
        reject(new StreamStartTimeoutError());
      }, dependencies.config.graphStreamStartTimeoutMs);
    });
    const disconnected = new Promise<never>((_, reject) => {
      rejectStart = reject;
    });
    const token = await Promise.race([
      dependencies.auth.getAccessToken(),
      startTimeout,
      disconnected,
    ]);
    logger.debug("auth_succeeded", { request_id: id, duration_ms: elapsedMs(startedAt) });
    if (clientDisconnected) return;
    const conversation = await Promise.race([
      dependencies.copilot.createConversation(token, controller.signal),
      startTimeout,
      disconnected,
    ]);
    logger.debug("graph_conversation_created", { request_id: id, duration_ms: elapsedMs(startedAt) });
    if (clientDisconnected) return;
    const graphStream = await Promise.race([
      dependencies.copilot.chatStream(
        token,
        conversation.id,
        flattenMessages(input.messages),
        controller.signal,
        () => idleDeadline.activity(),
      ),
      startTimeout,
      disconnected,
    ]);
    logger.debug("graph_stream_ready", { request_id: id, duration_ms: elapsedMs(startedAt) });
    if (startTimer !== undefined) clearTimeout(startTimer);
    if (clientDisconnected) return;

    reply.hijack();
    reply.raw.writeHead(200, {
      "Content-Type": "text/event-stream; charset=utf-8",
      "Cache-Control": "no-cache, no-transform",
      Connection: "keep-alive",
      "X-Accel-Buffering": "no",
      "X-Request-Id": id,
    });
    downstreamStarted = true;

    const projector = createStreamProjector(conversation);
    const roleChunk = projector.roleChunk();
    if (roleChunk) {
      await writeSse(reply.raw, serializeSseData(roleChunk), downstreamController.signal);
    }

    iterator = graphStream[Symbol.asyncIterator]();
    while (true) {
      const result = await nextStreamEvent(
        iterator,
        dependencies.config,
        async () => {
          await writeSse(reply.raw, ": keep-alive\n\n", downstreamController.signal);
        },
        idleDeadline,
      );
      if (result.done) break;
      snapshotCount += 1;
      logger.trace("stream_snapshot", {
        request_id: id,
        snapshot: snapshotCount,
        message_count: result.value.copilotConversation.messages?.length ?? 0,
      });
      const contentChunk = projector.contentChunk(result.value.copilotConversation);
      if (contentChunk) {
        const deltaChars = contentChunk.choices[0]?.delta.content?.length ?? 0;
        contentDeltaCount += 1;
        outputChars += deltaChars;
        firstContentMs ??= elapsedMs(startedAt);
        logger.trace("stream_delta", {
          request_id: id,
          snapshot: snapshotCount,
          delta_chars: deltaChars,
          output_chars: outputChars,
        });
        await writeSse(reply.raw, serializeSseData(contentChunk), downstreamController.signal);
      }
    }

    await writeSse(reply.raw, serializeSseData(projector.finalChunk()), downstreamController.signal);
    await writeSse(reply.raw, serializeSseData("[DONE]"), downstreamController.signal);
    logger.info("request_completed", {
      request_id: id,
      mode: "stream",
      status: 200,
      duration_ms: elapsedMs(startedAt),
      first_content_ms: firstContentMs ?? null,
      upstream_snapshots: snapshotCount,
      content_deltas: contentDeltaCount,
      output_chars: outputChars,
    });
    completed = true;
    cleanup();
    reply.raw.end();
    return reply;
  } catch (error) {
    if (clientDisconnected || error instanceof StreamClientDisconnectedError) {
      logger.debug("stream_cancelled", {
        request_id: id,
        mode: "stream",
        duration_ms: elapsedMs(startedAt),
        reason: clientDisconnected ? "client_disconnected" : "downstream_disconnected",
        upstream_snapshots: snapshotCount,
        output_chars: outputChars,
      });
      return;
    }
    logger.error(downstreamStarted ? "stream_failed" : "request_failed", {
      request_id: id,
      mode: "stream",
      duration_ms: elapsedMs(startedAt),
      upstream_snapshots: snapshotCount,
      output_chars: outputChars,
      ...errorLogFields(error),
    });
    if (!downstreamStarted) {
      return sendGatewayError(reply, error);
    }

    controller.abort();
    const errorWriteTimer = setTimeout(() => {
      downstreamController.abort();
      if (!reply.raw.destroyed) reply.raw.destroy();
    }, Math.min(1_000, dependencies.config.graphStreamIdleTimeoutMs));
    try {
      await writeSse(reply.raw, serializeSseData({
        error: {
          message: "Microsoft Graph Copilot streaming failed before completion.",
          type: "api_error",
          code: streamErrorCode(error),
        },
      }), downstreamController.signal);
    } catch {
      // The downstream connection may have closed while the error was being written.
    } finally {
      clearTimeout(errorWriteTimer);
    }
    cleanup();
    reply.raw.end();
    return reply;
  } finally {
    cleanup();
    if (!completed) {
      controller.abort();
      downstreamController.abort();
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
  const logger = dependencies.logger ?? createGatewayLogger(
    dependencies.config.logLevel,
    dependencies.config.logFormat,
  );
  const activeStreams = new Set<() => void>();

  app.addHook("preClose", async () => {
    for (const abort of activeStreams) abort();
  });

  app.addHook("onRequest", async (request, reply) => {
    if (dependencies.config.apiKey && !apiKeyIsValid(request, dependencies.config.apiKey)) {
      reply.header("x-request-id", requestId(request));
      logger.warn("request_rejected", {
        request_id: requestId(request),
        method: request.method,
        reason: "invalid_api_key",
      });
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
    const id = requestId(request);
    const startedAt = Date.now();
    reply.header("x-request-id", id);
    const parsed = chatRequestSchema.safeParse(request.body);
    if (!parsed.success) {
      logger.warn("request_rejected", {
        request_id: id,
        method: request.method,
        reason: "invalid_request",
      });
      return reply.code(400).send({
        error: { message: "Invalid chat completion request.", type: "invalid_request_error", code: "invalid_request" },
      });
    }

    const input = parsed.data;
    logger.info("request_started", {
      request_id: id,
      mode: input.stream === true ? "stream" : "sync",
      model: input.model,
      message_count: input.messages.length,
      tools_requested: Boolean(input.tools && input.tools.length > 0),
    });
    const tools = normalizeTools(input.tools ?? []);
    const choice = input.tool_choice ?? "auto";
    if (tools.length === 0 && choice !== "auto" && choice !== "none") {
      const error = new InvalidToolDefinitionError("tool_choice requires at least one function tool.");
      logger.warn("request_rejected", {
        request_id: id,
        method: request.method,
        reason: "invalid_tool_definition",
      });
      return sendGatewayError(reply, error);
    }
    if (tools.length > 0 && choice === "none") {
      try {
        validateToolDefinitions(tools);
      } catch (error) {
        logger.warn("request_rejected", {
          request_id: id,
          method: request.method,
          reason: "invalid_tool_definition",
        });
        return sendGatewayError(reply, error);
      }
    }
    if (tools.length > 0 && choice !== "none") {
      return toolChatCompletion(
        dependencies,
        input,
        tools,
        choice,
        request,
        reply,
        activeStreams,
        logger,
        startedAt,
      );
    }

    if (input.stream === true) {
      return streamChatCompletion(dependencies, input, request, reply, activeStreams, logger, startedAt);
    }

    try {
      const token = await dependencies.auth.getAccessToken();
      logger.debug("auth_succeeded", { request_id: id, duration_ms: elapsedMs(startedAt) });
      const conversation = await dependencies.copilot.createConversation(token);
      logger.debug("graph_conversation_created", { request_id: id, duration_ms: elapsedMs(startedAt) });
      const graphResponse = await dependencies.copilot.chat(token, conversation.id, flattenMessages(input.messages));
      const completion = toOpenAICompletion(conversation, graphResponse);
      const outputChars = graphResponse.messages?.reduce((total, message) => total + (message.text?.length ?? 0), 0) ?? 0;
      logger.info("request_completed", {
        request_id: id,
        mode: "sync",
        status: 200,
        duration_ms: elapsedMs(startedAt),
        output_chars: outputChars,
      });
      return reply.send(completion);
    } catch (error) {
      logger.error("request_failed", {
        request_id: id,
        mode: "sync",
        duration_ms: elapsedMs(startedAt),
        ...errorLogFields(error),
      });
      return sendGatewayError(reply, error);
    }
  });

  return app;
}

async function start(): Promise<void> {
  const config = loadConfig();
  const logger = createGatewayLogger(config.logLevel, config.logFormat);
  logger.info("gateway_starting", {
    host: config.host,
    port: config.port,
    log_level: config.logLevel,
    log_format: config.logFormat,
  });
  const auth = await createAuthService(config);
  const copilot = createCopilotClient(config);
  const app = buildServer({ config, auth, copilot, logger });
  await app.listen({ host: config.host, port: config.port });
  logger.info("gateway_listening", { host: config.host, port: config.port });
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  void start();
}
