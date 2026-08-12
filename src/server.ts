import "dotenv/config";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";
import Fastify, { type FastifyInstance, type FastifyRequest } from "fastify";
import { z } from "zod";
import { createAuthService, isAuthenticationRequiredError, type AuthService } from "./auth.js";
import { loadConfig, type GatewayConfig } from "./config.js";
import { createCopilotClient, GraphCopilotError, type CopilotClient } from "./graph-copilot.js";
import { GATEWAY_MODEL_ID, toOpenAICompletion } from "./openai-mapper.js";
import { flattenMessages } from "./prompt-adapter.js";

const chatRequestSchema = z.object({
  model: z.string().min(1),
  messages: z.array(z.object({
    role: z.enum(["system", "user", "assistant", "tool"]),
    content: z.string().min(1),
    tool_call_id: z.string().optional(),
  })).min(1),
  stream: z.boolean().optional(),
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

export function buildServer(dependencies: GatewayDependencies): FastifyInstance {
  const app = Fastify({ logger: false });

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
    if (input.stream) {
      return reply.code(400).send({
        error: {
          message: "Streaming is planned but is not implemented in phase 1.",
          type: "invalid_request_error",
          code: "streaming_not_supported",
        },
      });
    }
    if (input.tools && input.tools.length > 0) {
      return reply.code(400).send({
        error: {
          message: "OpenAI tool calling is planned but is not implemented in phase 1.",
          type: "invalid_request_error",
          code: "tools_not_supported",
        },
      });
    }

    try {
      const token = await dependencies.auth.getAccessToken();
      const conversation = await dependencies.copilot.createConversation(token);
      const graphResponse = await dependencies.copilot.chat(token, conversation.id, flattenMessages(input.messages));
      return reply.send(toOpenAICompletion(conversation, graphResponse));
    } catch (error) {
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
