import { homedir } from "node:os";
import path from "node:path";
import { z } from "zod";

const environmentSchema = z.object({
  M365_TENANT_ID: z.string().uuid(),
  M365_CLIENT_ID: z.string().uuid(),
  GATEWAY_HOST: z.string().default("127.0.0.1"),
  GATEWAY_PORT: z.coerce.number().int().min(1).max(65535).default(8787),
  GATEWAY_API_KEY: z.string().optional(),
  M365_TIME_ZONE: z.string().min(1).default("Australia/Sydney"),
  M365_TOKEN_CACHE_DIR: z.string().optional().transform((value) => value === "" ? undefined : value),
  GRAPH_BASE_URL: z.url().default("https://graph.microsoft.com/beta"),
  GRAPH_STREAM_START_TIMEOUT_MS: z.coerce.number().int().min(1).default(30_000),
  GRAPH_STREAM_IDLE_TIMEOUT_MS: z.coerce.number().int().min(1).default(60_000),
  GRAPH_STREAM_MAX_EVENT_BYTES: z.coerce.number().int().min(1_024).default(2 * 1024 * 1024),
  GATEWAY_SSE_HEARTBEAT_MS: z.coerce.number().int().min(0).default(15_000),
});

export interface GatewayConfig {
  tenantId: string;
  clientId: string;
  host: string;
  port: number;
  apiKey?: string;
  timeZone: string;
  tokenCacheDirectory: string;
  graphBaseUrl: string;
  graphStreamStartTimeoutMs: number;
  graphStreamIdleTimeoutMs: number;
  graphStreamMaxEventBytes: number;
  gatewaySseHeartbeatMs: number;
}

export function loadConfig(environment: NodeJS.ProcessEnv = process.env): GatewayConfig {
  const result = environmentSchema.safeParse(environment);
  if (!result.success) {
    const details = result.error.issues.map((issue) => `${issue.path.join(".")}: ${issue.message}`).join("; ");
    throw new Error(`Invalid gateway configuration: ${details}`);
  }

  const env = result.data;
  return {
    tenantId: env.M365_TENANT_ID,
    clientId: env.M365_CLIENT_ID,
    host: env.GATEWAY_HOST,
    port: env.GATEWAY_PORT,
    ...(env.GATEWAY_API_KEY ? { apiKey: env.GATEWAY_API_KEY } : {}),
    timeZone: env.M365_TIME_ZONE,
    tokenCacheDirectory:
      env.M365_TOKEN_CACHE_DIR ?? path.join(homedir(), ".m365-copilot-openai-gateway", "msal-cache"),
    graphBaseUrl: env.GRAPH_BASE_URL.replace(/\/$/, ""),
    graphStreamStartTimeoutMs: env.GRAPH_STREAM_START_TIMEOUT_MS,
    graphStreamIdleTimeoutMs: env.GRAPH_STREAM_IDLE_TIMEOUT_MS,
    graphStreamMaxEventBytes: env.GRAPH_STREAM_MAX_EVENT_BYTES,
    gatewaySseHeartbeatMs: env.GATEWAY_SSE_HEARTBEAT_MS,
  };
}
