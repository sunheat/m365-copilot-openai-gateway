# M365 Copilot OpenAI Gateway

Local TypeScript gateway that exposes a narrow OpenAI Chat Completions-compatible
surface over the Microsoft 365 Copilot Chat API preview.

The gateway is designed for a single signed-in developer on `127.0.0.1`. It
uses delegated Microsoft Entra authentication and never exposes Microsoft Graph
tokens to the IDE client.

## Gateway status

Implemented locally:

- `GET /health`
- `GET /v1/models` exposing `m365-copilot-preview`
- `POST /v1/chat/completions` for text and function-tool requests
- OpenAI-compatible `tools`, `tool_choice`, `message.tool_calls`, and streamed `delta.tool_calls`
- Microsoft Entra delegated browser sign-in using authorization code with PKCE and an encrypted, current-user MSAL token cache on Windows
- A new Graph Copilot conversation per request and web grounding disabled on every turn

The text-only streaming path calls Microsoft Graph's native `chatOverStream` endpoint, converts cumulative snapshots into OpenAI-compatible content deltas, and ends successful responses with a `finish_reason` chunk followed by `[DONE]`. Tool-enabled turns use a strict gateway-side text protocol because Graph does not expose native tool calling; those responses are validated and fully buffered before OpenAI-compatible output is emitted. `stream_options.include_usage` is accepted but no fabricated token usage is emitted because Graph does not provide token counts.

The proposed native Graph-to-OpenAI streaming architecture is documented in
[Phase 2 SSE Streaming Design](docs/phase-2-sse-streaming-design.md).

The implemented text-protocol bridge for IDE tool loops is documented in
[Phase 3 OpenAI Tool-Calling Compatibility Design](docs/phase-3-tool-calling-design.md).

## Local setup

1. In the Microsoft Entra app registration, add `http://localhost` as a **Mobile and desktop applications** redirect URI. Keep public client flows enabled.
2. Copy `.env.example` to `.env` and set `M365_TENANT_ID` and `M365_CLIENT_ID` from the app registration. Keep `GATEWAY_HOST=127.0.0.1`.
3. Run `npm install`.
4. Run `npm run auth:login`. Complete sign-in in the browser window that opens, including MFA if required.
5. Run `npm run dev`.

The initial browser sign-in uses authorization code with PKCE and creates the local encrypted cache. Later runs acquire and refresh the delegated access token silently; another login is only needed after sign-out, consent changes, cache deletion, or a revoked refresh token. This flow remains compatible with Microsoft Entra Security Defaults, which blocks device-code authentication.

## Diagnostics

The gateway writes concise diagnostic records to the dev console. Control the
verbosity with `GATEWAY_LOG_LEVEL`: `silent`, `error`, `warn`, `info`, `debug`,
or `trace`. The default is `info`. Set `GATEWAY_LOG_FORMAT=json` when a
structured log consumer or shell redirection is preferred; the default
`pretty` format is easier to read interactively.

Logs include request IDs, stage timings, response status, first-content timing,
stream snapshot/delta counts, output lengths, and stable error codes. They do
not include prompts, responses, Microsoft Graph tokens, API keys, or enterprise
content. `trace` adds per-snapshot metadata but still does not log message text.
Chat completion responses echo the same correlation value in `X-Request-Id`.

## Test the gateway

```powershell
Invoke-RestMethod -Method Get http://127.0.0.1:8787/health

$body = @{
  model = "m365-copilot-preview"
  messages = @(@{ role = "user"; content = "Reply with exactly: gateway works" })
} | ConvertTo-Json -Depth 5

Invoke-RestMethod -Method Post `
  -Uri http://127.0.0.1:8787/v1/chat/completions `
  -ContentType "application/json" `
  -Body $body
```

For Continue, use `http://127.0.0.1:8787/v1` as the OpenAI-compatible base URL and select `m365-copilot-preview`. If you set `GATEWAY_API_KEY`, enter the same value as the client API key.

## Security boundaries

- The server binds to loopback by default. Do not expose it to a LAN or the Internet without adding a stronger authentication and transport design.
- `.env` and the MSAL cache are ignored by Git. The app registration client ID is not a secret; no client secret is used.
- This project does not log prompts, responses, or Microsoft access tokens.
