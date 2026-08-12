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
- `POST /v1/chat/completions` for text-only requests, including native SSE streaming when `stream: true`
- Microsoft Entra delegated device-code sign-in with an encrypted, current-user MSAL token cache on Windows
- A new Graph Copilot conversation per request and web grounding disabled on every turn

The streaming path calls Microsoft Graph's native `chatOverStream` endpoint, converts cumulative snapshots into OpenAI-compatible content deltas, and ends successful responses with a `finish_reason` chunk followed by `[DONE]`. `stream_options.include_usage` is accepted but no fabricated token usage is emitted because Graph does not provide token counts. OpenAI tool calling remains unsupported and needs a separate gateway-side protocol.

The proposed native Graph-to-OpenAI streaming architecture is documented in
[Phase 2 SSE Streaming Design](docs/phase-2-sse-streaming-design.md).

## Local setup

1. Copy `.env.example` to `.env` and set `M365_TENANT_ID` and `M365_CLIENT_ID` from your Microsoft Entra app registration. Keep `GATEWAY_HOST=127.0.0.1`.
2. Run `npm install`.
3. Run `npm run auth:login`, open the displayed Microsoft device-login URL, enter the code, and sign in with the M365 Copilot-licensed user.
4. Run `npm run dev`.

The initial interactive sign-in creates the local encrypted cache. Later runs acquire and refresh the delegated access token silently; another login is only needed after sign-out, consent changes, cache deletion, or a revoked refresh token.

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
