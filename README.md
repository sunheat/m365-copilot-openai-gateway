# M365 Copilot OpenAI Gateway

Local TypeScript gateway that exposes a narrow OpenAI Chat Completions-compatible
surface over the Microsoft 365 Copilot Chat API preview.

The gateway is designed for a single signed-in developer on `127.0.0.1`. It
uses delegated Microsoft Entra authentication and never exposes Microsoft Graph
tokens to the IDE client.

## Status

Initial project scaffold. The first implementation will provide Entra device
code authentication, Microsoft Graph Copilot conversations, `/health`,
`/v1/models`, and non-streaming `/v1/chat/completions`.
