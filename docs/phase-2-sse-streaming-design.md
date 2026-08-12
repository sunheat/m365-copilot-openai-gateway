# Phase 2 SSE Streaming Design

Status: Proposed  
Target release: Phase 2  
Last updated: 2026-08-12

## 1. Summary

Phase 2 adds OpenAI-compatible Server-Sent Events (SSE) streaming to
`POST /v1/chat/completions` when the request contains `"stream": true`.

The gateway will use Microsoft Graph's native streamed Copilot endpoint:

```text
POST /beta/copilot/conversations/{conversationId}/chatOverStream
```

It will not wait for a complete synchronous answer and split it into artificial
chunks. Microsoft Graph already returns `text/event-stream`; the gateway will
parse those upstream events, extract progressive answer text, convert cumulative
Copilot snapshots into content deltas, and serialize those deltas as OpenAI
`chat.completion.chunk` events.

The existing non-streaming behavior remains available when `stream` is absent or
false.

## 2. Motivation

Continue's OpenAI-compatible chat path consumes streamed chat completions. The
current phase-1 gateway rejects `stream: true`, so a successful direct HTTP test
does not make the gateway usable as Continue's model backend.

Microsoft documents two Copilot Chat continuation modes: a synchronous `chat`
endpoint and a native SSE `chatOverStream` endpoint. The streamed endpoint
returns `200 OK`, uses `Content-Type: text/event-stream`, and places a complete
`copilotConversation` snapshot in each SSE event. See the
[Microsoft 365 Copilot Chat API overview](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/overview)
and the
[`chatOverStream` API reference](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/ai-services/chat/copilotconversation-chatoverstream).

OpenAI-compatible clients instead expect each SSE `data` field to contain a
`chat.completion.chunk` whose `choices[0].delta` holds only newly generated
content. All chunks for one completion must use a stable ID and timestamp. The
stream ends with a chunk carrying a non-null `finish_reason`, followed by
`data: [DONE]`. See the
[OpenAI Chat Completions API reference](https://developers.openai.com/api/reference/resources/chat).

## 3. Goals

- Accept OpenAI Chat Completions requests with `stream: true`.
- Stream real upstream Copilot progress with minimal buffering.
- Emit a conservative OpenAI-compatible chunk sequence that Continue can read.
- Preserve phase-1 authentication, prompt flattening, web-grounding, error, and
  rate-limit behavior.
- Propagate client cancellation to Microsoft Graph promptly.
- Apply backpressure so a slow client cannot create unbounded memory growth.
- Keep prompts, responses, Microsoft tokens, and enterprise data out of logs and
  test fixtures.
- Retain the synchronous path for `stream: false`.

## 4. Non-goals

- OpenAI tools or function calling.
- `/v1/completions`, Responses API, or Realtime API compatibility.
- Token-accurate usage reporting. The Copilot stream does not expose OpenAI token
  counts.
- Persisting or resuming Copilot conversation IDs between HTTP requests.
- Replaying interrupted streams.
- Exposing Graph attributions, adaptive cards, or sensitivity labels through a
  nonstandard OpenAI extension in Phase 2.
- Supporting production deployment of the Microsoft Graph beta API. Microsoft
  explicitly marks this API as preview and unsupported for production use.

## 5. Architecture decision

### 5.1 Selected approach: Graph SSE to OpenAI SSE

```text
Continue
  | POST /v1/chat/completions { stream: true }
  v
Fastify route
  | authenticate -> create Copilot conversation
  v
Graph chatOverStream
  | SSE containing cumulative copilotConversation snapshots
  v
Graph SSE parser
  | validated conversation objects
  v
Snapshot-to-delta projector
  | assistant text suffixes
  v
OpenAI chunk serializer
  | SSE chat.completion.chunk events + [DONE]
  v
Continue
```

This provides genuine progressive delivery and avoids adding latency solely for
compatibility. It also keeps the synchronous Graph endpoint unchanged for
non-streaming clients.

### 5.2 Rejected approach: synthetic streaming

Calling `/chat`, waiting for the full response, and splitting the final string
would satisfy the wire format but not the behavior users expect from streaming.
It would preserve the full time-to-first-token delay, waste the newly available
Graph streaming capability, and complicate cancellation without benefit.

Synthetic streaming is acceptable only as an explicit emergency fallback if the
preview `chatOverStream` endpoint is temporarily unavailable. It will not be part
of the Phase 2 implementation.

## 6. Request behavior

The existing request schema remains the compatibility boundary, with these
changes:

- Missing `stream` or `stream: false`: keep the synchronous phase-1 path.
- `stream: true`: select the new streamed path.
- `stream_options` is accepted to avoid rejecting clients that send it.
- `stream_options.include_usage` is tolerated, but the gateway does not emit a
  fabricated usage record because Graph provides no token counts.
- `n` is supported only when missing or equal to `1`.
- `tools` remains rejected until the tool protocol is designed.

The requested `model` remains a client-facing compatibility value. Responses use
the stable gateway model ID `m365-copilot-preview`.

For every request, the server creates a new Copilot conversation, flattens the
OpenAI message history into one Graph prompt, and disables web grounding in the
`chatOverStream` request exactly as it does for synchronous chat.

## 7. Upstream Graph client

Extend `CopilotClient` with a streaming method conceptually equivalent to:

```ts
chatStream(
  accessToken: string,
  conversationId: string,
  prompt: string,
  signal: AbortSignal,
): Promise<AsyncIterable<GraphChatStreamEvent>>;
```

The method calls:

```text
POST {graphBaseUrl}/copilot/conversations/{conversationId}/chatOverStream
Authorization: Bearer <delegated-token>
Content-Type: application/json
Accept: text/event-stream
```

The body matches the synchronous path:

```json
{
  "message": { "text": "<flattened prompt>" },
  "locationHint": { "timeZone": "Australia/Sydney" },
  "contextualResources": {
    "webContext": { "isWebEnabled": false }
  }
}
```

Before returning the iterable, the Graph client must verify that:

- the HTTP status is successful;
- the response body exists; and
- the media type is `text/event-stream`, ignoring optional parameters such as
  `charset=utf-8`.

Failures before a stream is established remain normal `GraphCopilotError`
instances and preserve `Retry-After` metadata.

## 8. Graph SSE parsing

Use a small dependency-free parser based on `ReadableStream<Uint8Array>` and a
streaming `TextDecoder`. Do not parse raw network chunks as events: a network
chunk can contain half an event, multiple events, or half a UTF-8 character.

The parser must:

1. Decode UTF-8 incrementally.
2. Recognize LF and CRLF line endings.
3. Finish an event on a blank line.
4. Join multiple `data:` lines with `\n`, following SSE rules.
5. Ignore comment lines and fields other than `data` and optional `id`.
6. Parse each completed `data` value as JSON.
7. Validate only fields used by the gateway and tolerate additional beta fields.
8. Ignore snapshots whose `messages` array is empty.
9. Apply a maximum event size to prevent unbounded buffering if the upstream
   response is malformed.

The Graph SSE `id` is upstream sequencing metadata. It must not be exposed as the
OpenAI completion ID.

## 9. Cumulative snapshot to content delta

Microsoft's examples show that each event contains a `copilotConversation`, not
an OpenAI token delta. Intermediate events may repeat the full answer, contain an
empty `messages` array, or present a longer cumulative version of earlier text.

Maintain this per-request state:

```ts
interface StreamProjectionState {
  completionId: string;
  created: number;
  emittedText: string;
  emittedRole: boolean;
  emittedContent: boolean;
}
```

For each valid snapshot:

1. Select the last message with non-empty textual content.
2. Normalize no content other than removing an optional UTF-8 BOM. Whitespace,
   Markdown, and line endings are model output and must be preserved.
3. If the text equals `emittedText`, emit nothing.
4. If the text starts with `emittedText`, emit only the new suffix and update
   `emittedText`.
5. If no content has been emitted yet, accept the snapshot as the initial text.
6. If a later snapshot diverges from already emitted text, fail the stream. SSE
   cannot retract bytes already consumed by the client, so silently appending a
   replacement would corrupt the answer.

A divergent snapshot is treated as an upstream protocol error. After downstream
headers have been sent, the gateway emits a best-effort OpenAI-shaped error event
and closes without `[DONE]`. Absence of `[DONE]` tells compatible clients that
the completion did not finish successfully.

## 10. OpenAI SSE output

After Graph has returned a successful SSE response, set:

```http
Content-Type: text/event-stream; charset=utf-8
Cache-Control: no-cache, no-transform
Connection: keep-alive
X-Accel-Buffering: no
```

Do not send SSE headers before Graph accepts the request. This preserves the
ability to return a normal JSON `401`, `429`, or `502` for setup failures.

Every downstream event uses exactly this framing:

```text
data: <single-line JSON>\n\n
```

The normal sequence is:

1. Role chunk:

   ```json
   {
     "id": "chatcmpl-<conversation-id>",
     "object": "chat.completion.chunk",
     "created": 1780000000,
     "model": "m365-copilot-preview",
     "choices": [
       { "index": 0, "delta": { "role": "assistant", "content": "" }, "finish_reason": null }
     ]
   }
   ```

2. One or more content chunks using the same `id`, `created`, and `model`:

   ```json
   {
     "id": "chatcmpl-<conversation-id>",
     "object": "chat.completion.chunk",
     "created": 1780000000,
     "model": "m365-copilot-preview",
     "choices": [
       { "index": 0, "delta": { "content": "new text" }, "finish_reason": null }
     ]
   }
   ```

3. Final chunk after a clean upstream EOF and at least one content chunk:

   ```json
   {
     "id": "chatcmpl-<conversation-id>",
     "object": "chat.completion.chunk",
     "created": 1780000000,
     "model": "m365-copilot-preview",
     "choices": [
       { "index": 0, "delta": {}, "finish_reason": "stop" }
     ]
   }
   ```

4. Terminal sentinel:

   ```text
   data: [DONE]

   ```

Do not expose Graph conversation snapshots, SSE IDs, access tokens,
attributions, adaptive cards, or sensitivity labels in the OpenAI chunks.

## 11. Completion and error semantics

### 11.1 Before downstream SSE begins

Return the existing JSON error envelope and HTTP status mapping:

| Condition | HTTP status | Error code |
| --- | ---: | --- |
| Gateway API key failure | 401 | `invalid_api_key` |
| Entra login required | 401 | `m365_login_required` |
| Invalid OpenAI request | 400 | `invalid_request` |
| Unsupported tools or options | 400 | Existing compatibility code |
| Graph throttling | 429 | `rate_limit_exceeded` plus `Retry-After` |
| Other Graph failure | 502 | `graph_<status>` |

### 11.2 After downstream SSE begins

HTTP status and headers can no longer be changed. On malformed Graph events,
snapshot divergence, idle timeout, or upstream transport failure:

- emit one best-effort SSE event whose JSON has a top-level `error` object;
- never include raw Graph response bodies or tenant data in the error;
- do not emit a final `finish_reason: "stop"` chunk;
- do not emit `[DONE]`; and
- close the connection.

The error event is a pragmatic compatibility measure rather than a standard
Chat Completion chunk. The missing `[DONE]` is the authoritative indication of
an incomplete response.

An upstream EOF is successful only if at least one non-empty assistant content
delta was emitted. EOF with no answer is an upstream protocol error.

## 12. Cancellation, timeouts, and backpressure

Create one `AbortController` per streamed request. Abort the Graph fetch when:

- the client aborts the request;
- the downstream socket closes before normal completion;
- the stream-start timeout expires;
- the upstream idle timeout expires; or
- the server shuts down.

Remove event listeners and timers in a `finally` block. Client cancellation is
expected and should not be logged as an application error.

Suggested configurable defaults:

| Setting | Default | Meaning |
| --- | ---: | --- |
| `GRAPH_STREAM_START_TIMEOUT_MS` | 30,000 | Time to receive successful upstream headers |
| `GRAPH_STREAM_IDLE_TIMEOUT_MS` | 60,000 | Maximum time between upstream bytes/events |
| `GRAPH_STREAM_MAX_EVENT_BYTES` | 2 MiB | Maximum buffered Graph SSE event |
| `GATEWAY_SSE_HEARTBEAT_MS` | 15,000 | Downstream SSE comment interval while Graph is silent |

Heartbeat comments use `: keep-alive\n\n` and do not affect OpenAI content.

Honor Node response backpressure: when `reply.raw.write()` returns `false`, wait
for `drain` before reading and serializing more events. Never accumulate the
entire Graph response in memory.

## 13. Proposed module changes

```text
src/
  graph-copilot.ts        Add chatStream() and upstream response validation
  graph-sse-parser.ts     Convert byte stream into validated Graph events
  openai-stream.ts        Snapshot projection and chunk serialization
  server.ts               Route stream:true requests and manage lifecycle
  types.ts                Graph stream and OpenAI chunk types
test/
  graph-sse-parser.test.ts
  openai-stream.test.ts
  server-streaming.test.ts
```

No new runtime dependency is required. Native `fetch`, Web Streams,
`TextDecoder`, and Node HTTP primitives are sufficient on the project's minimum
Node version.

## 14. Test strategy

### 14.1 Parser unit tests

- One SSE event split across multiple byte chunks.
- Multiple SSE events in one byte chunk.
- UTF-8 multibyte characters split across chunks.
- LF and CRLF endings.
- Multiple `data:` lines and ignored comment/unknown fields.
- Empty `messages` snapshots.
- Malformed JSON, missing message text, oversized events, and unexpected EOF.

Use synthetic fixtures. Do not check real enterprise answers or identifiers into
the repository.

### 14.2 Projection unit tests

- Initial role chunk is emitted once.
- Cumulative `"Hello"` then `"Hello world"` produces `"Hello"` then
  `" world"`.
- Duplicate snapshots produce no duplicate output.
- Empty intermediate snapshots produce no output.
- Whitespace and Markdown are preserved exactly.
- Divergent snapshots fail once content has been emitted.
- All chunks share the same completion ID, model, and timestamp.
- Clean completion emits `finish_reason: "stop"` and `[DONE]` exactly once.

### 14.3 Route integration tests

- `stream: false` remains byte-for-byte compatible with Phase 1.
- `stream: true` uses `chatStream`, not synchronous `chat`.
- Headers are correct and chunks arrive before upstream completion.
- Authentication and Graph failures before headers remain JSON responses.
- Graph `429` preserves `Retry-After`.
- Failure after streaming begins closes without `[DONE]`.
- Client disconnect aborts the mocked Graph stream.
- A slow downstream consumer exercises the backpressure path.

Use a real loopback listener for at least one test. Request injection libraries
often buffer responses and cannot prove incremental delivery.

### 14.4 Manual acceptance tests

1. Run `npm run auth:login` if required and start the gateway.
2. Use `curl.exe -N` or a small Node client to confirm incremental SSE chunks.
3. Configure Continue with:

   ```yaml
   models:
     - name: Microsoft 365 Copilot Gateway
       provider: openai
       model: m365-copilot-preview
       apiBase: http://127.0.0.1:8787/v1
       apiKey: <GATEWAY_API_KEY or a non-empty placeholder>
       roles:
         - chat
         - edit
   ```

4. Verify normal chat, code explanation, and a small edit request.
5. Cancel a long response from Continue and confirm the Graph request stops.
6. Confirm logs contain no prompt, answer, Graph token, or enterprise content.

## 15. Acceptance criteria

Phase 2 is complete when:

- Continue can receive and render a response through the gateway with
  `stream: true`.
- The first content chunk arrives before the Graph stream completes.
- The output is valid OpenAI Chat Completions SSE with stable metadata, a final
  `finish_reason`, and exactly one `[DONE]` on success.
- Duplicate cumulative Graph snapshots do not duplicate visible text.
- Authentication, throttling, cancellation, malformed events, timeouts, and
  client disconnects have automated coverage.
- The existing non-streaming API and tests continue to pass.
- Type checking, unit tests, production build, and a live M365 Copilot streaming
  smoke test pass.

## 16. Implementation order

1. Add Graph SSE parser and synthetic fixtures.
2. Add cumulative snapshot projector and OpenAI chunk serializer.
3. Add `CopilotClient.chatStream()` with abort and timeout support.
4. Add the Fastify streaming route with headers, backpressure, and lifecycle
   cleanup.
5. Add loopback integration tests.
6. Run a live Graph `chatOverStream` smoke test to validate assumptions about
   snapshot progression and EOF behavior.
7. Test Continue chat and edit roles end to end.

The live smoke test is intentionally late: parser and projection behavior should
first be deterministic and fully testable without storing enterprise data.

## 17. Risks and follow-up decisions

- **Graph beta drift:** event shape or termination behavior can change. Keep
  validation tolerant of unknown fields and isolate Graph parsing from OpenAI
  serialization.
- **Snapshot revisions:** SSE cannot retract already emitted text. The selected
  fail-closed behavior prevents silent answer corruption but may surface an
  interrupted response if Graph revises a prefix.
- **Attributions:** dropping attribution metadata preserves a narrow OpenAI
  surface but loses citations. A later phase can map supported citations to
  annotations or append formatted references after a separate design review.
- **Usage accounting:** Graph does not provide token counts. Do not estimate and
  present them as authoritative usage.
- **Tool calling:** streaming text compatibility does not make the model agentic.
  Continue tool use requires a distinct, validated tool-call protocol and remains
  Phase 3 work.
