# Phase 3 OpenAI Tool-Calling Compatibility Design

## 1. Objective

Phase 3 adds a narrow OpenAI Chat Completions function-calling surface so IDE
clients can perform a human-directed tool loop while Microsoft 365 Copilot
remains the text-only upstream model.

Microsoft Graph Copilot Chat does not accept tool definitions and does not emit
native tool calls. The gateway therefore presents the available functions to
Copilot as a bounded text protocol, validates the complete response locally,
and only then exposes an OpenAI-compatible `tool_calls` response to the client.

This is a compatibility experiment, not a claim that Microsoft 365 Copilot has
native tool-calling support.

## 2. Supported OpenAI surface

The first release supports Chat Completions function tools:

- `tools[].type: "function"`;
- function `name`, optional `description`, and JSON Schema `parameters`;
- `tool_choice: "none" | "auto" | "required"`;
- a named function choice;
- assistant messages containing one prior `tool_calls` entry;
- tool-result messages containing `tool_call_id` and text content;
- non-streaming `message.tool_calls` responses; and
- streaming `delta.tool_calls` responses ending with
  `finish_reason: "tool_calls"` and `[DONE]`.

The gateway accepts `parallel_tool_calls`, but this phase emits at most one tool
call per assistant response. Custom tools, built-in OpenAI tools, multiple
choices, and more than one tool call in an assistant history message are out of
scope.

## 3. Request validation

Only function tools are accepted. Tool names must be unique and match
`^[A-Za-z0-9_-]{1,64}$`. Each `parameters` value must be a valid JSON Schema.

`tool_choice` behavior:

- `none`: use the normal text-only path and do not present tools upstream;
- omitted or `auto`: Copilot may return final text or one tool call;
- `required`: Copilot must return one of the supplied tools; and
- named function: Copilot must return that function.

Malformed definitions, duplicate names, unknown named choices, and invalid
schemas are rejected with HTTP 400 before Microsoft Graph is called.

## 4. Upstream protocol

For a tool-enabled turn the gateway generates a random nonce and appends a
control record to the JSON-framed conversation transcript. Copilot must return
exactly one JSON object and no Markdown fence or explanatory text.

Final answer envelope:

```json
{
  "type": "final",
  "nonce": "turn nonce",
  "content": "answer text"
}
```

Tool-call envelope:

```json
{
  "type": "tool_call",
  "nonce": "turn nonce",
  "name": "read_file",
  "arguments": {
    "path": "src/server.ts"
  }
}
```

The random nonce prevents source text containing a plausible static envelope
from being mistaken for the gateway control response. It is not a secret and
does not make prompt injection impossible; local structural validation remains
the authority.

## 5. Local validation

The gateway accepts an envelope only when all of these checks pass:

1. The entire trimmed response is one JSON object.
2. It contains only the fields allowed for its declared type.
3. The nonce exactly matches the current turn.
4. The requested tool is allowed by `tool_choice` and exists in the request.
5. `arguments` is an object that validates against that function's JSON Schema.

Unknown tools, schema failures, Markdown fences, extra prose, extra fields,
multiple calls, or mixed final text and tool-call content fail closed. No local
tool is executed by the gateway.

## 6. One correction retry

If Copilot's first response is not a valid envelope, the gateway sends one
fixed correction message in the same Graph conversation. The correction names
the validation category without copying the invalid response and repeats the
required nonce and response shapes.

If the second response is still invalid, the gateway returns HTTP 502 with the
stable code `tool_protocol_error`. There is no unbounded retry loop.

## 7. Streaming behavior

Tool-enabled requests are fully buffered upstream, including requests where
the client sets `stream: true`. This is required because the gateway cannot
know whether partial text is a final answer or an incomplete JSON tool envelope.

After validation:

- a final answer is emitted as one content chunk, a `stop` chunk, then `[DONE]`;
- a tool call is emitted as one `delta.tool_calls` chunk, a `tool_calls` finish
  chunk, then `[DONE]`; and
- protocol failure occurs before downstream SSE headers, so the client receives
  a normal JSON HTTP error.

Requests without active tools retain the native low-latency Graph SSE path.

## 8. Message-history projection

The JSON transcript preserves:

- system, user, assistant, and tool roles;
- assistant final text;
- the prior assistant tool-call ID, function name, and argument string; and
- the matching tool-result `tool_call_id` and content.

This lets an IDE execute a tool locally and submit the result in the next Chat
Completions request without reusing the stateful upstream Graph conversation.
The gateway remains stateless across HTTP requests and creates a new Graph
conversation for every request.

## 9. Security and diagnostics

- Tool definitions and tool results are untrusted transcript data.
- The gateway validates but never executes tool calls.
- Logs record only request IDs, counts, outcome kind, timings, lengths, and
  stable error codes. Tool arguments, prompts, results, and response text are
  never logged.
- Web grounding remains disabled on the initial and correction turns.
- Invalid envelopes never become assistant text or executable tool calls.

## 10. Acceptance criteria

- Existing text-only sync and SSE behavior remains unchanged.
- Valid auto, required, and named tool choices map to OpenAI-compatible output.
- `tool_choice: "none"` retains the native text path.
- Arguments are validated against the submitted JSON Schema.
- Invalid first output can recover once; invalid second output fails closed.
- Streaming tool calls contain stable IDs and valid argument JSON.
- Prior assistant tool calls and tool results survive transcript framing.
- Full type-check, unit test, build, and diff checks pass.
- A live, harmless tool-decision smoke test succeeds when local Microsoft
  credentials are available.

## 11. Follow-up decisions

Before adding parallel calls or more permissive extraction, measure at least 50
tool-decision cases. Stop treating the Chat API as an agent backend if valid
single-call structure remains below 80% after the one correction retry or if
prompt injection can produce a locally accepted unintended call.

Potential later work includes parallel calls, incremental argument streaming,
prefix-hash Graph conversation reuse, and a Responses API adapter. None of
those are required for this phase.
