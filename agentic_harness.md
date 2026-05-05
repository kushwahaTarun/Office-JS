# Agentic Harness Architecture for Excel AI Add-in

## 1. Executive Summary

The Excel AI Assistant (Office.js add-in + Node/Express backend + OpenAI GPT-5) currently uses a **hand-rolled agent loop**: a giant text prompt instructs the LLM to emit raw JSON tool calls, which the frontend parses with regex and executes via Office.js. It works, but is fragile and pays a structural cost on every conversation:

- Tool calls returned as **free-text JSON**, parsed with **4 fallback regex strategies** ([src/services/agent.ts:182-233](src/services/agent.ts#L182-L233)).
- **System prompt duplicated** in two places ([agent.ts:9-180](src/services/agent.ts#L9-L180), [backend/index.js:71-202](backend/index.js#L71-L202)) — already drifted.
- Conversation history is **string concatenation** ([agent.ts:337](src/services/agent.ts#L337)), no role separation.
- **No schema validation** on tool arguments; bad args reach Office.js and throw.
- No streaming, no observability, no structured retries.

**Proposal:** adopt an **Agentic Harness** — a thin, battle-tested orchestration layer that owns the tool-calling loop, schema validation, conversation state, and streaming.

**Recommendation:** Vercel AI SDK (`ai` v6 — Dec 2025) with `tool()` + `streamText({ tools, stopWhen })` on the backend. Migrate the agent loop from frontend to backend in a phased rollout.

**Packages to install (backend):**

```bash
npm install ai @ai-sdk/openai zod
```

**Effort:** 2 weeks (1 engineer) for core migration; +1 week for streaming UX polish.

---

## 2. Problem Statement

### 2.1 What is an "Agentic Harness"?

An **agentic harness** is the orchestration layer between an LLM and tools. It owns three responsibilities:

1. **Tool schema enforcement** — declare tools with structured schemas (Zod / JSON Schema). The LLM is forced to emit valid arguments; the harness validates before executing.
2. **The agent loop** — call LLM → execute tool → feed result back → repeat until the model is done. The harness owns iteration, step caps, error recovery.
3. **Conversation state** — properly-roled message history (`system` / `user` / `assistant` / `tool`) instead of concatenated strings.

You give it **tools + a system prompt + a user message**. It returns **the final answer + the trace**. No JSON parsers, no loop-control code.

### 2.2 Why the Current Hand-Rolled Loop is a Liability

| Pain Point                                             | Where in code                                                                                         | Impact                                                                                              |
| ------------------------------------------------------ | ----------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| Free-text JSON parsing with 4 regex fallbacks          | [agent.ts:182-233](src/services/agent.ts#L182-L233)                                                   | Silent failures when LLM emits unexpected formatting; brittle to prompt changes                     |
| System prompt duplicated frontend + backend            | [agent.ts:9-180](src/services/agent.ts#L9-L180), [backend/index.js:71-202](backend/index.js#L71-L202) | Already drifted (frontend has more tools than backend) — bug surface                                |
| No tool-argument schema validation                     | [agent.ts:336](src/services/agent.ts#L336)                                                            | Bad args reach Office.js → user-facing errors                                                       |
| Conversation history is `JSON.stringify` concatenation | [agent.ts:337](src/services/agent.ts#L337)                                                            | Cannot use OpenAI's native multi-turn caching; wastes tokens                                        |
| Full system prompt re-sent every iteration             | [agent.ts:297-316](src/services/agent.ts#L297-L316)                                                   | Wastes ~3-4k tokens × N iterations; no prompt caching                                               |
| 15-iteration cap with no early termination signal      | [agent.ts:7](src/services/agent.ts#L7)                                                                | Hard cap; user sees "max iteration" message instead of clean failure                                |
| No streaming                                           | [chat.ts:13-40](src/services/chat.ts#L13-L40)                                                         | User waits 5-30s with a spinner; no progressive feedback                                            |
| No observability / trace                               | —                                                                                                     | Cannot answer "why did the agent pick this tool?" in production                                     |
| Tool calls JSON spec is **prose in the prompt**        | [agent.ts:13-107](src/services/agent.ts#L13-L107)                                                     | Source of truth diverges from TypeScript types in [office.ts:21-51](src/services/office.ts#L21-L51) |

### 2.3 Symptoms the Team Has Already Observed

- LLM occasionally wraps JSON in markdown fences → parse miss → silent failure (worked around with regex #1 in [agent.ts:191](src/services/agent.ts#L191)).
- LLM emits `Math.random()` literally inside JSON → had to add a sanitizer ([agent.ts:184](src/services/agent.ts#L184)).
- LLM picks unsupported chart type `"Bar"` → Office.js throws → caught only as a generic error string.

These are exactly the failure modes a harness eliminates by design.

---

## 3. Goals & Non-Goals

### 3.1 Goals

- **Reliability:** Eliminate JSON-parsing failures by using native structured tool calls.
- **Single source of truth:** Tools are defined once (with Zod schemas) and both validate args and document themselves to the LLM.
- **Observability:** Every step (LLM call, tool call, tool result) is captured as a structured event, not a string.
- **Token efficiency:** Use prompt caching and proper roled messages so the system prompt is sent once per conversation, not 15 times.
- **Streaming UX:** Stream the agent's progress to the chat (tool-by-tool), not just the final answer.
- **Provider portability:** Swap GPT-5 for Claude / Gemini / Azure OpenAI without rewriting the loop.

### 3.2 Non-Goals (out of scope for this phase)

- Multi-agent orchestration (we have one agent — Excel).
- Long-term memory / vector store.
- Replacing Auth0 or the Office.js layer.
- Adding new Excel tools (the existing 30+ tools will be ported as-is).

---

## 4. Options Evaluated

| Option                                         | Provider lock-in         | TS support  | Streaming        | Schema validation  | Maturity               | Verdict                                                                       |
| ---------------------------------------------- | ------------------------ | ----------- | ---------------- | ------------------ | ---------------------- | ----------------------------------------------------------------------------- |
| **A. OpenAI native function calling** (no SDK) | OpenAI only              | ✓ (SDK)     | ✓                | Manual JSON Schema | Stable                 | Solid baseline. Forces us to still hand-write the loop.                       |
| **B. Vercel AI SDK** (`ai` v6)                 | None — provider-agnostic | First-class | ✓ (`streamText`) | ✓ (Zod)            | Stable, widely adopted | **Recommended.** Lightweight, owns the loop, easy migration.                  |
| **C. LangChain.js / LangGraph.js**             | None                     | ✓ but heavy | ✓                | ✓ (Zod)            | Mature, complex        | Overkill — graph orchestration we don't need. Steep learning curve.           |
| **D. OpenAI Agents SDK** (`@openai/agents`)    | OpenAI only              | ✓           | ✓                | ✓ (Zod)            | New (TS port of Swarm) | Promising but provider-locked; no `useChat`-equivalent for client-side tools. |
| **E. Mastra**                                  | None                     | TS-first    | ✓                | ✓                  | Newer                  | Good TS DX but smaller community; risk for production.                        |

### 4.1 Why Vercel AI SDK Wins for This Project

1. **TypeScript-first** — no impedance mismatch with our Node + TS stack.
2. **`tool()` + Zod = single source of truth** — `inputSchema` validates LLM output AND generates the JSON Schema sent to the model. Our [office.ts](src/services/office.ts) `ToolCall` union maps cleanly to Zod.
3. **`stopWhen: stepCountIs(15)`** replaces our manual loop and `MAX_STEPS` ([agent.ts:7](src/services/agent.ts#L7)).
4. **Provider-agnostic** — same code runs against `openai("gpt-5")`, `anthropic("claude-sonnet-4-6")`, or Gemini. Future-proof against price/quality shifts.
5. **Streaming built-in** — `streamText` emits `tool-call` / `tool-result` events we pipe over SSE for live progress.
6. **First-class client-side tool execution** — define a tool without `execute`, client runs it via Office.js, returns result via `addToolOutput` ([official docs](https://ai-sdk.dev/docs/ai-sdk-ui/chatbot-tool-usage)). Exact fit for our architecture.
7. **`needsApproval: true`** (SDK 6) — built-in human-in-the-loop confirmation for destructive ops like `delete_rows`, `clear_range`.
8. **Battle-tested in production** by Vercel, Perplexity, many YC startups.
9. **No graph machinery** — unlike LangGraph, the loop is just a function call.

---

## 5. Proposed Architecture

### 5.1 Where Does the Agent Live?

**Decision: move the agent loop to the backend.** Today the loop runs on the frontend ([agent.ts](src/services/agent.ts)) and the backend is a thin OpenAI proxy. This was fine for the MVP but:

- Office.js tool execution **must** stay on the frontend (only the add-in can talk to Excel).
- The LLM-side loop **should** run on the backend — secrets, retries, token usage, and observability all belong server-side.

The new design splits responsibilities cleanly:

```
┌───────────────────────────────────────────────────────────────────────┐
│  FRONTEND (Office.js add-in — React + TS)                             │
│                                                                       │
│  • Reads workbook context via Office.js                               │
│  • Streams chat over SSE                                              │
│  • Executes tool calls the backend asks for, posts result back        │
│  • Renders progress (tool-by-tool) and final answer                   │
└──────────────────────────────┬────────────────────────────────────────┘
                               │  SSE / WebSocket
                               │  (workbook context + user message)
                               │  (tool results posted back per step)
                               ▼
┌───────────────────────────────────────────────────────────────────────┐
│  BACKEND (Node + Express + Vercel AI SDK)                             │
│                                                                       │
│  • Auth0 JWT middleware (existing)                                    │
│  • streamText({ model, tools, system, messages, stopWhen })           │
│  • Tool schemas defined ONCE with Zod (mirrors office.ts types)       │
│  • For each tool call:                                                │
│       - validate args with Zod                                        │
│       - send "execute_tool" event to client                           │
│       - await client's tool result                                    │
│       - feed result back to the model                                 │
│  • Emits structured events: text-delta, tool-call, tool-result, done  │
└──────────────────────────────┬────────────────────────────────────────┘
                               │
                               ▼
                          OpenAI GPT-5
                  (or any model — provider-agnostic)
```

### 5.2 The Tool Execution Bridge (Key Design Decision)

In a normal Vercel AI SDK setup, `tool({ execute })` runs the tool **on the server**. But our tools must run on the **client** (Office.js requires the Excel add-in context). We solve this with a **server-driven tool bridge**:

1. Backend defines tools with Zod schemas but **no `execute`**.
2. Model emits a tool call → SDK pauses the stream awaiting a result.
3. Backend forwards the call to the client over SSE.
4. Client executes via Office.js ([office.ts executeTool](src/services/office.ts)).
5. Client posts the result back via `addToolOutput` → backend resumes the stream.
6. Loop continues until the model stops calling tools.

This is the standard Vercel AI SDK pattern for client-side tools ([docs](https://ai-sdk.dev/docs/ai-sdk-ui/chatbot-tool-usage)).

### 5.3 Module-Level Changes

| File                                             | Today                                             | After                                                                                                           |
| ------------------------------------------------ | ------------------------------------------------- | --------------------------------------------------------------------------------------------------------------- |
| [src/services/agent.ts](src/services/agent.ts)   | 346 lines: prompt + 4 regex parsers + manual loop | **~80 lines:** opens SSE, dispatches tool calls to `executeTool`, posts results back                            |
| [src/services/office.ts](src/services/office.ts) | Keeps `ToolCall` union + 30+ tool implementations | **Unchanged.** Tool implementations stay; types become source for Zod schemas                                   |
| [src/services/chat.ts](src/services/chat.ts)     | One-shot POST                                     | Replaced by SSE consumer (`EventSource` or `fetch` + `ReadableStream`)                                          |
| `backend/index.js`                               | OpenAI proxy + duplicated system prompt           | **Replaced** with `backend/agent.ts`: `streamText` + Zod tool schemas + SSE endpoint                            |
| **NEW** `backend/tools.ts`                       | —                                                 | Single source of truth: Zod schemas for every tool, mirroring [office.ts:21-51](src/services/office.ts#L21-L51) |
| **NEW** `backend/system-prompt.ts`               | —                                                 | The system prompt, defined once. Auto-injected by `streamText`                                                  |

### 5.4 Tool Definition — Before vs After

**Today** (system prompt instructs the LLM in prose):

```
{"tool":"sort_range","sheet":"Sheet1","range":"A2:C10","sortColumn":0,"ascending":true}
  → Sort by column index (0-based)
```

→ then in [agent.ts:202](src/services/agent.ts#L202): `text.match(/\{[^{}]*"tool"\s*:[^{}]*\}/s)` and hope for the best.

**After** (in `backend/tools.ts`):

```ts
import { tool } from "ai";
import { z } from "zod";

export const sortRange = tool({
  description: "Sort a range by a given column. sortColumn is 0-based.",
  inputSchema: z.object({
    sheet: z.string(),
    range: z.string().regex(/^[A-Z]+\d+:[A-Z]+\d+$/, "Must be A1:B2 notation"),
    sortColumn: z.number().int().min(0),
    ascending: z.boolean().default(true),
  }),
  // No `execute` — handled on the client via the bridge
});
```

Benefits:

- The Zod schema **becomes** the JSON Schema sent to the LLM (no prose drift).
- Bad args (e.g. `"range": "Sheet1!A1"`) are rejected before reaching Office.js.
- TypeScript infers the arg shape end-to-end.

### 5.5 The Loop — Before vs After

**Today** ([agent.ts:294-345](src/services/agent.ts#L294-L345)):

```ts
for (let llmCall = 0; llmCall < MAX_STEPS; llmCall++) {
  const response = await sendChatMessage(...);
  const toolCalls = parseToolCalls(response);   // 4 regex fallbacks
  if (!toolCalls.length) return response;
  for (const toolCall of toolCalls) {
    if (toolCall.tool === "final_answer") return ...;
    const result = await executeTool(toolCall);
    history.push(`Step ${n}: ${JSON.stringify(toolCall)} → ${JSON.stringify(result)}`);
  }
}
```

**After** (in `backend/agent.ts`):

```ts
import { streamText, stepCountIs } from "ai";
import { openai } from "@ai-sdk/openai";
import { tools } from "./tools";
import { systemPrompt } from "./system-prompt";

export function runAgent(userMessage: string, workbookContext: string) {
  return streamText({
    model: openai("gpt-5"),
    system: systemPrompt,
    messages: [
      { role: "user", content: `${workbookContext}\n\n${userMessage}` },
    ],
    tools,
    stopWhen: stepCountIs(15),
  });
}
```

That's the entire loop. No regex. No history concatenation. No `final_answer` sentinel — the model stops calling tools naturally and emits final text, which the SDK surfaces as the assistant message.

---

## 6. Migration Plan (Phased)

### Phase 1 — Backend Harness Skeleton (Days 1-3)

- Install `ai`, `@ai-sdk/openai`, `zod`.
- `backend/tools.ts`: Zod schemas for all 30+ tools (mechanical port from [office.ts:21-51](src/services/office.ts#L21-L51)).
- `backend/system-prompt.ts`: extract prompt from [backend/index.js:71-202](backend/index.js#L71-L202).
- `backend/agent.ts`: `streamText` wrapper.
- New endpoint `POST /agent/stream` (SSE); keep `/query-answer-final-output` for rollback.

### Phase 2 — Client-Side Tool Bridge (Days 4-6)

- Replace [chat.ts](src/services/chat.ts) with an SSE consumer.
- New `src/services/agentClient.ts`: opens SSE; on `tool-call` → runs `executeTool`, posts result via `addToolOutput`; on `text-delta` → appends to chat; on `done` → finalizes.
- Shrink [agent.ts](src/services/agent.ts) `runAgent` to ~80 lines (SSE-driven).

### Phase 3 — Streaming UX (Days 7-9)

- Live per-step status in chat ("Sorting A2:C10…" as it happens).
- Tool-call cards with collapsible args/results.
- Remove `onStep` callback hack ([agent.ts:286](src/services/agent.ts#L286)) — SSE drives the UI.

### Phase 4 — Hardening & Cutover (Days 10-12)

- Structured logging for `tool-call` / `tool-result` / `step-finish` events.
- A/B flag: 10% → new endpoint, 90% → old. After 1 week clean: flip 100%, delete old code.

### Phase 5 — Optional: Provider Portability (Days 13-14)

- Add `@ai-sdk/anthropic`, validate against Claude. Model switch via env var.

---

## 7. Risks & Mitigations

| Risk                                                 | Likelihood | Impact                 | Mitigation                                                                      |
| ---------------------------------------------------- | ---------- | ---------------------- | ------------------------------------------------------------------------------- |
| SSE connection drops mid-loop                        | Medium     | User loses progress    | Add resume-token; retry logic on client; max 15-step cap unchanged              |
| GPT-5 tool-calling quality differs from prose-prompt | Medium     | Behavioral regressions | Phase 4 A/B rollout catches this before full cutover                            |
| Zod schemas reject args the old loop accepted        | Low        | Some queries fail      | Schemas mirror existing TS types; pilot users in Phase 4                        |
| Vercel AI SDK breaking changes                       | Low        | Migration churn        | Pin major version (v6.x); codemod available for v6 → v7 (`npx @ai-sdk/codemod`) |
| Bundle size grows on backend                         | Negligible | —                      | Vercel AI SDK is small; only on backend, not in add-in                          |

---

## 8. Effort & Cost Estimate

| Item               | Estimate                                                        |
| ------------------ | --------------------------------------------------------------- |
| Engineering        | 2 weeks (1 engineer) for Phases 1-4; +3 days for Phase 5        |
| New runtime cost   | ~Neutral. Token savings from prompt caching offset SSE overhead |
| Risk of regression | **Low** with A/B rollout                                        |
| Reversibility      | **High** — old endpoint kept until Phase 4 cutover              |

---

## 9. Success Criteria

After cutover, we should observe:

1. **Zero JSON parse failures** in logs (today: occasional regex misses).
2. **Single source of truth** for tool definitions (today: 3 places — types, frontend prompt, backend prompt).
3. **Per-step UI feedback** in <500ms (today: full wait until final answer).
4. **Tool argument validation errors caught server-side** before reaching Office.js.
5. **Token usage per conversation drops ≥30%** from prompt caching + roled messages.
6. **Agent traces** queryable in logs (model decisions, tool args, results) for every conversation.

---

## 10. Open Questions for Review

1. **Backend hosting** — current backend is local Express. Is there a target deploy environment (Vercel, AWS, Render)? Affects SSE timeout config.
2. **Auth pass-through on SSE** — JWT on initial connection is straightforward; do we need re-auth on long sessions?
3. **Telemetry vendor** — should we add OpenTelemetry / Datadog / Vercel Observability hooks in Phase 4?
4. **Model strategy** — stick with GPT-5, or use Phase 5 to evaluate Claude / Gemini for cost/quality?

---

## 11. References

- Vercel AI SDK 6 announcement: https://vercel.com/blog/ai-sdk-6
- Vercel AI SDK — agents / loop control (`stopWhen`, `stepCountIs`): https://ai-sdk.dev/docs/agents/loop-control
- Vercel AI SDK — client-side tool execution (`addToolOutput`): https://ai-sdk.dev/docs/ai-sdk-ui/chatbot-tool-usage
- OpenAI function calling spec: https://platform.openai.com/docs/guides/function-calling
- OpenAI Agents SDK (TypeScript): https://openai.github.io/openai-agents-js/
- Current code:
  - [src/services/agent.ts](src/services/agent.ts) — hand-rolled loop to be replaced
  - [src/services/office.ts](src/services/office.ts) — tool implementations (kept)
  - [backend/index.js](backend/index.js) — proxy to be replaced
