# Agentic Harness Architecture for Excel AI Add-in

**Author:** Tarun Kushwaha
**Status:** Proposal — for review
**Audience:** Engineering Manager / Tech Lead
**Date:** 2026-04-28

---

## 1. Executive Summary

We currently run an Excel AI Assistant (Office.js add-in + Node/Express backend + OpenAI GPT-5) that uses a **hand-rolled agent loop**: a single giant text prompt instructs the LLM to emit raw JSON tool calls, which our frontend parses with regex and executes via Office.js.

The loop works, but it is fragile, hard to extend, and pays a structural cost on every conversation:

- Tool calls are returned as **free-text JSON**, parsed with **4 fallback regex strategies** ([src/services/agent.ts:182-233](src/services/agent.ts#L182-L233)).
- The **system prompt is duplicated** in two places ([src/services/agent.ts:9-180](src/services/agent.ts#L9-L180) and [backend/index.js:71-202](backend/index.js#L71-L202)) and drifts.
- Conversation history is concatenated as **plain strings** ([src/services/agent.ts:337](src/services/agent.ts#L337)) with no role separation.
- There is **no schema validation** on tool arguments — a malformed `range` reaches Office.js and throws at runtime.
- There is **no streaming, no observability, and no structured retries**.

This document proposes adopting an **Agentic Harness** — a thin, well-tested orchestration layer that owns the tool-calling loop, schema validation, conversation state, and streaming — so we can stop maintaining bespoke parsing/looping code and focus on Excel tools and UX.

**Recommendation:** Adopt the **Vercel AI SDK** (`ai` package) with `tool()` + `generateText({ tools, stopWhen })` on the backend, and migrate the agent loop from the frontend into the backend in a phased rollout. Estimated effort: **2 weeks (1 engineer)** for the core migration; **+1 week** for streaming UX polish.

---

## 2. Problem Statement

### 2.1 What is an "Agentic Harness"?

An **agentic harness** is the orchestration layer that sits between an LLM and your tools. It owns three responsibilities:

1. **Tool schema enforcement** — declare tools with structured schemas (name, JSON Schema/Zod params, description). The LLM is forced to emit valid arguments; the harness validates before executing.
2. **The agent loop** — call LLM → execute tool → feed result back → call LLM again → repeat until the model declares it is done. The harness owns iteration, max-step caps, and error recovery.
3. **Conversation state** — the harness maintains a properly-roled message history (`system` / `user` / `assistant` / `tool`) instead of one big concatenated string.

You give the harness **tools + a system prompt + a user message**. It returns **the final answer + the trace of what happened**. You no longer write JSON parsers or loop-control code.

### 2.2 Why the Current Hand-Rolled Loop is a Liability

| Pain Point | Where in code | Impact |
|---|---|---|
| Free-text JSON parsing with 4 regex fallbacks | [agent.ts:182-233](src/services/agent.ts#L182-L233) | Silent failures when LLM emits unexpected formatting; brittle to prompt changes |
| System prompt duplicated frontend + backend | [agent.ts:9-180](src/services/agent.ts#L9-L180), [backend/index.js:71-202](backend/index.js#L71-L202) | Already drifted (frontend has more tools than backend) — bug surface |
| No tool-argument schema validation | [agent.ts:336](src/services/agent.ts#L336) | Bad args reach Office.js → user-facing errors |
| Conversation history is `JSON.stringify` concatenation | [agent.ts:337](src/services/agent.ts#L337) | Cannot use OpenAI's native multi-turn caching; wastes tokens |
| Full system prompt re-sent every iteration | [agent.ts:297-316](src/services/agent.ts#L297-L316) | Wastes ~3-4k tokens × N iterations; no prompt caching |
| 15-iteration cap with no early termination signal | [agent.ts:7](src/services/agent.ts#L7) | Hard cap; user sees "max iteration" message instead of clean failure |
| No streaming | [chat.ts:13-40](src/services/chat.ts#L13-L40) | User waits 5-30s with a spinner; no progressive feedback |
| No observability / trace | — | Cannot answer "why did the agent pick this tool?" in production |
| Tool calls JSON spec is **prose in the prompt** | [agent.ts:13-107](src/services/agent.ts#L13-L107) | Source of truth diverges from TypeScript types in [office.ts:21-51](src/services/office.ts#L21-L51) |

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

| Option | Provider lock-in | TS support | Streaming | Schema validation | Maturity | Verdict |
|---|---|---|---|---|---|---|
| **A. OpenAI native function calling** (no SDK) | OpenAI only | ✓ (SDK) | ✓ | Manual (JSON Schema) | Stable | Solid baseline. Forces us to still hand-write the loop. |
| **B. Vercel AI SDK (`ai` package)** | None — provider-agnostic | First-class | ✓ (`streamText`) | ✓ (Zod) | Stable, widely adopted | **Recommended.** Lightweight, owns the loop, easy migration. |
| **C. LangChain.js / LangGraph.js** | None | ✓ but heavy | ✓ | ✓ (Zod) | Mature, complex | Overkill — graph orchestration we don't need. Steep learning curve. |
| **D. OpenAI Agents SDK** (`@openai/agents`) | OpenAI only | ✓ | ✓ | ✓ (Zod) | New (2024+) | Promising but provider-locked; less momentum than Vercel AI SDK in TS ecosystem. |
| **E. Mastra** | None | TS-first | ✓ | ✓ | Newer | Good TS DX but smaller community; risk for production. |

### 4.1 Why Vercel AI SDK Wins for This Project

1. **Already a Node + TS shop** — Vercel AI SDK is TypeScript-first; no impedance mismatch.
2. **`tool()` + Zod** gives us a single source of truth: the Zod schema validates LLM output AND generates the JSON Schema sent to the model. Our [office.ts](src/services/office.ts) `ToolCall` union type maps cleanly to Zod.
3. **`stopWhen: stepCountIs(15)`** replaces our manual loop and `MAX_STEPS` constant ([agent.ts:7](src/services/agent.ts#L7)).
4. **Provider-agnostic** — same code runs against `openai("gpt-5")`, `anthropic("claude-sonnet-4-6")`, or `google("gemini-2.5")`. Future-proof against price/quality shifts.
5. **Streaming built-in** — `streamText` emits `tool-call` and `tool-result` events we can pipe over SSE to the chat UI for live progress.
6. **Used in production** by Vercel, Perplexity, many YC startups — battle-tested.
7. **Small bundle, no graph machinery** — unlike LangGraph, it does not impose a graph DSL. The loop is just a function call.

---

## 5. Proposed Architecture

### 5.1 Where Does the Agent Live?

**Decision: move the agent loop to the backend.**

Today the loop runs on the **frontend** ([agent.ts](src/services/agent.ts)) and the backend is a thin OpenAI proxy. This was fine for the MVP but breaks down because:

- Office.js tool execution **must** stay on the frontend (only the add-in can talk to Excel).
- The LLM-side loop **should** run on the backend — secrets, retries, token usage, observability all belong server-side.

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

A subtlety: in a normal Vercel AI SDK setup, `tool({ execute })` runs the tool **on the server**. But our tools must run on the **client** (Office.js requires the Excel add-in context).

We solve this with a **server-driven tool bridge**:

1. Backend defines tools with Zod schemas but **no `execute` function**.
2. When the model emits a tool call, the SDK pauses the stream waiting for a tool result.
3. Backend forwards the tool call to the client over SSE.
4. Client executes the tool via Office.js ([office.ts executeTool](src/services/office.ts)).
5. Client posts the result back; backend resumes the stream by feeding the result to the model.
6. Loop continues until the model stops calling tools.

This pattern is the standard Vercel AI SDK approach for **client-side tools** — documented and supported.

### 5.3 Module-Level Changes

| File | Today | After |
|---|---|---|
| [src/services/agent.ts](src/services/agent.ts) | 346 lines: prompt + 4 regex parsers + manual loop | **~80 lines:** opens SSE, dispatches tool calls to `executeTool`, posts results back |
| [src/services/office.ts](src/services/office.ts) | Keeps `ToolCall` union + 30+ tool implementations | **Unchanged.** Tool implementations stay; types become source for Zod schemas |
| [src/services/chat.ts](src/services/chat.ts) | One-shot POST | Replaced by SSE consumer (`EventSource` or `fetch` + `ReadableStream`) |
| `backend/index.js` | OpenAI proxy + duplicated system prompt | **Replaced** with `backend/agent.ts`: `streamText` + Zod tool schemas + SSE endpoint |
| **NEW** `backend/tools.ts` | — | Single source of truth: Zod schemas for every tool, mirroring [office.ts:21-51](src/services/office.ts#L21-L51) |
| **NEW** `backend/system-prompt.ts` | — | The system prompt, defined once. Auto-injected by `streamText` |

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

export async function runAgent(userMessage: string, workbookContext: string) {
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
- Add `ai`, `@ai-sdk/openai`, `zod` to `backend/package.json`.
- Create `backend/tools.ts` with Zod schemas for **all 30+ tools** (mechanical port from [office.ts:21-51](src/services/office.ts#L21-L51)).
- Create `backend/system-prompt.ts`: move the prompt out of [backend/index.js:71-202](backend/index.js#L71-L202).
- Create `backend/agent.ts`: `streamText` wrapper.
- New endpoint `POST /agent/stream` (SSE), keeping `/query-answer-final-output` alive for rollback.

### Phase 2 — Client-Side Tool Bridge (Days 4-6)
- Replace [chat.ts](src/services/chat.ts) with an SSE consumer.
- New `src/services/agentClient.ts`:
  - Opens SSE connection.
  - On `tool-call` event → calls `executeTool` from [office.ts](src/services/office.ts) → POSTs result to `/agent/stream/tool-result`.
  - On `text-delta` → appends to chat UI.
  - On `done` → finalizes message.
- Replace [agent.ts](src/services/agent.ts) `runAgent` with the SSE-driven version (~80 lines).

### Phase 3 — Streaming UX (Days 7-9)
- Surface per-step status in the chat ("Sorting range A2:C10…" appears live, not after completion).
- Show tool-call cards with collapsible args/results — better debuggability for users.
- Remove `onStep` callback hack ([agent.ts:286](src/services/agent.ts#L286)) since SSE events drive UI directly.

### Phase 4 — Hardening & Cutover (Days 10-12)
- Add structured logging on the backend (`tool-call`, `tool-result`, `step-finish` events) → enable observability.
- A/B flag: 10% traffic → new `/agent/stream`, 90% → old `/query-answer-final-output`. Watch for regressions.
- After 1 week of clean metrics: flip 100%, delete old endpoint and old [agent.ts](src/services/agent.ts).

### Phase 5 — Optional: Provider Portability (Days 13-14)
- Add `@ai-sdk/anthropic` and validate the same agent against Claude.
- Document model-switch via env var (`MODEL=gpt-5` vs `MODEL=claude-sonnet-4-6`).

---

## 7. Risks & Mitigations

| Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|
| SSE connection drops mid-loop | Medium | User loses progress | Add resume-token; retry logic on client; max 15-step cap unchanged |
| GPT-5 tool-calling quality differs from prose-prompt | Medium | Behavioral regressions | Phase 4 A/B rollout catches this before full cutover |
| Zod schemas reject args the old loop accepted | Low | Some queries fail | Schemas mirror existing TS types; pilot users in Phase 4 |
| Vercel AI SDK breaking changes | Low | Migration churn | Pin major version; SDK is stable at v4+ |
| Bundle size grows on backend | Negligible | — | Vercel AI SDK is small; only on backend, not in add-in |

---

## 8. Effort & Cost Estimate

| Item | Estimate |
|---|---|
| Engineering | 2 weeks (1 engineer) for Phases 1-4; +3 days for Phase 5 |
| New runtime cost | ~Neutral. Token savings from prompt caching offset SSE overhead |
| Risk of regression | **Low** with A/B rollout |
| Reversibility | **High** — old endpoint kept until Phase 4 cutover |

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

- Vercel AI SDK — `tool()` + `generateText` / `streamText`: https://ai-sdk.dev
- Vercel AI SDK — multi-step agents (`stopWhen`): https://ai-sdk.dev/docs/foundations/agents
- Vercel AI SDK — client-side tool execution pattern: https://ai-sdk.dev/cookbook/next/human-in-the-loop
- OpenAI function calling spec: https://platform.openai.com/docs/guides/function-calling
- Current code:
  - [src/services/agent.ts](src/services/agent.ts) — hand-rolled loop to be replaced
  - [src/services/office.ts](src/services/office.ts) — tool implementations (kept)
  - [backend/index.js](backend/index.js) — proxy to be replaced
