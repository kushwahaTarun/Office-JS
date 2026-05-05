// Thin HTTP client for the backend agentic harness (POST /agent/step).
// One request = one LLM call. The agent loop is driven by the caller
// (see agent.ts).

export interface AgentToolCall {
  toolCallId: string;
  toolName: string;
  input: Record<string, unknown>;
}

// Loose ModelMessage shape — we don't import the AI SDK on the frontend
// to keep the bundle lean. The backend validates this on the way in.
export type AgentMessage = Record<string, unknown>;

export interface AgentStepResponse {
  messages: AgentMessage[];
  toolCalls: AgentToolCall[];
  finalText: string | null;
  finishReason?: string;
}

interface StepRequest {
  token: string | null;
  messages: AgentMessage[];
  signal?: AbortSignal;
}

export async function postAgentStep({
  token,
  messages,
  signal,
}: StepRequest): Promise<AgentStepResponse> {
  const url = `${import.meta.env.VITE_API_BASE_URL}/agent/step`;

  const headers: HeadersInit = { "Content-Type": "application/json" };
  if (token) headers["Authorization"] = `Bearer ${token}`;

  const response = await fetch(url, {
    method: "POST",
    body: JSON.stringify({ messages }),
    headers,
    signal,
  });

  if (response.status === 401) {
    throw new Error("Unauthorized. Please login again.", {
      cause: "CAUGHT_ERROR: AUTHENTICATION ERROR",
    });
  }

  if (response.status === 429) {
    throw new Error("Too many requests, please try again after a few seconds.", {
      cause: "CAUGHT_ERROR: NETWORK ERROR",
    });
  }

  if (response.status !== 200) {
    let detail = "";
    try {
      const body = await response.json();
      detail = body?.message ? `: ${body.message}` : "";
    } catch {
      // ignore parse errors
    }
    throw new Error(`Network error, please try again${detail}.`, {
      cause: "CAUGHT_ERROR: NETWORK ERROR",
    });
  }

  const data = (await response.json()) as AgentStepResponse & { status: string };
  return {
    messages: data.messages ?? [],
    toolCalls: data.toolCalls ?? [],
    finalText: data.finalText ?? null,
    finishReason: data.finishReason,
  };
}
