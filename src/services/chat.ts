const WORKFLOW_NAME = import.meta.env.VITE_WORKFLOW_NAME as string;

interface ChatRequest {
  token: string | null;
  tenant: string | null;
  query: string;
}

interface ChatApiResponse {
  chatId: string;
  queryIds: string[];
  results: Array<{
    query: string;
    answer: string;
    results: unknown;
    split_view: boolean;
  }>;
}

export async function sendChatMessage({ token, tenant, query }: ChatRequest): Promise<string> {
  if (!tenant)
    throw new Error("Missing tenant.", { cause: "CAUGHT_ERROR: MISSING FIELDS" });

  const url = `https://pre-dev.fluidgpt.ai/query-answer-agentic-workflow-final-output`;

  const body = JSON.stringify({
    queries: [{ query }],
    agentic_workflow_name: WORKFLOW_NAME,
  });

  const headers = new Headers({ "Content-Type": "application/json" });
  if (token) headers.append("Authorization", `Bearer ${token}`);

  const response = await fetch(url, {
    method: "POST",
    body,
    headers,
    credentials: token ? undefined : "include",
  });

  if (response.status === 401)
    throw new Error("Missing Authorization.", { cause: "CAUGHT_ERROR: AUTH ERROR" });
  if (response.status === 429)
    throw new Error("Too many requests, please try again after a few seconds.", {
      cause: "CAUGHT_ERROR: NETWORK ERROR",
    });
  if (response.status !== 200)
    throw new Error("Network error, please try again.", { cause: "CAUGHT_ERROR: NETWORK ERROR" });

  const data: ChatApiResponse = await response.json();
  return data.results[0]?.answer ?? "No response received.";
}
