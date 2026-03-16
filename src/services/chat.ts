interface ChatRequest {
  token: string | null;
  tenant: string | null;
  query: string;
  signal?: AbortSignal;
}

interface ChatApiResponse {
  status: string;
  message: string;
}

export async function sendChatMessage({ query, signal }: ChatRequest): Promise<string> {
  const url = `${import.meta.env.VITE_API_BASE_URL}/query-answer-final-output`;

  const response = await fetch(url, {
    method: "POST",
    body: JSON.stringify({ msg: query }),
    headers: new Headers({ "Content-Type": "application/json" }),
    signal,
  });

  if (response.status === 429)
    throw new Error("Too many requests, please try again after a few seconds.", {
      cause: "CAUGHT_ERROR: NETWORK ERROR",
    });
  if (response.status !== 200)
    throw new Error("Network error, please try again.", { cause: "CAUGHT_ERROR: NETWORK ERROR" });

  const data: ChatApiResponse = await response.json();
  return data.message ?? "No response received.";
}
