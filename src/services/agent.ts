// Client-side agent loop. The Vercel AI SDK drives the LLM ↔ tool
// round-trips entirely in the browser — the backend `/agent/step` route
// is no longer involved.
import { generateText, stepCountIs } from "ai";
import { createOpenAI } from "@ai-sdk/openai";
import { getWorkbookContext } from "./office";
import type { WorkbookContext } from "./office";
import { getWorkbookInsights } from "./contextUtils";
import { parseAIResponse } from "./markdownUtils";
import { SYSTEM_PROMPT } from "../agent/systemPrompt";
import { createTools } from "../agent/tools";

const MAX_STEPS = 20;
const MODEL_ID = "gpt-5";

// SECURITY NOTE: this API key ships in the client bundle. Treat it as
// public — the org-wide proxy token must be rotated to per-user tokens
// before any wider release.
const openai = createOpenAI({
  baseURL: import.meta.env.VITE_OPENAI_BASE_URL,
  apiKey: import.meta.env.VITE_OPENAI_API_KEY,
});

function formatContext(ctx: WorkbookContext): string {
  const selected = ctx.selectedRange;
  const activeSheet = ctx.sheetsMetadata.find((m) => m.name === ctx.activeSheet);
  const sheetsMeta = ctx.sheetsMetadata
    .map(
      (m) =>
        `- ${m.name}: ${m.rowCount} rows, ${m.columnCount} cols ${m.hasData ? "(contains data)" : "(empty)"}`,
    )
    .join("\n");

  const insights = getWorkbookInsights(ctx.sheetsMetadata);
  const insightsStr =
    insights.recommendations.length > 0
      ? `\nWorkbook Insights:\n${insights.recommendations.map((r) => `- ${r}`).join("\n")}`
      : "";

  const lastRowNote =
    activeSheet && activeSheet.rowCount > 0
      ? `IMPORTANT: Active sheet has exactly ${activeSheet.rowCount} rows of used data. Last data row = row ${activeSheet.rowCount}. Next empty row = row ${activeSheet.rowCount + 1}. Do NOT use the preview row count (capped at 50) to determine row positions.`
      : "";

  return [
    `Active sheet: ${ctx.activeSheet}`,
    `Sheets Metadata:\n${sheetsMeta}`,
    lastRowNote,
    `Current Selection: ${selected.address} (${selected.rowCount}x${selected.columnCount} range)`,
    `Selection Values: ${JSON.stringify(selected.values)}`,
    insightsStr,
    `Active Sheet Preview (first up to 50 rows shown — NOT the total row count):`,
    ctx.sheetData,
  ].join("\n");
}

export async function runAgent(
  userQuery: string,
  // Auth0 token + tenant are kept for signature compatibility with
  // ChatComponent. Neither is sent to the LLM proxy now.
  _token: string,
  _tenant: string,
  onStep: (status: string) => void,
  signal?: AbortSignal,
): Promise<string> {
  onStep("Reading workbook…");
  const context = await getWorkbookContext();

  if (signal?.aborted) throw new DOMException("Aborted", "AbortError");

  const userContent = [
    "Current Excel context:",
    formatContext(context),
    "",
    `User request: ${userQuery}`,
  ].join("\n");

  onStep("Thinking…");

  const tools = createTools((name) => onStep(`${name.replace(/_/g, " ")}…`));

  try {
    const result = await generateText({
      model: openai.chat(MODEL_ID),
      system: SYSTEM_PROMPT,
      messages: [{ role: "user", content: userContent }],
      tools,
      stopWhen: stepCountIs(MAX_STEPS),
      abortSignal: signal,
    });

    if (!result.text || result.text.trim().length === 0) {
      return "Done.";
    }
    return parseAIResponse(result.text).text;
  } catch (err) {
    if (err instanceof Error && err.name === "AbortError") throw err;

    const message = err instanceof Error ? err.message : "Agent failed";
    throw new Error(`Network error, please try again: ${message}.`, {
      cause: "CAUGHT_ERROR: NETWORK ERROR",
    });
  }
}
