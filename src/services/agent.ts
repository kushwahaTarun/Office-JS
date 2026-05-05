import { getWorkbookContext, executeTool } from "./office";
import type { ExecutableTool, WorkbookContext } from "./office";
import { getWorkbookInsights } from "./contextUtils";
import { parseAIResponse } from "./markdownUtils";
import { postAgentStep } from "./agentClient";
import type { AgentMessage, AgentToolCall } from "./agentClient";

const MAX_STEPS = 15;

function formatContext(ctx: WorkbookContext): string {
  const selected = ctx.selectedRange;
  const activeSheet = ctx.sheetsMetadata.find((m) => m.name === ctx.activeSheet);
  const sheetsMeta = ctx.sheetsMetadata
    .map(
      (m) =>
        `- ${m.name}: ${m.rowCount} rows, ${m.columnCount} cols ${m.hasData ? "(contains data)" : "(empty)"}`
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

function toToolResultMessage(
  call: AgentToolCall,
  output: unknown,
  isError = false
): AgentMessage {
  return {
    role: "tool",
    content: [
      {
        type: "tool-result",
        toolCallId: call.toolCallId,
        toolName: call.toolName,
        output: isError
          ? { type: "error-text", value: String(output) }
          : { type: "json", value: output as unknown },
      },
    ],
  };
}

export async function runAgent(
  userQuery: string,
  token: string,
  // tenant kept for signature compatibility with the legacy call site;
  // the harness no longer needs it (auth is via the JWT alone).
  _tenant: string,
  onStep: (status: string) => void,
  signal?: AbortSignal
): Promise<string> {
  onStep("Reading workbook…");
  const context = await getWorkbookContext();

  const initialUserContent = [
    "Current Excel context:",
    formatContext(context),
    "",
    `User request: ${userQuery}`,
  ].join("\n");

  let messages: AgentMessage[] = [{ role: "user", content: initialUserContent }];

  for (let step = 0; step < MAX_STEPS; step++) {
    if (signal?.aborted) throw new DOMException("Aborted", "AbortError");

    onStep(step === 0 ? "Thinking…" : "Continuing…");
    const response = await postAgentStep({ token, messages, signal });
    messages = response.messages;

    if (response.finalText !== null) {
      return parseAIResponse(response.finalText).text;
    }

    if (response.toolCalls.length === 0) {
      // No tool calls and no final text — treat as done with empty answer
      return "Done.";
    }

    const toolMessages: AgentMessage[] = [];
    for (const call of response.toolCalls) {
      if (signal?.aborted) throw new DOMException("Aborted", "AbortError");

      onStep(`${call.toolName.replace(/_/g, " ")}…`);

      const executable = {
        tool: call.toolName,
        ...call.input,
      } as unknown as ExecutableTool;

      try {
        const result = await executeTool(executable);
        toolMessages.push(toToolResultMessage(call, result));
      } catch (err) {
        const msg = err instanceof Error ? err.message : "Tool execution failed";
        toolMessages.push(toToolResultMessage(call, msg, true));
      }
    }

    messages = [...messages, ...toolMessages];
  }

  return "I completed several steps but reached the maximum iteration limit. Please check your spreadsheet for the changes made.";
}
