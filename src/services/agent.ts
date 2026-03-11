import { getWorkbookContext, executeTool } from "./office";
import type { ToolCall, ExecutableTool, WorkbookContext } from "./office";
import { sendChatMessage } from "./chat";

const MAX_STEPS = 8;

const SYSTEM_PROMPT = `You are an Excel AI assistant. You help users read and modify their spreadsheets.

You have access to the following tools. To call a tool, respond with ONLY a raw JSON object — no markdown fences, no explanation before or after.

{"tool":"get_workbook_info"}
  → Returns detailed metadata about all sheets (name, rowCount, columnCount) and the active sheet.

{"tool":"read_range","sheet":"Sheet1","range":"A1:C10"}
  → Reads cell values and formulas from the given range. Returns values, formulas, and dimensions.

{"tool":"write_cell","sheet":"Sheet1","cell":"B3","value":42}
  → Writes a single value (string, number, or boolean) to a cell.

{"tool":"write_range","sheet":"Sheet1","range":"A1:C2","values":[["Name","Age","City"],["Alice",30,"NYC"]], "isFormula": false}
  → Writes a 2D array of values or formulas to a range. If "isFormula" is true, values are treated as formulas.

{"tool":"add_formula","sheet":"Sheet1","cell":"D1","formula":"=SUM(A1:C1)"}
  → Writes an Excel formula to a cell. Always prefix with =.

{"tool":"clear_range","sheet":"Sheet1","range":"A1:C10"}
  → Clears content and formatting from the given range.

{"tool":"final_answer","answer":"I completed the task. Here is what I did: ..."}
  → Ends the task. Always call this last with a clear summary for the user.

Rules:
1. Think step-by-step before acting.
2. Read data first if you need to understand the sheet structure before modifying.
3. Use the provided metadata (active sheet, selected range) to contextualize the user's request.
4. If writing values (like numbers/strings), check if the target range is empty first by reading it, unless the user explicitly asks to overwrite.
5. Always call final_answer when done — even if no changes were made.
6. If the user only asks a question (no sheet changes needed), call final_answer directly.`;

function parseToolCall(text: string): ToolCall | null {
  // Match outermost JSON object containing a "tool" key
  const match = text.match(/\{[^{}]*"tool"\s*:[^{}]*\}/s);
  if (!match) return null;
  try {
    return JSON.parse(match[0]) as ToolCall;
  } catch {
    return null;
  }
}

function formatContext(ctx: WorkbookContext): string {
  const selected = ctx.selectedRange;
  const sheetsMeta = ctx.sheetsMetadata.map(m => `- ${m.name}: ${m.rowCount} rows, ${m.columnCount} cols ${m.hasData ? "(contains data)" : "(empty)"}`).join("\n");

  return [
    `Active sheet: ${ctx.activeSheet}`,
    `Sheets Metadata:\n${sheetsMeta}`,
    `Current Selection: ${selected.address} (${selected.rowCount}x${selected.columnCount} range)`,
    `Selection Values: ${JSON.stringify(selected.values)}`,
    `Active Sheet Preview (Top 50 rows, 20 cols):`,
    ctx.sheetData,
  ].join("\n");
}

export async function runAgent(
  userQuery: string,
  token: string,
  tenant: string,
  onStep: (status: string) => void
): Promise<string> {
  onStep("Reading workbook…");
  const context = await getWorkbookContext();

  const history: string[] = [];

  for (let step = 0; step < MAX_STEPS; step++) {
    const prompt = [
      SYSTEM_PROMPT,
      "",
      "Current Excel context:",
      formatContext(context),
      history.length > 0 ? "\nSteps taken so far:\n" + history.join("\n") : "",
      "",
      `User request: ${userQuery}`,
    ]
      .join("\n")
      .trim();

    const response = await sendChatMessage({ token, tenant, query: prompt });

    const toolCall = parseToolCall(response);

    // LLM returned a plain-text answer — use it directly
    if (!toolCall) return response;

    // Agent signals completion
    if (toolCall.tool === "final_answer") return toolCall.answer;

    // Execute the tool and record the result
    const label = toolCall.tool.replace(/_/g, " ");
    onStep(`${label}…`);

    try {
      const result = await executeTool(toolCall as ExecutableTool);
      history.push(
        `Step ${step + 1}: ${JSON.stringify(toolCall)} → ${JSON.stringify(result)}`
      );
    } catch (err) {
      const msg = err instanceof Error ? err.message : "Tool execution failed";
      history.push(`Step ${step + 1}: ${JSON.stringify(toolCall)} → Error: ${msg}`);
    }
  }

  return "I completed several steps but reached the maximum iteration limit. Please check your spreadsheet for the changes made.";
}
