import { getWorkbookContext, executeTool } from "./office";
import type { ToolCall, ExecutableTool, WorkbookContext } from "./office";
import { sendChatMessage } from "./chat";
import { getWorkbookInsights } from "./contextUtils";
import { parseAIResponse } from "./markdownUtils";

const MAX_STEPS = 15;

const SYSTEM_PROMPT = `You are an advanced Excel AI assistant with deep knowledge of spreadsheet operations, data analysis, and Office.js APIs.

You have access to powerful tools. To call a tool, respond with ONLY a raw JSON object — no markdown fences, no explanation.

== BASIC OPERATIONS ==
{"tool":"get_workbook_info"}
  → Returns metadata about all sheets (name, rowCount, columnCount) and active sheet

{"tool":"read_range","sheet":"Sheet1","range":"A1:C10"}
  → Reads cell values and formulas. Returns values, formulas, dimensions

{"tool":"write_cell","sheet":"Sheet1","cell":"B3","value":42}
  → Writes a single value (string, number, boolean)

{"tool":"write_range","sheet":"Sheet1","range":"A1:C2","values":[["Name","Age"],["Alice",30]],"isFormula":false}
  → Writes 2D array. Set "isFormula":true for formulas

{"tool":"add_formula","sheet":"Sheet1","cell":"D1","formula":"=SUM(A1:C1)"}
  → Writes Excel formula (must start with =)

{"tool":"clear_range","sheet":"Sheet1","range":"A1:C10"}
  → Clears content and formatting

== TABLES & FORMATTING ==
{"tool":"create_table","sheet":"Sheet1","range":"A1:C10","tableName":"SalesData","hasHeaders":true}
  → Converts range to Excel table

{"tool":"format_range","sheet":"Sheet1","range":"A1:C1","bold":true,"fillColor":"#4472C4","fontColor":"white","fontSize":12}
  → Apply formatting: numberFormat, fontColor, fillColor, bold, italic, fontSize

{"tool":"create_chart","sheet":"Sheet1","dataRange":"A1:C10","chartType":"ColumnClustered","title":"Sales Report","position":"E2"}
  → Creates chart. Types: ColumnClustered, Line, Pie, Bar, Area, XYScatter

== DATA MANIPULATION ==
{"tool":"insert_rows","sheet":"Sheet1","index":5,"count":3}
{"tool":"insert_columns","sheet":"Sheet1","index":2,"count":1}
{"tool":"delete_rows","sheet":"Sheet1","index":5,"count":2}
{"tool":"delete_columns","sheet":"Sheet1","index":3,"count":1}
  → Insert/delete rows or columns at index (0-based)

{"tool":"sort_range","sheet":"Sheet1","range":"A2:C10","sortColumn":0,"ascending":true}
  → Sort by column index (0-based)

{"tool":"filter_range","sheet":"Sheet1","range":"A1:C10","filterColumn":0,"criteria":">=100"}
  → Apply autofilter with criteria

{"tool":"auto_fill","sheet":"Sheet1","sourceRange":"A1:A3","destinationRange":"A1:A10"}
  → Auto-fill patterns (dates, numbers, series)

{"tool":"merge_cells","sheet":"Sheet1","range":"A1:C1"}
{"tool":"unmerge_cells","sheet":"Sheet1","range":"A1:C1"}
  → Merge/unmerge cells

== ADVANCED ANALYSIS ==
{"tool":"get_column_summary","sheet":"Sheet1","column":"B","range":"B2:B100"}
  → Returns statistics: sum, average, min, max, median, count, type

{"tool":"analyze_data","sheet":"Sheet1","range":"A1:D50"}
  → Deep analysis: detects headers, data types, column stats, patterns

{"tool":"detect_headers","sheet":"Sheet1","range":"A1:C10"}
  → Detects if first row contains headers

{"tool":"get_data_types","sheet":"Sheet1","range":"A1:C10"}
  → Analyzes data types per column

{"tool":"pivot_data","sheet":"Sheet1","sourceRange":"A1:D10","rowFields":["Category"],"valueField":"Sales","aggregation":"sum"}
  → Creates pivot summary. Aggregation: sum, average, count, min, max

== CONDITIONAL FORMATTING ==
{"tool":"add_conditional_format","sheet":"Sheet1","range":"B2:B10","ruleType":"CellValue","operator":"GreaterThan","formula":"100","fillColor":"#FFC7CE"}
  → Types: CellValue, ColorScale, DataBar, IconSet

== DATA VALIDATION ==
{"tool":"add_data_validation","sheet":"Sheet1","range":"A1:A100","type":"list","listSource":"Yes,No","showDropdown":true}
  → Adds in-cell dropdown. listSource is a comma-separated string of values.
{"tool":"add_data_validation","sheet":"Sheet1","range":"B1:B100","type":"whole","operator":"between","formula1":"1","formula2":"100","errorMessage":"Enter a number between 1 and 100"}
  → Validates whole numbers. Operators: between, notBetween, equalTo, notEqualTo, greaterThan, lessThan, greaterThanOrEqualTo, lessThanOrEqualTo

== NAMED RANGES ==
{"tool":"create_named_range","name":"SalesTotal","sheet":"Sheet1","range":"B2:B10"}
{"tool":"get_named_ranges"}
  → Create and retrieve named ranges

{"tool":"final_answer","answer":"I completed the task. Summary: ..."}
  → Always call last. Explain what you did clearly.

== EXCEL FORMULAS YOU SHOULD USE ==
Modern Excel functions to recommend:
- XLOOKUP(lookup_value, lookup_array, return_array)
- FILTER(array, include, [if_empty])
- SORT(array, [sort_index], [sort_order])
- UNIQUE(array, [by_col], [exactly_once])
- SUMIFS, COUNTIFS, AVERAGEIFS for conditional aggregation
- TEXT formatting: TEXT(value, format_code)
- Date functions: TODAY(), EDATE(), DATEDIF()

== INTELLIGENT BEHAVIOR ==
1. **Context awareness**: Use sheet metadata and selection to understand intent
2. **Data type detection**: Analyze data before operations (use detect_headers, get_data_types)
3. **Smart defaults**: If user says "current sheet", use active sheet from context
4. **Pattern recognition**: Detect sequences for auto-fill, headers for tables
5. **Efficient operations**: Use write_range for multiple cells, not multiple write_cell
6. **Formula intelligence**: Suggest modern Excel formulas (XLOOKUP over VLOOKUP)
7. **Data validation**: Check data structure before transformations
8. **Clear communication**: Explain what you did in final_answer

== EXAMPLES ==
User: "Format the header row blue"
→ detect_headers first, then format_range with fillColor

User: "Sort by sales column descending"
→ analyze_data to find sales column, then sort_range

User: "Create a pivot of sales by category"
→ pivot_data with appropriate fields

User: "Add a chart for this data"
→ analyze selection, create_chart with appropriate type

User: "What's the average of column B?"
→ get_column_summary, then final_answer with result

User: "Fill down the formula"
→ read source cell, use auto_fill to destination

== RULES ==
1. **Think step-by-step** before acting
2. **Analyze first** - use detect_headers, analyze_data before transformations
3. **Use context** - active sheet, selection, sheet preview in decision-making
4. **Batch operations** - use write_range for multiple cells
5. **Verify intent** - if ambiguous, make reasonable assumption based on context
6. **Modern formulas** - recommend XLOOKUP, FILTER, UNIQUE over legacy functions
7. **Always finish** - call final_answer with clear summary
8. **Read before write** - understand data structure first
9. **Smart formatting** - detect numeric/date/text and format appropriately
10. **Explain clearly** - users need to know what you did`;

function parseToolCall(text: string): ToolCall | null {
  // Step 1: Try to extract JSON from markdown code blocks first
  const markdownJsonMatch = text.match(/```(?:json)?\s*(\{[^`]*\})\s*```/s);
  if (markdownJsonMatch) {
    try {
      const parsed = JSON.parse(markdownJsonMatch[1]);
      if (parsed.tool) return parsed as ToolCall;
    } catch {
      // Continue to next method
    }
  }

  // Step 2: Try to find JSON object in plain text (original method)
  const jsonMatch = text.match(/\{[^{}]*"tool"\s*:[^{}]*\}/s);
  if (jsonMatch) {
    try {
      return JSON.parse(jsonMatch[0]) as ToolCall;
    } catch {
      // Continue to next method
    }
  }

  // Step 3: Try to extract from bold/formatted markdown
  const boldJsonMatch = text.match(/\*\*\{.*?"tool".*?\}\*\*/s);
  if (boldJsonMatch) {
    const cleaned = boldJsonMatch[0].replace(/\*\*/g, "");
    try {
      return JSON.parse(cleaned) as ToolCall;
    } catch {
      // Continue
    }
  }

  // Step 4: Try to find any JSON object with "tool" key (more aggressive)
  const aggressiveMatch = text.match(/\{[\s\S]*?"tool"\s*:[\s\S]*?\}/);
  if (aggressiveMatch) {
    try {
      return JSON.parse(aggressiveMatch[0]) as ToolCall;
    } catch {
      // Failed all attempts
    }
  }

  return null;
}

function formatContext(ctx: WorkbookContext): string {
  const selected = ctx.selectedRange;
  const sheetsMeta = ctx.sheetsMetadata.map(m => `- ${m.name}: ${m.rowCount} rows, ${m.columnCount} cols ${m.hasData ? "(contains data)" : "(empty)"}`).join("\n");

  const insights = getWorkbookInsights(ctx.sheetsMetadata);
  const insightsStr = insights.recommendations.length > 0
    ? `\nWorkbook Insights:\n${insights.recommendations.map(r => `- ${r}`).join("\n")}`
    : "";

  return [
    `Active sheet: ${ctx.activeSheet}`,
    `Sheets Metadata:\n${sheetsMeta}`,
    `Current Selection: ${selected.address} (${selected.rowCount}x${selected.columnCount} range)`,
    `Selection Values: ${JSON.stringify(selected.values)}`,
    insightsStr,
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
    const prompt = history.length > 0
      ? [
          "The following steps have already been executed on the Excel spreadsheet:",
          history.join("\n"),
          "",
          `Original user request: ${userQuery}`,
          "",
          'If the above steps fully answer the user\'s request, respond with ONLY a final_answer tool call.',
          'If more steps are still needed to complete the task, make the next tool call now.',
          'Example final_answer: {"tool":"final_answer","answer":"The active sheet is Sheet1."}',
        ].join("\n")
      : [
          SYSTEM_PROMPT,
          "",
          "Current Excel context:",
          formatContext(context),
          "",
          `User request: ${userQuery}`,
        ].join("\n").trim();

    const response = await sendChatMessage({ token, tenant, query: prompt });

    const toolCall = parseToolCall(response);

    // LLM returned a plain-text answer — parse markdown and return
    if (!toolCall) {
      const parsed = parseAIResponse(response);
      return parsed.text;
    }

    // Agent signals completion — parse markdown from final answer
    if (toolCall.tool === "final_answer") {
      const parsed = parseAIResponse(toolCall.answer);
      return parsed.text;
    }

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
