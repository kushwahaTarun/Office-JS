export interface WorkbookContext {
  activeSheet: string;
  sheets: string[];
  selectedCell: string;
  selectedValue: unknown;
  sheetData: string;
}

export type ToolCall =
  | { tool: "get_workbook_info" }
  | { tool: "read_range"; sheet: string; range: string }
  | { tool: "write_cell"; sheet: string; cell: string; value: unknown }
  | { tool: "write_range"; sheet: string; range: string; values: unknown[][] }
  | { tool: "add_formula"; sheet: string; cell: string; formula: string }
  | { tool: "final_answer"; answer: string };

export type ExecutableTool = Exclude<ToolCall, { tool: "final_answer" }>;

export async function getWorkbookContext(): Promise<WorkbookContext> {
  try {
    return await Excel.run(async (ctx) => {
      const wb = ctx.workbook;
      const activeSheet = wb.worksheets.getActiveWorksheet();
      const sheets = wb.worksheets;
      const selected = wb.getSelectedRange();
      const usedRange = activeSheet.getUsedRangeOrNullObject();

      activeSheet.load("name");
      sheets.load("items/name");
      selected.load(["address", "values"]);
      usedRange.load(["isNullObject", "values"]);

      await ctx.sync();

      let sheetData = "Empty sheet";
      if (!usedRange.isNullObject) {
        const rows = usedRange.values.slice(0, 50);
        sheetData =
          rows
            .map((row, r) =>
              row
                .slice(0, 20)
                .map((cell, c) =>
                  cell !== "" && cell !== null ? `${colLetter(c)}${r + 1}=${cell}` : ""
                )
                .filter(Boolean)
                .join("  ")
            )
            .filter(Boolean)
            .join("\n") || "Empty sheet";
      }

      return {
        activeSheet: activeSheet.name,
        sheets: sheets.items.map((s) => s.name),
        selectedCell: selected.address,
        selectedValue: selected.values?.[0]?.[0] ?? null,
        sheetData,
      };
    });
  } catch {
    return {
      activeSheet: "Sheet1",
      sheets: ["Sheet1"],
      selectedCell: "A1",
      selectedValue: null,
      sheetData: "Could not read sheet data.",
    };
  }
}

function colLetter(i: number): string {
  let s = "";
  let n = i;
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

export async function executeTool(tool: ExecutableTool): Promise<unknown> {
  switch (tool.tool) {
    case "get_workbook_info": {
      const ctx = await getWorkbookContext();
      return { activeSheet: ctx.activeSheet, sheets: ctx.sheets };
    }

    case "read_range": {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRange(tool.range);
        range.load("values");
        await ctx.sync();
        return range.values;
      });
    }

    case "write_cell": {
      await Excel.run(async (ctx) => {
        ctx.workbook.worksheets
          .getItem(tool.sheet)
          .getRange(tool.cell).values = [[tool.value]];
        await ctx.sync();
      });
      return `Wrote "${tool.value}" to ${tool.sheet}!${tool.cell}`;
    }

    case "write_range": {
      await Excel.run(async (ctx) => {
        ctx.workbook.worksheets
          .getItem(tool.sheet)
          .getRange(tool.range).values = tool.values;
        await ctx.sync();
      });
      return `Wrote ${tool.values.length} row(s) to ${tool.sheet}!${tool.range}`;
    }

    case "add_formula": {
      await Excel.run(async (ctx) => {
        ctx.workbook.worksheets
          .getItem(tool.sheet)
          .getRange(tool.cell).formulas = [[tool.formula]];
        await ctx.sync();
      });
      return `Formula ${tool.formula} added to ${tool.sheet}!${tool.cell}`;
    }
  }
}
