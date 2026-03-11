export interface SheetMetadata {
  name: string;
  rowCount: number;
  columnCount: number;
  hasData: boolean;
}

export interface WorkbookContext {
  activeSheet: string;
  sheets: string[];
  sheetsMetadata: SheetMetadata[];
  selectedRange: {
    address: string;
    rowCount: number;
    columnCount: number;
    values: unknown[][];
  };
  sheetData: string;
}

export type ToolCall =
  | { tool: "get_workbook_info" }
  | { tool: "read_range"; sheet: string; range: string }
  | { tool: "write_cell"; sheet: string; cell: string; value: unknown }
  | { tool: "write_range"; sheet: string; range: string; values: unknown[][]; isFormula?: boolean }
  | { tool: "add_formula"; sheet: string; cell: string; formula: string }
  | { tool: "clear_range"; sheet: string; range: string }
  | { tool: "final_answer"; answer: string };

export type ExecutableTool = Exclude<ToolCall, { tool: "final_answer" }>;

export async function getWorkbookContext(): Promise<WorkbookContext> {
  try {
    return await Excel.run(async (ctx) => {
      const wb = ctx.workbook;
      const activeSheet = wb.worksheets.getActiveWorksheet();
      const sheets = wb.worksheets;
      const selected = wb.getSelectedRange();

      activeSheet.load("name");
      sheets.load("items/name");
      selected.load(["address", "values", "rowCount", "columnCount"]);

      await ctx.sync();

      // Load metadata for all sheets
      const sheetsMetadata: SheetMetadata[] = [];
      for (const sheet of sheets.items) {
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load(["address", "rowCount", "columnCount", "isNullObject"]);
        await ctx.sync();

        sheetsMetadata.push({
          name: sheet.name,
          rowCount: usedRange.isNullObject ? 0 : usedRange.rowCount,
          columnCount: usedRange.isNullObject ? 0 : usedRange.columnCount,
          hasData: !usedRange.isNullObject
        });
      }

      const activeUsedRange = activeSheet.getUsedRangeOrNullObject();
      activeUsedRange.load(["isNullObject", "values"]);
      await ctx.sync();

      let sheetData = "Empty sheet";
      if (!activeUsedRange.isNullObject) {
        const rows = activeUsedRange.values.slice(0, 50);
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
        sheetsMetadata,
        selectedRange: {
          address: selected.address,
          rowCount: selected.rowCount,
          columnCount: selected.columnCount,
          values: selected.values,
        },
        sheetData,
      };
    });
  } catch (error) {
    console.error("Error getting workbook context:", error);
    return {
      activeSheet: "Sheet1",
      sheets: ["Sheet1"],
      sheetsMetadata: [{ name: "Sheet1", rowCount: 1, columnCount: 1, hasData: false }],
      selectedRange: {
        address: "A1",
        rowCount: 1,
        columnCount: 1,
        values: [[null]],
      },
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
        range.load(["values", "formulas", "address", "rowCount", "columnCount"]);
        await ctx.sync();
        return {
          address: range.address,
          values: range.values,
          formulas: range.formulas,
          dimensions: { rows: range.rowCount, cols: range.columnCount }
        };
      });
    }

    case "write_cell": {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.worksheets
          .getItem(tool.sheet)
          .getRange(tool.cell);

        // Basic overwrite protection: do not write if user didn't ask to overwrite and cell is not empty
        // (Implementation detail: for now we just write, but we could add a check here)
        range.values = [[tool.value]];
        await ctx.sync();
      });
      return `Wrote "${tool.value}" to ${tool.sheet}!${tool.cell}`;
    }

    case "write_range": {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.worksheets
          .getItem(tool.sheet)
          .getRange(tool.range);

        if (tool.isFormula) {
          range.formulas = tool.values as string[][];
        } else {
          range.values = tool.values;
        }
        await ctx.sync();
      });
      return `Wrote ${tool.values.length} row(s) to ${tool.sheet}!${tool.range}${tool.isFormula ? " (as formulas)" : ""}`;
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

    case "clear_range": {
      await Excel.run(async (ctx) => {
        ctx.workbook.worksheets
          .getItem(tool.sheet)
          .getRange(tool.range)
          .clear();
        await ctx.sync();
      });
      return `Cleared range ${tool.sheet}!${tool.range}`;
    }
  }
}
