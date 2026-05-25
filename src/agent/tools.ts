import { tool } from "ai";
import { z } from "zod";
import { executeTool } from "../services/office";
import type { ExecutableTool } from "../services/office";

const A1Range = z
  .string()
  .regex(/^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/, "Must be A1 or A1:B2 notation (no sheet prefix)");

const A1Cell = z
  .string()
  .regex(/^[A-Za-z]+\d+$/, "Must be a single A1 cell reference (e.g. B3)");

const ChartType = z.enum(["ColumnClustered", "Line", "Pie", "Bar", "Area", "XYScatter"]);

// Factory so each tool's execute can notify the UI of the active tool
// name before running it. Without this, the UI can only show "Thinking…"
// for the entire multi-step run.
export function createTools(onToolStart: (name: string) => void) {
  const run = async <T extends ExecutableTool["tool"]>(
    name: T,
    args: Omit<Extract<ExecutableTool, { tool: T }>, "tool">,
  ): Promise<unknown> => {
    onToolStart(name);
    return executeTool({ tool: name, ...args } as ExecutableTool);
  };

  return {
    get_workbook_info: tool({
      description:
        "Return metadata about all sheets (name, rowCount, columnCount, hasData) and the active sheet name.",
      inputSchema: z.object({}),
      execute: async () => run("get_workbook_info", {} as never),
    }),

    read_range: tool({
      description:
        "Read cell values and formulas for a range. Returns values, formulas, and dimensions.",
      inputSchema: z.object({ sheet: z.string(), range: A1Range }),
      execute: async (input) => run("read_range", input),
    }),

    write_cell: tool({
      description: "Write a single value (string, number, or boolean) to one cell.",
      inputSchema: z.object({
        sheet: z.string(),
        cell: A1Cell,
        value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
      }),
      execute: async (input) => run("write_cell", input),
    }),

    write_range: tool({
      description:
        "Write a 2D array to a range. Set isFormula=true to write the strings as formulas instead of values.",
      inputSchema: z.object({
        sheet: z.string(),
        range: A1Range,
        values: z.array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()]))),
        isFormula: z.boolean().optional(),
      }),
      execute: async (input) => run("write_range", input),
    }),

    add_formula: tool({
      description: "Write a single Excel formula (must start with '=') to one cell.",
      inputSchema: z.object({
        sheet: z.string(),
        cell: A1Cell,
        formula: z.string().regex(/^=/, "Formula must start with '='"),
      }),
      execute: async (input) => run("add_formula", input),
    }),

    clear_range: tool({
      description: "Clear values and formatting from a range.",
      inputSchema: z.object({ sheet: z.string(), range: A1Range }),
      execute: async (input) => run("clear_range", input),
    }),

    create_table: tool({
      description: "Convert a range into an Excel table.",
      inputSchema: z.object({
        sheet: z.string(),
        range: A1Range,
        tableName: z.string().optional(),
        hasHeaders: z.boolean().optional(),
      }),
      execute: async (input) => run("create_table", input),
    }),

    format_range: tool({
      description:
        "Apply formatting to a range: numberFormat, fontColor, fillColor, bold, italic, fontSize.",
      inputSchema: z.object({
        sheet: z.string(),
        range: A1Range,
        numberFormat: z.string().optional(),
        fontColor: z.string().optional(),
        fillColor: z.string().optional(),
        bold: z.boolean().optional(),
        italic: z.boolean().optional(),
        fontSize: z.number().optional(),
      }),
      execute: async (input) => run("format_range", input),
    }),

    create_chart: tool({
      description:
        "Create a chart from a data range. chartType is one of ColumnClustered, Line, Pie, Bar, Area, XYScatter. Note: 'Bar' is unsupported by the host — use ColumnClustered for bar/column charts.",
      inputSchema: z.object({
        sheet: z.string(),
        dataRange: A1Range,
        chartType: ChartType,
        title: z.string().optional(),
        position: A1Cell.optional(),
      }),
      execute: async (input) => run("create_chart", input),
    }),

    insert_rows: tool({
      description: "Insert N rows at the given 0-based index.",
      inputSchema: z.object({
        sheet: z.string(),
        index: z.number().int().min(0),
        count: z.number().int().min(1),
      }),
      execute: async (input) => run("insert_rows", input),
    }),

    insert_columns: tool({
      description: "Insert N columns at the given 0-based index.",
      inputSchema: z.object({
        sheet: z.string(),
        index: z.number().int().min(0),
        count: z.number().int().min(1),
      }),
      execute: async (input) => run("insert_columns", input),
    }),

    delete_rows: tool({
      description: "Delete N rows starting at the given 0-based index.",
      inputSchema: z.object({
        sheet: z.string(),
        index: z.number().int().min(0),
        count: z.number().int().min(1),
      }),
      execute: async (input) => run("delete_rows", input),
    }),

    delete_columns: tool({
      description: "Delete N columns starting at the given 0-based index.",
      inputSchema: z.object({
        sheet: z.string(),
        index: z.number().int().min(0),
        count: z.number().int().min(1),
      }),
      execute: async (input) => run("delete_columns", input),
    }),

    sort_range: tool({
      description: "Sort a range by a 0-based column index. Defaults to ascending.",
      inputSchema: z.object({
        sheet: z.string(),
        range: A1Range,
        sortColumn: z.number().int().min(0),
        ascending: z.boolean().optional(),
      }),
      execute: async (input) => run("sort_range", input),
    }),

    filter_range: tool({
      description:
        "Apply an autofilter to a range with a criteria string (e.g. '>=100', '*foo*').",
      inputSchema: z.object({
        sheet: z.string(),
        range: A1Range,
        filterColumn: z.number().int().min(0),
        criteria: z.string(),
      }),
      execute: async (input) => run("filter_range", input),
    }),

    get_column_summary: tool({
      description:
        "Compute statistics (sum, average, min, max, median, count, type) for a column. Provide the column letter and an optional explicit range.",
      inputSchema: z.object({
        sheet: z.string(),
        column: z.string(),
        range: z.string().optional(),
      }),
      execute: async (input) => run("get_column_summary", input),
    }),

    auto_fill: tool({
      description:
        "Auto-fill values from a source range to a destination range (sequences, dates, formulas).",
      inputSchema: z.object({
        sheet: z.string(),
        sourceRange: A1Range,
        destinationRange: A1Range,
      }),
      execute: async (input) => run("auto_fill", input),
    }),

    merge_cells: tool({
      description: "Merge cells across the given range.",
      inputSchema: z.object({ sheet: z.string(), range: A1Range }),
      execute: async (input) => run("merge_cells", input),
    }),

    unmerge_cells: tool({
      description: "Unmerge previously merged cells in the given range.",
      inputSchema: z.object({ sheet: z.string(), range: A1Range }),
      execute: async (input) => run("unmerge_cells", input),
    }),

    add_conditional_format: tool({
      description:
        "Add conditional formatting. ruleType: CellValue, ColorScale, DataBar, IconSet, TopBottom. For CellValue supply operator + formula. For TopBottom supply topBottomType (TopItems/BottomItems/TopPercent/BottomPercent) and rank.",
      inputSchema: z.object({
        sheet: z.string(),
        range: A1Range,
        ruleType: z.enum(["CellValue", "ColorScale", "DataBar", "IconSet", "TopBottom"]),
        operator: z.string().optional(),
        formula: z.string().optional(),
        fillColor: z.string().optional(),
        rank: z.number().optional(),
        topBottomType: z
          .enum(["TopItems", "BottomItems", "TopPercent", "BottomPercent"])
          .optional(),
      }),
      execute: async (input) => run("add_conditional_format", input),
    }),

    create_named_range: tool({
      description: "Create a workbook-level named range.",
      inputSchema: z.object({ name: z.string(), sheet: z.string(), range: A1Range }),
      execute: async (input) => run("create_named_range", input),
    }),

    get_named_ranges: tool({
      description: "Return all workbook-level named ranges.",
      inputSchema: z.object({}),
      execute: async () => run("get_named_ranges", {} as never),
    }),

    analyze_data: tool({
      description:
        "Deep analysis of a range: detects headers, data types per column, and basic per-column statistics.",
      inputSchema: z.object({ sheet: z.string(), range: A1Range }),
      execute: async (input) => run("analyze_data", input),
    }),

    detect_headers: tool({
      description: "Heuristically detect whether the first row of a range is a header row.",
      inputSchema: z.object({ sheet: z.string(), range: z.string().optional() }),
      execute: async (input) => run("detect_headers", input),
    }),

    get_data_types: tool({
      description: "Per-column data type detection for a range.",
      inputSchema: z.object({ sheet: z.string(), range: A1Range }),
      execute: async (input) => run("get_data_types", input),
    }),

    pivot_data: tool({
      description:
        "Compute a pivot summary of a source range. Aggregation: sum (default), average, count, min, max.",
      inputSchema: z.object({
        sheet: z.string(),
        sourceRange: A1Range,
        rowFields: z.array(z.string()).min(1),
        columnFields: z.array(z.string()).optional(),
        valueField: z.string(),
        aggregation: z.enum(["sum", "average", "count", "min", "max"]).optional(),
      }),
      execute: async (input) => run("pivot_data", input),
    }),

    add_data_validation: tool({
      description:
        "Add data validation to a range. type=list uses listSource as a comma-separated string. type=whole/decimal/date/textLength uses operator + formula1/formula2.",
      inputSchema: z.object({
        sheet: z.string(),
        range: A1Range,
        type: z.enum(["list", "whole", "decimal", "date", "textLength"]),
        listSource: z.string().optional(),
        operator: z.string().optional(),
        formula1: z.string().optional(),
        formula2: z.string().optional(),
        showDropdown: z.boolean().optional(),
        errorMessage: z.string().optional(),
      }),
      execute: async (input) => run("add_data_validation", input),
    }),

    create_sheet: tool({
      description:
        "Create a new worksheet. If a sheet with that name already exists the call returns a notice — do NOT retry.",
      inputSchema: z.object({
        name: z.string(),
        position: z.number().int().min(0).optional(),
      }),
      execute: async (input) => run("create_sheet", input),
    }),

    copy_range: tool({
      description:
        "Copy values from one sheet/range into another sheet starting at the target cell.",
      inputSchema: z.object({
        sourceSheet: z.string(),
        sourceRange: A1Range,
        targetSheet: z.string(),
        targetCell: A1Cell,
      }),
      execute: async (input) => run("copy_range", input),
    }),
  };
}
