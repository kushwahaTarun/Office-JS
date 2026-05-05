// Single source of truth for tool definitions.
// Each tool declares an input schema (Zod). No `execute` is provided —
// tools are executed on the client (Office.js add-in), so the SDK pauses
// after emitting a tool call and the client posts back a tool result on
// the next step.

const { tool } = require("ai");
const { z } = require("zod");

const A1Range = z
  .string()
  .regex(/^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/, "Must be A1 or A1:B2 notation (no sheet prefix)");

const A1Cell = z
  .string()
  .regex(/^[A-Za-z]+\d+$/, "Must be a single A1 cell reference (e.g. B3)");

const ChartType = z.enum(["ColumnClustered", "Line", "Pie", "Bar", "Area", "XYScatter"]);

const tools = {
  get_workbook_info: tool({
    description:
      "Return metadata about all sheets (name, rowCount, columnCount, hasData) and the active sheet name.",
    inputSchema: z.object({}),
  }),

  read_range: tool({
    description: "Read cell values and formulas for a range. Returns values, formulas, and dimensions.",
    inputSchema: z.object({
      sheet: z.string(),
      range: A1Range,
    }),
  }),

  write_cell: tool({
    description: "Write a single value (string, number, or boolean) to one cell.",
    inputSchema: z.object({
      sheet: z.string(),
      cell: A1Cell,
      value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
    }),
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
  }),

  add_formula: tool({
    description: "Write a single Excel formula (must start with '=') to one cell.",
    inputSchema: z.object({
      sheet: z.string(),
      cell: A1Cell,
      formula: z.string().regex(/^=/, "Formula must start with '='"),
    }),
  }),

  clear_range: tool({
    description: "Clear values and formatting from a range.",
    inputSchema: z.object({
      sheet: z.string(),
      range: A1Range,
    }),
  }),

  create_table: tool({
    description: "Convert a range into an Excel table.",
    inputSchema: z.object({
      sheet: z.string(),
      range: A1Range,
      tableName: z.string().optional(),
      hasHeaders: z.boolean().optional(),
    }),
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
  }),

  insert_rows: tool({
    description: "Insert N rows at the given 0-based index.",
    inputSchema: z.object({
      sheet: z.string(),
      index: z.number().int().min(0),
      count: z.number().int().min(1),
    }),
  }),

  insert_columns: tool({
    description: "Insert N columns at the given 0-based index.",
    inputSchema: z.object({
      sheet: z.string(),
      index: z.number().int().min(0),
      count: z.number().int().min(1),
    }),
  }),

  delete_rows: tool({
    description: "Delete N rows starting at the given 0-based index.",
    inputSchema: z.object({
      sheet: z.string(),
      index: z.number().int().min(0),
      count: z.number().int().min(1),
    }),
  }),

  delete_columns: tool({
    description: "Delete N columns starting at the given 0-based index.",
    inputSchema: z.object({
      sheet: z.string(),
      index: z.number().int().min(0),
      count: z.number().int().min(1),
    }),
  }),

  sort_range: tool({
    description: "Sort a range by a 0-based column index. Defaults to ascending.",
    inputSchema: z.object({
      sheet: z.string(),
      range: A1Range,
      sortColumn: z.number().int().min(0),
      ascending: z.boolean().optional(),
    }),
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
  }),

  get_column_summary: tool({
    description:
      "Compute statistics (sum, average, min, max, median, count, type) for a column. Provide the column letter and an optional explicit range.",
    inputSchema: z.object({
      sheet: z.string(),
      column: z.string(),
      range: z.string().optional(),
    }),
  }),

  auto_fill: tool({
    description:
      "Auto-fill values from a source range to a destination range (sequences, dates, formulas).",
    inputSchema: z.object({
      sheet: z.string(),
      sourceRange: A1Range,
      destinationRange: A1Range,
    }),
  }),

  merge_cells: tool({
    description: "Merge cells across the given range.",
    inputSchema: z.object({
      sheet: z.string(),
      range: A1Range,
    }),
  }),

  unmerge_cells: tool({
    description: "Unmerge previously merged cells in the given range.",
    inputSchema: z.object({
      sheet: z.string(),
      range: A1Range,
    }),
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
      topBottomType: z.enum(["TopItems", "BottomItems", "TopPercent", "BottomPercent"]).optional(),
    }),
  }),

  create_named_range: tool({
    description: "Create a workbook-level named range.",
    inputSchema: z.object({
      name: z.string(),
      sheet: z.string(),
      range: A1Range,
    }),
  }),

  get_named_ranges: tool({
    description: "Return all workbook-level named ranges.",
    inputSchema: z.object({}),
  }),

  analyze_data: tool({
    description:
      "Deep analysis of a range: detects headers, data types per column, and basic per-column statistics.",
    inputSchema: z.object({
      sheet: z.string(),
      range: A1Range,
    }),
  }),

  detect_headers: tool({
    description: "Heuristically detect whether the first row of a range is a header row.",
    inputSchema: z.object({
      sheet: z.string(),
      range: z.string().optional(),
    }),
  }),

  get_data_types: tool({
    description: "Per-column data type detection for a range.",
    inputSchema: z.object({
      sheet: z.string(),
      range: A1Range,
    }),
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
  }),

  create_sheet: tool({
    description:
      "Create a new worksheet. If a sheet with that name already exists the call returns a notice — do NOT retry.",
    inputSchema: z.object({
      name: z.string(),
      position: z.number().int().min(0).optional(),
    }),
  }),

  copy_range: tool({
    description: "Copy values from one sheet/range into another sheet starting at the target cell.",
    inputSchema: z.object({
      sourceSheet: z.string(),
      sourceRange: A1Range,
      targetSheet: z.string(),
      targetCell: A1Cell,
    }),
  }),
};

module.exports = { tools };
