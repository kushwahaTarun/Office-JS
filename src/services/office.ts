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
  | { tool: "create_table"; sheet: string; range: string; tableName?: string; hasHeaders?: boolean }
  | { tool: "format_range"; sheet: string; range: string; numberFormat?: string; fontColor?: string; fillColor?: string; bold?: boolean; italic?: boolean; fontSize?: number }
  | { tool: "create_chart"; sheet: string; dataRange: string; chartType: "ColumnClustered" | "Line" | "Pie" | "Bar" | "Area" | "XYScatter"; title?: string; position?: string }
  | { tool: "insert_rows"; sheet: string; index: number; count: number }
  | { tool: "insert_columns"; sheet: string; index: number; count: number }
  | { tool: "delete_rows"; sheet: string; index: number; count: number }
  | { tool: "delete_columns"; sheet: string; index: number; count: number }
  | { tool: "sort_range"; sheet: string; range: string; sortColumn: number; ascending?: boolean }
  | { tool: "filter_range"; sheet: string; range: string; filterColumn: number; criteria: string }
  | { tool: "get_column_summary"; sheet: string; column: string; range?: string }
  | { tool: "auto_fill"; sheet: string; sourceRange: string; destinationRange: string }
  | { tool: "merge_cells"; sheet: string; range: string }
  | { tool: "unmerge_cells"; sheet: string; range: string }
  | { tool: "add_conditional_format"; sheet: string; range: string; ruleType: "CellValue" | "ColorScale" | "DataBar" | "IconSet"; operator?: string; formula?: string; fillColor?: string }
  | { tool: "create_named_range"; name: string; sheet: string; range: string }
  | { tool: "get_named_ranges" }
  | { tool: "analyze_data"; sheet: string; range: string }
  | { tool: "detect_headers"; sheet: string; range?: string }
  | { tool: "get_data_types"; sheet: string; range: string }
  | { tool: "pivot_data"; sheet: string; sourceRange: string; rowFields: string[]; columnFields?: string[]; valueField: string; aggregation?: "sum" | "average" | "count" | "min" | "max" }
  | { tool: "final_answer"; answer: string };

export type ExecutableTool = Exclude<ToolCall, { tool: "final_answer" }>;

export async function getWorkbookContext(): Promise<WorkbookContext> {
  // Check if Excel API is available
  if (typeof Excel === "undefined") {
    console.error("Excel API is not available. Are you running in Excel?");
    return {
      activeSheet: "Sheet1",
      sheets: ["Sheet1"],
      sheetsMetadata: [{ name: "Sheet1", rowCount: 0, columnCount: 0, hasData: false }],
      selectedRange: {
        address: "A1",
        rowCount: 1,
        columnCount: 1,
        values: [[""]],
      },
      sheetData: "Excel API not available. Please run this add-in in Excel.",
    };
  }

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
  // Check if Excel API is available
  if (typeof Excel === "undefined") {
    throw new Error("Excel API is not available. This add-in must be run in Microsoft Excel.");
  }

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

    case "create_table": {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRange(tool.range);
        const table = ws.tables.add(range, tool.hasHeaders ?? true);
        if (tool.tableName) {
          table.name = tool.tableName;
        }
        table.load("name");
        await ctx.sync();
        return `Created table ${table.name} from ${tool.range} on ${tool.sheet}`;
      });
    }

    case "format_range": {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.worksheets
          .getItem(tool.sheet)
          .getRange(tool.range);

        if (tool.numberFormat) {
          range.numberFormat = [[tool.numberFormat]];
        }
        if (tool.fontColor) {
          range.format.font.color = tool.fontColor;
        }
        if (tool.fillColor) {
          range.format.fill.color = tool.fillColor;
        }
        if (tool.bold !== undefined) {
          range.format.font.bold = tool.bold;
        }
        if (tool.italic !== undefined) {
          range.format.font.italic = tool.italic;
        }
        if (tool.fontSize) {
          range.format.font.size = tool.fontSize;
        }

        await ctx.sync();
      });
      return `Formatted range ${tool.sheet}!${tool.range}`;
    }

    case "create_chart": {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRange(tool.dataRange);
        const chart = ws.charts.add(tool.chartType as Excel.ChartType, range, "Auto");

        if (tool.title) {
          chart.title.text = tool.title;
        }
        if (tool.position) {
          chart.setPosition(tool.position);
        }

        chart.load("name");
        await ctx.sync();
        return `Created ${tool.chartType} chart from ${tool.dataRange} on ${tool.sheet}`;
      });
    }

    case "insert_rows": {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRangeByIndexes(tool.index, 0, tool.count, 1);
        range.insert("Down");
        await ctx.sync();
      });
      return `Inserted ${tool.count} row(s) at index ${tool.index} on ${tool.sheet}`;
    }

    case "insert_columns": {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRangeByIndexes(0, tool.index, 1, tool.count);
        range.insert("Right");
        await ctx.sync();
      });
      return `Inserted ${tool.count} column(s) at index ${tool.index} on ${tool.sheet}`;
    }

    case "delete_rows": {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRangeByIndexes(tool.index, 0, tool.count, 1);
        range.delete("Up");
        await ctx.sync();
      });
      return `Deleted ${tool.count} row(s) at index ${tool.index} on ${tool.sheet}`;
    }

    case "delete_columns": {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRangeByIndexes(0, tool.index, 1, tool.count);
        range.delete("Left");
        await ctx.sync();
      });
      return `Deleted ${tool.count} column(s) at index ${tool.index} on ${tool.sheet}`;
    }

    case "sort_range": {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRange(tool.range);
        range.sort.apply(
          [{ key: tool.sortColumn, ascending: tool.ascending ?? true }],
          false
        );
        await ctx.sync();
      });
      return `Sorted ${tool.range} by column ${tool.sortColumn} (${tool.ascending ?? true ? "ascending" : "descending"})`;
    }

    case "filter_range": {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRange(tool.range);
        const autoFilter = ws.autoFilter;
        autoFilter.apply(range);
        autoFilter.apply(range, tool.filterColumn, {
          criterion1: tool.criteria,
          filterOn: Excel.FilterOn.values
        });
        await ctx.sync();
      });
      return `Applied filter to column ${tool.filterColumn} in ${tool.range} with criteria: ${tool.criteria}`;
    }

    case "get_column_summary": {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const rangeStr = tool.range || `${tool.column}:${tool.column}`;
        const range = ws.getRange(rangeStr);
        range.load(["values", "rowCount"]);
        await ctx.sync();

        const values = range.values.flat().filter(v => v !== null && v !== "");
        const numbers = values.filter(v => typeof v === "number") as number[];

        if (numbers.length === 0) {
          return {
            column: tool.column,
            count: values.length,
            type: "text",
            uniqueValues: [...new Set(values)].length
          };
        }

        const sum = numbers.reduce((a, b) => a + b, 0);
        const avg = sum / numbers.length;
        const sorted = [...numbers].sort((a, b) => a - b);
        const min = sorted[0];
        const max = sorted[sorted.length - 1];
        const median = sorted[Math.floor(sorted.length / 2)];

        return {
          column: tool.column,
          count: numbers.length,
          sum,
          average: avg,
          min,
          max,
          median,
          type: "numeric"
        };
      });
    }

    case "auto_fill": {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const sourceRange = ws.getRange(tool.sourceRange);
        const destRange = ws.getRange(tool.destinationRange);
        sourceRange.autoFill(destRange, "FillDefault");
        await ctx.sync();
      });
      return `Auto-filled from ${tool.sourceRange} to ${tool.destinationRange}`;
    }

    case "merge_cells": {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.worksheets
          .getItem(tool.sheet)
          .getRange(tool.range);
        range.merge(false);
        await ctx.sync();
      });
      return `Merged cells in ${tool.range}`;
    }

    case "unmerge_cells": {
      await Excel.run(async (ctx) => {
        const range = ctx.workbook.worksheets
          .getItem(tool.sheet)
          .getRange(tool.range);
        range.unmerge();
        await ctx.sync();
      });
      return `Unmerged cells in ${tool.range}`;
    }

    case "add_conditional_format": {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRange(tool.range);

        let format: Excel.ConditionalFormat;

        if (tool.ruleType === "CellValue" && tool.operator && tool.formula) {
          format = range.conditionalFormats.add("CellValue");
          const cellFormat = format.cellValue;
          cellFormat.rule = {
            formula1: tool.formula,
            operator: tool.operator as Excel.ConditionalCellValueOperator
          };
          if (tool.fillColor) {
            cellFormat.format.fill.color = tool.fillColor;
          }
        } else if (tool.ruleType === "ColorScale") {
          format = range.conditionalFormats.add("ColorScale");
        } else if (tool.ruleType === "DataBar") {
          format = range.conditionalFormats.add("DataBar");
        } else if (tool.ruleType === "IconSet") {
          format = range.conditionalFormats.add("IconSet");
        } else {
          throw new Error(`Unsupported rule type: ${tool.ruleType}`);
        }

        await ctx.sync();
        return `Added ${tool.ruleType} conditional formatting to ${tool.range}`;
      });
    }

    case "create_named_range": {
      await Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRange(tool.range);
        ctx.workbook.names.add(tool.name, range);
        await ctx.sync();
      });
      return `Created named range "${tool.name}" for ${tool.sheet}!${tool.range}`;
    }

    case "get_named_ranges": {
      return Excel.run(async (ctx) => {
        const names = ctx.workbook.names;
        names.load("items/name, items/formula");
        await ctx.sync();
        return names.items.map(n => ({ name: n.name, formula: n.formula }));
      });
    }

    case "analyze_data": {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRange(tool.range);
        range.load(["values", "rowCount", "columnCount"]);
        await ctx.sync();

        const values = range.values;
        const rowCount = values.length;
        const colCount = values[0]?.length || 0;

        // Detect headers
        const firstRow = values[0];
        const hasHeaders = firstRow.every(cell => typeof cell === "string");

        // Analyze each column
        const columnAnalysis = [];
        for (let col = 0; col < colCount; col++) {
          const columnValues = values.slice(hasHeaders ? 1 : 0).map(row => row[col]);
          const nonEmpty = columnValues.filter(v => v !== null && v !== "");
          const numbers = nonEmpty.filter(v => typeof v === "number") as number[];
          const strings = nonEmpty.filter(v => typeof v === "string");

          let dataType = "mixed";
          if (numbers.length === nonEmpty.length) dataType = "numeric";
          else if (strings.length === nonEmpty.length) dataType = "text";

          const summary: Record<string, unknown> = {
            column: colLetter(col),
            header: hasHeaders ? firstRow[col] : null,
            dataType,
            totalRows: columnValues.length,
            nonEmptyRows: nonEmpty.length,
            emptyRows: columnValues.length - nonEmpty.length
          };

          if (dataType === "numeric" && numbers.length > 0) {
            const sum = numbers.reduce((a, b) => a + b, 0);
            summary.sum = sum;
            summary.average = sum / numbers.length;
            summary.min = Math.min(...numbers);
            summary.max = Math.max(...numbers);
          }

          if (dataType === "text") {
            summary.uniqueValues = new Set(strings).size;
          }

          columnAnalysis.push(summary);
        }

        return {
          range: tool.range,
          dimensions: { rows: rowCount, columns: colCount },
          hasHeaders,
          columnAnalysis
        };
      });
    }

    case "detect_headers": {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const rangeStr = tool.range || ws.getUsedRange().address;
        const range = ws.getRange(rangeStr);
        range.load(["values"]);
        await ctx.sync();

        const firstRow = range.values[0];
        const secondRow = range.values[1];

        const hasHeaders =
          firstRow.every(cell => typeof cell === "string") &&
          firstRow.some(cell => cell !== "") &&
          (secondRow === undefined ||
           firstRow.some((cell, i) => typeof cell !== typeof secondRow[i]));

        return {
          hasHeaders,
          detectedHeaders: hasHeaders ? firstRow : null,
          confidence: hasHeaders ? "high" : "low"
        };
      });
    }

    case "get_data_types": {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const range = ws.getRange(tool.range);
        range.load(["values", "columnCount"]);
        await ctx.sync();

        const columnTypes = [];
        for (let col = 0; col < range.columnCount; col++) {
          const columnValues = range.values.map(row => row[col]).filter(v => v !== null && v !== "");
          const types = columnValues.map(v => typeof v);
          const uniqueTypes = [...new Set(types)];

          let detectedType = "empty";
          if (uniqueTypes.length === 1) {
            detectedType = uniqueTypes[0];
          } else if (uniqueTypes.length > 1) {
            detectedType = "mixed";
          }

          columnTypes.push({
            column: colLetter(col),
            type: detectedType,
            sampleValues: columnValues.slice(0, 3)
          });
        }

        return { range: tool.range, columnTypes };
      });
    }

    case "pivot_data": {
      return Excel.run(async (ctx) => {
        const ws = ctx.workbook.worksheets.getItem(tool.sheet);
        const sourceRange = ws.getRange(tool.sourceRange);
        sourceRange.load(["values"]);
        await ctx.sync();

        const data = sourceRange.values;
        const headers = data[0] as string[];
        const rows = data.slice(1);

        // Simple pivot implementation
        const rowFieldIndices = tool.rowFields.map(f => headers.indexOf(f));
        const valueFieldIndex = headers.indexOf(tool.valueField);
        const aggregation = tool.aggregation || "sum";

        const pivotMap = new Map<string, number[]>();

        for (const row of rows) {
          const key = rowFieldIndices.map(i => row[i]).join("|");
          const value = Number(row[valueFieldIndex]);

          if (!pivotMap.has(key)) {
            pivotMap.set(key, []);
          }
          pivotMap.get(key)!.push(value);
        }

        const pivotResult = Array.from(pivotMap.entries()).map(([key, values]) => {
          let aggregatedValue: number;
          switch (aggregation) {
            case "sum":
              aggregatedValue = values.reduce((a, b) => a + b, 0);
              break;
            case "average":
              aggregatedValue = values.reduce((a, b) => a + b, 0) / values.length;
              break;
            case "count":
              aggregatedValue = values.length;
              break;
            case "min":
              aggregatedValue = Math.min(...values);
              break;
            case "max":
              aggregatedValue = Math.max(...values);
              break;
            default:
              aggregatedValue = values.reduce((a, b) => a + b, 0);
          }

          return [...key.split("|"), aggregatedValue];
        });

        return {
          pivotHeaders: [...tool.rowFields, `${aggregation}(${tool.valueField})`],
          pivotData: pivotResult,
          summary: `Pivoted data by ${tool.rowFields.join(", ")} with ${aggregation} of ${tool.valueField}`
        };
      });
    }
  }
}
