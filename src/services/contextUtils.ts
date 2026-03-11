/**
 * Context utilities for enhanced Excel AI capabilities
 * Provides pattern detection, data type inference, and smart suggestions
 */

export interface DataPattern {
  type: "numeric_sequence" | "date_sequence" | "text_pattern" | "formula_pattern" | "none";
  confidence: number;
  suggestion?: string;
}

export interface ColumnInsight {
  column: string;
  dataType: "numeric" | "text" | "date" | "boolean" | "mixed" | "empty";
  hasHeader: boolean;
  headerName?: string;
  stats?: {
    count: number;
    unique?: number;
    min?: number;
    max?: number;
    average?: number;
  };
  suggestions: string[];
}

/**
 * Detect patterns in a range of values for smart auto-fill suggestions
 */
export function detectPattern(values: unknown[][]): DataPattern {
  if (values.length === 0 || values[0].length === 0) {
    return { type: "none", confidence: 0 };
  }

  const flatValues = values.flat().filter(v => v !== null && v !== "");

  if (flatValues.length < 2) {
    return { type: "none", confidence: 0 };
  }

  // Check for numeric sequence
  if (flatValues.every(v => typeof v === "number")) {
    const numbers = flatValues as number[];
    const diffs = numbers.slice(1).map((n, i) => n - numbers[i]);
    const avgDiff = diffs.reduce((a, b) => a + b, 0) / diffs.length;
    const isSequence = diffs.every(d => Math.abs(d - avgDiff) < 0.01);

    if (isSequence) {
      return {
        type: "numeric_sequence",
        confidence: 0.95,
        suggestion: `Detected sequence with step ${avgDiff}. Use auto_fill to continue.`
      };
    }
  }

  // Check for date sequence
  const datePattern = /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/;
  if (flatValues.every(v => typeof v === "string" && datePattern.test(v as string))) {
    return {
      type: "date_sequence",
      confidence: 0.9,
      suggestion: "Detected date pattern. Use auto_fill for date series."
    };
  }

  // Check for text pattern (e.g., "Item 1", "Item 2")
  if (flatValues.every(v => typeof v === "string")) {
    const strings = flatValues as string[];
    const numbersInStrings = strings.map(s => s.match(/\d+/)?.[0]);

    if (numbersInStrings.every(n => n !== undefined)) {
      return {
        type: "text_pattern",
        confidence: 0.85,
        suggestion: "Detected numbered text pattern. Use auto_fill to continue series."
      };
    }
  }

  // Check for formulas
  if (flatValues.some(v => typeof v === "string" && (v as string).startsWith("="))) {
    return {
      type: "formula_pattern",
      confidence: 0.95,
      suggestion: "Detected formulas. Use auto_fill to copy formula pattern."
    };
  }

  return { type: "none", confidence: 0 };
}

/**
 * Analyze column and provide intelligent insights
 */
export function analyzeColumn(
  values: unknown[],
  headerValue?: unknown
): ColumnInsight {
  const nonEmpty = values.filter(v => v !== null && v !== "");

  if (nonEmpty.length === 0) {
    return {
      column: "",
      dataType: "empty",
      hasHeader: false,
      suggestions: ["Column is empty. Consider adding data or removing."]
    };
  }

  const types = nonEmpty.map(v => typeof v);
  const uniqueTypes = [...new Set(types)];

  let dataType: ColumnInsight["dataType"] = "mixed";
  if (uniqueTypes.length === 1) {
    if (uniqueTypes[0] === "number") dataType = "numeric";
    else if (uniqueTypes[0] === "string") dataType = "text";
    else if (uniqueTypes[0] === "boolean") dataType = "boolean";
  }

  const suggestions: string[] = [];
  const stats: ColumnInsight["stats"] = { count: nonEmpty.length };

  // Numeric analysis
  if (dataType === "numeric") {
    const numbers = nonEmpty as number[];
    stats.min = Math.min(...numbers);
    stats.max = Math.max(...numbers);
    stats.average = numbers.reduce((a, b) => a + b, 0) / numbers.length;

    suggestions.push(`Use get_column_summary for detailed statistics`);

    if (stats.min >= 0 && stats.max <= 1) {
      suggestions.push("Values appear to be percentages. Consider formatting as percentage.");
    }

    if (numbers.every(n => Number.isInteger(n))) {
      suggestions.push("All integers. Consider using integer formatting.");
    }
  }

  // Text analysis
  if (dataType === "text") {
    const strings = nonEmpty as string[];
    stats.unique = new Set(strings).size;

    if (stats.unique < nonEmpty.length * 0.5) {
      suggestions.push(`Only ${stats.unique} unique values. Consider creating a filter or pivot.`);
    }

    if (strings.every(s => s.length < 20 && /^[A-Z]/.test(s))) {
      suggestions.push("Appears to be categorical data. Good for grouping/pivoting.");
    }
  }

  // Date detection
  const datePattern = /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/;
  if (dataType === "text" && (nonEmpty as string[]).every(v => datePattern.test(v))) {
    dataType = "date";
    suggestions.push("Dates detected. Consider using date formulas or formatting.");
  }

  return {
    column: "",
    dataType,
    hasHeader: typeof headerValue === "string" && headerValue !== "",
    headerName: typeof headerValue === "string" ? headerValue : undefined,
    stats,
    suggestions
  };
}

/**
 * Suggest appropriate chart type based on data structure
 */
export function suggestChartType(
  data: unknown[][],
  hasHeaders: boolean
): { chartType: string; confidence: number; reason: string } {
  const dataRows = hasHeaders ? data.slice(1) : data;
  const columnCount = data[0]?.length || 0;

  if (columnCount === 0 || dataRows.length === 0) {
    return { chartType: "ColumnClustered", confidence: 0, reason: "No data" };
  }

  // Single column of numbers → Bar chart
  if (columnCount === 1) {
    return {
      chartType: "Bar",
      confidence: 0.8,
      reason: "Single numeric column works well with bar chart"
    };
  }

  // Two columns: categories + values → Column chart
  if (columnCount === 2) {
    const col1Types = dataRows.map(r => typeof r[0]);
    const col2Types = dataRows.map(r => typeof r[1]);

    if (col1Types.every(t => t === "string") && col2Types.every(t => t === "number")) {
      return {
        chartType: "ColumnClustered",
        confidence: 0.95,
        reason: "Category-value pairs ideal for column chart"
      };
    }
  }

  // Multiple numeric columns → Line chart (time series assumption)
  if (columnCount >= 3) {
    const allNumeric = dataRows.every(row =>
      row.slice(1).every(cell => typeof cell === "number")
    );

    if (allNumeric) {
      return {
        chartType: "Line",
        confidence: 0.85,
        reason: "Multiple numeric columns suggest time series or trends"
      };
    }
  }

  // Few rows, many columns → Pie chart
  if (dataRows.length <= 7 && columnCount === 2) {
    return {
      chartType: "Pie",
      confidence: 0.7,
      reason: "Few categories work well with pie chart"
    };
  }

  return {
    chartType: "ColumnClustered",
    confidence: 0.6,
    reason: "Default choice for general data"
  };
}

/**
 * Suggest appropriate number format based on values
 */
export function suggestNumberFormat(values: number[]): string {
  if (values.length === 0) return "General";

  const allIntegers = values.every(n => Number.isInteger(n));
  const allPercentages = values.every(n => n >= 0 && n <= 1);
  const allCurrency = values.some(n => n > 100 && Number.isInteger(n * 100));
  const allSmall = values.every(n => Math.abs(n) < 0.01 && n !== 0);

  if (allPercentages) return "0.00%";
  if (allCurrency) return "$#,##0.00";
  if (allSmall) return "0.00E+00"; // Scientific notation
  if (allIntegers) return "#,##0";

  return "0.00";
}

/**
 * Extract insights from workbook context for better AI decisions
 */
export interface WorkbookInsights {
  primarySheet: string;
  largestSheet: string;
  totalCells: number;
  hasMultipleSheets: boolean;
  recommendations: string[];
}

export function getWorkbookInsights(
  sheetsMetadata: Array<{ name: string; rowCount: number; columnCount: number; hasData: boolean }>
): WorkbookInsights {
  const sheetsWithData = sheetsMetadata.filter(s => s.hasData);

  if (sheetsWithData.length === 0) {
    return {
      primarySheet: sheetsMetadata[0]?.name || "Sheet1",
      largestSheet: sheetsMetadata[0]?.name || "Sheet1",
      totalCells: 0,
      hasMultipleSheets: sheetsMetadata.length > 1,
      recommendations: ["Workbook is empty. Start by adding data to a sheet."]
    };
  }

  const largestSheet = sheetsWithData.reduce((max, sheet) =>
    (sheet.rowCount * sheet.columnCount) > (max.rowCount * max.columnCount) ? sheet : max
  );

  const totalCells = sheetsWithData.reduce(
    (sum, s) => sum + s.rowCount * s.columnCount,
    0
  );

  const recommendations: string[] = [];

  if (sheetsMetadata.length > 3) {
    recommendations.push(`You have ${sheetsMetadata.length} sheets. Consider consolidating if possible.`);
  }

  if (largestSheet.rowCount > 1000) {
    recommendations.push(`${largestSheet.name} has ${largestSheet.rowCount} rows. Consider using tables and filters for better performance.`);
  }

  return {
    primarySheet: largestSheet.name,
    largestSheet: largestSheet.name,
    totalCells,
    hasMultipleSheets: sheetsMetadata.length > 1,
    recommendations
  };
}
