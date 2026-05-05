# Complete Tool Reference Guide

This document shows all 27 tools, their parameters, and exact JSON format for your AI agent.

## 📍 Where Tools Are Defined

- **Type Definitions:** [office.ts:21-48](src/services/office.ts#L21-L48)
- **Implementations:** [office.ts:160-648](src/services/office.ts#L160-L648)

---

## 🔧 All 27 Tools

### 1. BASIC OPERATIONS (7 tools)

#### `get_workbook_info`
Get metadata about all sheets in the workbook.

**Parameters:** None

**JSON Format:**
```json
{"tool":"get_workbook_info"}
```

**Returns:**
```json
{
  "activeSheet": "Sheet1",
  "sheets": ["Sheet1", "Sheet2", "Sheet3"]
}
```

---

#### `read_range`
Read cell values and formulas from a range.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range address (e.g., "A1:C10")

**JSON Format:**
```json
{"tool":"read_range","sheet":"Sheet1","range":"A1:C10"}
```

**Returns:**
```json
{
  "address": "Sheet1!A1:C10",
  "values": [["Name","Age","City"],["Alice",30,"NYC"]],
  "formulas": [["Name","Age","City"],["Alice","30","NYC"]],
  "dimensions": {"rows":2,"cols":3}
}
```

**Implementation:** [office.ts:172-185](src/services/office.ts#L172-L185)

---

#### `write_cell`
Write a single value to a cell.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `cell` (string, required) - Cell address (e.g., "B3")
- `value` (string|number|boolean, required) - Value to write

**JSON Format:**
```json
{"tool":"write_cell","sheet":"Sheet1","cell":"B3","value":42}
```

**Returns:**
```
"Wrote \"42\" to Sheet1!B3"
```

---

#### `write_range`
Write a 2D array of values to a range.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range address
- `values` (array[][], required) - 2D array of values
- `isFormula` (boolean, optional) - If true, values are treated as formulas

**JSON Format:**
```json
{"tool":"write_range","sheet":"Sheet1","range":"A1:C2","values":[["Name","Age","City"],["Alice",30,"NYC"]],"isFormula":false}
```

**With Formulas:**
```json
{"tool":"write_range","sheet":"Sheet1","range":"D1:D2","values":[["Total"],["=SUM(A1:C1)"]],"isFormula":true}
```

**Returns:**
```
"Wrote 2 row(s) to Sheet1!A1:C2"
```

**Implementation:** [office.ts:199-212](src/services/office.ts#L199-L212)

---

#### `add_formula`
Add an Excel formula to a cell.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `cell` (string, required) - Cell address
- `formula` (string, required) - Formula (must start with =)

**JSON Format:**
```json
{"tool":"add_formula","sheet":"Sheet1","cell":"D1","formula":"=SUM(A1:C1)"}
```

**Returns:**
```
"Formula =SUM(A1:C1) added to Sheet1!D1"
```

---

#### `clear_range`
Clear content and formatting from a range.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range address

**JSON Format:**
```json
{"tool":"clear_range","sheet":"Sheet1","range":"A1:C10"}
```

**Returns:**
```
"Cleared range Sheet1!A1:C10"
```

---

### 2. TABLES & FORMATTING (3 tools)

#### `create_table`
Convert a range to an Excel table.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range address
- `tableName` (string, optional) - Custom table name
- `hasHeaders` (boolean, optional, default: true) - First row is headers

**JSON Format:**
```json
{"tool":"create_table","sheet":"Sheet1","range":"A1:D10","tableName":"SalesData","hasHeaders":true}
```

**Returns:**
```
"Created table SalesData from A1:D10 on Sheet1"
```

**Implementation:** [office.ts:216-227](src/services/office.ts#L216-L227)

---

#### `format_range`
Apply formatting to a range (colors, fonts, number formats).

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range address
- `numberFormat` (string, optional) - e.g., "0.00", "$#,##0.00", "0.00%"
- `fontColor` (string, optional) - Hex color (e.g., "#FFFFFF")
- `fillColor` (string, optional) - Background color (e.g., "#4472C4")
- `bold` (boolean, optional) - Make text bold
- `italic` (boolean, optional) - Make text italic
- `fontSize` (number, optional) - Font size

**JSON Format:**
```json
{"tool":"format_range","sheet":"Sheet1","range":"A1:D1","bold":true,"fillColor":"#4472C4","fontColor":"#FFFFFF","fontSize":12}
```

**Number Formatting:**
```json
{"tool":"format_range","sheet":"Sheet1","range":"C2:C10","numberFormat":"$#,##0.00"}
```

**Returns:**
```
"Formatted range Sheet1!A1:D1"
```

**Implementation:** [office.ts:230-257](src/services/office.ts#L230-L257)

---

#### `create_chart`
Create a chart from data range.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `dataRange` (string, required) - Data range for chart
- `chartType` (string, required) - One of: "ColumnClustered", "Line", "Pie", "Bar", "Area", "XYScatter"
- `title` (string, optional) - Chart title
- `position` (string, optional) - Cell address for chart position

**JSON Format:**
```json
{"tool":"create_chart","sheet":"Sheet1","dataRange":"A1:C10","chartType":"ColumnClustered","title":"Sales Report","position":"E2"}
```

**Chart Types:**
- `ColumnClustered` - Column chart
- `Line` - Line chart
- `Pie` - Pie chart
- `Bar` - Bar chart
- `Area` - Area chart
- `XYScatter` - Scatter plot

**Returns:**
```
"Created ColumnClustered chart from A1:C10 on Sheet1"
```

**Implementation:** [office.ts:260-276](src/services/office.ts#L260-L276)

---

### 3. DATA MANIPULATION (8 tools)

#### `insert_rows`
Insert blank rows at a specific index.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `index` (number, required) - Row index (0-based)
- `count` (number, required) - Number of rows to insert

**JSON Format:**
```json
{"tool":"insert_rows","sheet":"Sheet1","index":5,"count":3}
```

**Returns:**
```
"Inserted 3 row(s) at index 5 on Sheet1"
```

---

#### `insert_columns`
Insert blank columns at a specific index.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `index` (number, required) - Column index (0-based, 0=A, 1=B, etc.)
- `count` (number, required) - Number of columns to insert

**JSON Format:**
```json
{"tool":"insert_columns","sheet":"Sheet1","index":2,"count":1}
```

---

#### `delete_rows`
Delete rows at a specific index.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `index` (number, required) - Starting row index (0-based)
- `count` (number, required) - Number of rows to delete

**JSON Format:**
```json
{"tool":"delete_rows","sheet":"Sheet1","index":5,"count":2}
```

---

#### `delete_columns`
Delete columns at a specific index.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `index` (number, required) - Starting column index (0-based)
- `count` (number, required) - Number of columns to delete

**JSON Format:**
```json
{"tool":"delete_columns","sheet":"Sheet1","index":3,"count":1}
```

---

#### `sort_range`
Sort a range by a specific column.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range to sort
- `sortColumn` (number, required) - Column index to sort by (0-based)
- `ascending` (boolean, optional, default: true) - Sort order

**JSON Format:**
```json
{"tool":"sort_range","sheet":"Sheet1","range":"A2:D10","sortColumn":2,"ascending":false}
```

**Returns:**
```
"Sorted A2:D10 by column 2 (descending)"
```

**Implementation:** [office.ts:319-329](src/services/office.ts#L319-L329)

---

#### `filter_range`
Apply autofilter to a range.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range to filter
- `filterColumn` (number, required) - Column index to filter (0-based)
- `criteria` (string, required) - Filter criteria (e.g., ">=100", "North")

**JSON Format:**
```json
{"tool":"filter_range","sheet":"Sheet1","range":"A1:D10","filterColumn":2,"criteria":">=3000"}
```

**Returns:**
```
"Applied filter to column 2 in A1:D10 with criteria: >=3000"
```

---

#### `auto_fill`
Auto-fill a pattern from source to destination range.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `sourceRange` (string, required) - Source range with pattern
- `destinationRange` (string, required) - Destination range to fill

**JSON Format:**
```json
{"tool":"auto_fill","sheet":"Sheet1","sourceRange":"A1:A3","destinationRange":"A1:A10"}
```

**Returns:**
```
"Auto-filled from A1:A3 to A1:A10"
```

---

#### `merge_cells`
Merge cells in a range.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range to merge

**JSON Format:**
```json
{"tool":"merge_cells","sheet":"Sheet1","range":"A1:D1"}
```

---

#### `unmerge_cells`
Unmerge cells in a range.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range to unmerge

**JSON Format:**
```json
{"tool":"unmerge_cells","sheet":"Sheet1","range":"A1:D1"}
```

---

### 4. ADVANCED ANALYSIS (5 tools)

#### `get_column_summary`
Get statistical summary of a column.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `column` (string, required) - Column letter (e.g., "B")
- `range` (string, optional) - Specific range (e.g., "B2:B100")

**JSON Format:**
```json
{"tool":"get_column_summary","sheet":"Sheet1","column":"C","range":"C2:C50"}
```

**Returns (for numeric data):**
```json
{
  "column": "C",
  "count": 48,
  "sum": 156000,
  "average": 3250,
  "min": 1000,
  "max": 5000,
  "median": 3200,
  "type": "numeric"
}
```

**Returns (for text data):**
```json
{
  "column": "B",
  "count": 48,
  "type": "text",
  "uniqueValues": 12
}
```

**Implementation:** [office.ts:347-384](src/services/office.ts#L347-L384)

---

#### `analyze_data`
Deep analysis of data range (headers, types, column stats).

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range to analyze

**JSON Format:**
```json
{"tool":"analyze_data","sheet":"Sheet1","range":"A1:D50"}
```

**Returns:**
```json
{
  "range": "A1:D50",
  "dimensions": {"rows":50,"columns":4},
  "hasHeaders": true,
  "columnAnalysis": [
    {
      "column": "A",
      "header": "Name",
      "dataType": "text",
      "totalRows": 49,
      "nonEmptyRows": 49,
      "emptyRows": 0,
      "uniqueValues": 45
    },
    {
      "column": "C",
      "header": "Sales",
      "dataType": "numeric",
      "sum": 156000,
      "average": 3183.67,
      "min": 1000,
      "max": 5000
    }
  ]
}
```

**Implementation:** [office.ts:471-528](src/services/office.ts#L471-L528)

---

#### `detect_headers`
Detect if first row contains headers.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, optional) - Range to check (defaults to used range)

**JSON Format:**
```json
{"tool":"detect_headers","sheet":"Sheet1","range":"A1:D10"}
```

**Returns:**
```json
{
  "hasHeaders": true,
  "detectedHeaders": ["Name","Region","Sales","Date"],
  "confidence": "high"
}
```

---

#### `get_data_types`
Analyze data types for each column.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range to analyze

**JSON Format:**
```json
{"tool":"get_data_types","sheet":"Sheet1","range":"A1:D10"}
```

**Returns:**
```json
{
  "range": "A1:D10",
  "columnTypes": [
    {"column":"A","type":"string","sampleValues":["Alice","Bob","Charlie"]},
    {"column":"B","type":"string","sampleValues":["North","South","East"]},
    {"column":"C","type":"number","sampleValues":[1500,2300,4200]},
    {"column":"D","type":"string","sampleValues":["1/1/2024","1/2/2024"]}
  ]
}
```

---

#### `pivot_data`
Create a pivot table summary.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `sourceRange` (string, required) - Source data range
- `rowFields` (string[], required) - Fields to group by
- `valueField` (string, required) - Field to aggregate
- `columnFields` (string[], optional) - Column fields
- `aggregation` (string, optional, default: "sum") - "sum", "average", "count", "min", "max"

**JSON Format:**
```json
{"tool":"pivot_data","sheet":"Sheet1","sourceRange":"A1:D10","rowFields":["Region"],"valueField":"Sales","aggregation":"sum"}
```

**Returns:**
```json
{
  "pivotHeaders": ["Region","sum(Sales)"],
  "pivotData": [
    ["North",15000],
    ["South",12000],
    ["East",18000],
    ["West",11000]
  ],
  "summary": "Pivoted data by Region with sum of Sales"
}
```

**Implementation:** [office.ts:587-645](src/services/office.ts#L587-L645)

---

### 5. CONDITIONAL FORMATTING (1 tool)

#### `add_conditional_format`
Add conditional formatting rules.

**Parameters:**
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range to format
- `ruleType` (string, required) - "CellValue", "ColorScale", "DataBar", "IconSet"
- `operator` (string, optional) - For CellValue: "GreaterThan", "LessThan", etc.
- `formula` (string, optional) - For CellValue: comparison value
- `fillColor` (string, optional) - Background color for CellValue

**JSON Format (Color Scale):**
```json
{"tool":"add_conditional_format","sheet":"Sheet1","range":"C2:C10","ruleType":"ColorScale"}
```

**JSON Format (Highlight > Value):**
```json
{"tool":"add_conditional_format","sheet":"Sheet1","range":"C2:C10","ruleType":"CellValue","operator":"GreaterThan","formula":"3000","fillColor":"#FFC7CE"}
```

**JSON Format (Data Bars):**
```json
{"tool":"add_conditional_format","sheet":"Sheet1","range":"C2:C10","ruleType":"DataBar"}
```

**Implementation:** [office.ts:420-449](src/services/office.ts#L420-L449)

---

### 6. NAMED RANGES (2 tools)

#### `create_named_range`
Create a named range.

**Parameters:**
- `name` (string, required) - Name for the range
- `sheet` (string, required) - Sheet name
- `range` (string, required) - Range address

**JSON Format:**
```json
{"tool":"create_named_range","name":"SalesData","sheet":"Sheet1","range":"C2:C10"}
```

---

#### `get_named_ranges`
List all named ranges in workbook.

**Parameters:** None

**JSON Format:**
```json
{"tool":"get_named_ranges"}
```

**Returns:**
```json
[
  {"name":"SalesData","formula":"=Sheet1!$C$2:$C$10"},
  {"name":"Regions","formula":"=Sheet1!$B$2:$B$10"}
]
```

---

### 7. FINAL ANSWER (1 tool)

#### `final_answer`
Signal task completion with summary for user.

**Parameters:**
- `answer` (string, required) - Summary message for user

**JSON Format:**
```json
{"tool":"final_answer","answer":"I created a sales table with headers Name, Region, Sales, Date in cells A1:D1. The table is now ready for data entry."}
```

**This tool:**
- Ends the agent loop
- Returns the answer to the user
- Should be called when task is complete

---

## 🎯 Usage Examples

### Example 1: Create Table with Headers
```json
{"tool":"write_range","sheet":"Sheet1","range":"A1:D1","values":[["Name","Region","Sales","Date"]],"isFormula":false}
```

### Example 2: Format Headers
```json
{"tool":"format_range","sheet":"Sheet1","range":"A1:D1","bold":true,"fillColor":"#4472C4","fontColor":"#FFFFFF","fontSize":12}
```

### Example 3: Add Data
```json
{"tool":"write_range","sheet":"Sheet1","range":"A2:D3","values":[["Alice","North",3000,"1/1/2024"],["Bob","South",2500,"1/2/2024"]],"isFormula":false}
```

### Example 4: Create Chart
```json
{"tool":"create_chart","sheet":"Sheet1","dataRange":"A1:C10","chartType":"ColumnClustered","title":"Sales by Region"}
```

### Example 5: Complete Task
```json
{"tool":"final_answer","answer":"I've created a sales table with 4 columns (Name, Region, Sales, Date), formatted the headers with blue background, and added 2 sample rows of data."}
```

---

## 🔄 Multi-Step Workflow Example

For: "Create a sales table with headers Name, Region, Sales, Date in A1:D1"

**Step 1:** Write headers
```json
{"tool":"write_range","sheet":"Sheet1","range":"A1:D1","values":[["Name","Region","Sales","Date"]],"isFormula":false}
```

**Step 2:** Format headers (optional but nice)
```json
{"tool":"format_range","sheet":"Sheet1","range":"A1:D1","bold":true,"fillColor":"#4472C4","fontColor":"#FFFFFF"}
```

**Step 3:** Complete task
```json
{"tool":"final_answer","answer":"Created a sales table with headers Name, Region, Sales, Date in A1:D1. Headers are formatted with blue background and white bold text."}
```

---

## 📋 Quick Reference Table

| Tool | Use Case | Required Params |
|------|----------|-----------------|
| `write_range` | Write data | sheet, range, values |
| `read_range` | Read data | sheet, range |
| `format_range` | Apply formatting | sheet, range + format options |
| `create_table` | Make Excel table | sheet, range |
| `create_chart` | Add chart | sheet, dataRange, chartType |
| `sort_range` | Sort data | sheet, range, sortColumn |
| `analyze_data` | Data insights | sheet, range |
| `pivot_data` | Pivot summary | sheet, sourceRange, rowFields, valueField |
| `final_answer` | Complete task | answer |

---

## 📚 Full Documentation

- **Complete feature list:** [FEATURES.md](FEATURES.md)
- **Test scenarios:** [TEST_SCENARIOS.md](TEST_SCENARIOS.md)
- **Troubleshooting:** [TROUBLESHOOTING.md](TROUBLESHOOTING.md)

---

**Your AI agent should return these exact JSON formats to execute Excel operations!**
