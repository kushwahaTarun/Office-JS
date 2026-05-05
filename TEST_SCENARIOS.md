# Excel AI Assistant - Comprehensive Test Scenarios

## 🧪 Test Questions for All Features

### 📋 Setup Test Data First
Use these to create test datasets:

1. **"Create a sales table with headers Name, Region, Sales, Date in A1:D1 and add 10 rows of sample data"**
2. **"Fill A2:A11 with names: Alice, Bob, Charlie, David, Emma, Frank, Grace, Henry, Ivy, Jack"**
3. **"Fill B2:B11 with regions: North, South, East, West (repeat pattern)"**
4. **"Fill C2:C11 with random sales numbers between 1000 and 5000"**
5. **"Fill D2:D11 with dates from 1/1/2024 to 1/10/2024"**

---

## ✅ BASIC OPERATIONS (7 tools)

### Read Operations
6. **"What's the current workbook info?"** → Tests `get_workbook_info`
7. **"Read the data in range A1:D5"** → Tests `read_range`
8. **"What sheets do I have in this workbook?"** → Tests workbook metadata

### Write Operations
9. **"Write 'Total' in cell E1"** → Tests `write_cell`
10. **"Write a header row with Product, Quantity, Price in F1:H1"** → Tests `write_range`
11. **"Add a SUM formula in E2 to sum C2:C11"** → Tests `add_formula`
12. **"Clear the range F1:H10"** → Tests `clear_range`

---

## 📊 TABLES & FORMATTING (3 tools)

### Table Creation
13. **"Convert A1:D11 into an Excel table named SalesData"** → Tests `create_table`
14. **"Create a table from F1:H5 with headers"** → Tests table with hasHeaders
15. **"Make a table from the selected range"** → Tests context awareness

### Formatting
16. **"Format the header row A1:D1 with bold text and blue background"** → Tests `format_range`
17. **"Make column C currency format with green text"** → Tests number formatting
18. **"Format B2:B11 with yellow fill color and font size 14"** → Tests multiple format options
19. **"Make the sales column red and bold"** → Tests smart column detection

### Charts
20. **"Create a column chart from A1:C11"** → Tests `create_chart` basic
21. **"Create a pie chart showing sales by region with title 'Regional Sales'"** → Tests chart with title
22. **"Make a line chart from the sales data and place it at E5"** → Tests chart positioning
23. **"Add a bar chart for this selected data"** → Tests selection context

---

## 🔧 DATA MANIPULATION (8 tools)

### Row/Column Operations
24. **"Insert 3 blank rows at row 5"** → Tests `insert_rows`
25. **"Insert 2 columns before column B"** → Tests `insert_columns`
26. **"Delete rows 8 to 10"** → Tests `delete_rows`
27. **"Delete column F"** → Tests `delete_columns`

### Sorting
28. **"Sort A2:D11 by Sales column in descending order"** → Tests `sort_range`
29. **"Sort this data by Name alphabetically"** → Tests smart column detection + sort
30. **"Sort by Region then by Sales"** → Tests if AI suggests multiple sorts

### Filtering
31. **"Filter the data to show only sales greater than 3000"** → Tests `filter_range`
32. **"Apply a filter to show only North region"** → Tests text criteria
33. **"Filter column C to show values above average"** → Tests analysis + filter

### Auto-Fill & Patterns
34. **"Fill the pattern in A1:A3 down to A20"** → Tests `auto_fill` with pattern
35. **"Auto-fill the dates in column D from D2 to D50"** → Tests date sequence
36. **"Extend the number sequence 2, 4, 6 from E1:E3 to E1:E20"** → Tests numeric pattern
37. **"Copy the formula in E2 down to E11"** → Tests formula auto-fill

### Cell Merging
38. **"Merge cells A1:D1 for the title"** → Tests `merge_cells`
39. **"Unmerge the cells in A1:D1"** → Tests `unmerge_cells`
40. **"Merge A13:B13 and write 'Summary' in it"** → Tests merge + write

---

## 📈 ADVANCED ANALYSIS (5 tools)

### Column Statistics
41. **"What's the summary statistics for column C?"** → Tests `get_column_summary`
42. **"Give me the average, min, and max of the Sales column"** → Tests numeric analysis
43. **"Analyze column B for unique values"** → Tests text analysis
44. **"What are the statistics for C2:C11?"** → Tests range-specific summary

### Data Analysis
45. **"Analyze the data in A1:D11 and tell me what you find"** → Tests `analyze_data`
46. **"What patterns do you see in my data?"** → Tests deep analysis
47. **"Give me insights about this spreadsheet"** → Tests workbook insights
48. **"What data types are in each column of A1:D11?"** → Tests `get_data_types`

### Header Detection
49. **"Does this range have headers?"** → Tests `detect_headers`
50. **"Detect if A1:D11 has a header row"** → Tests header detection explicit
51. **"What are the column headers in this data?"** → Tests header extraction

### Pivot Analysis
52. **"Create a pivot showing total sales by region"** → Tests `pivot_data` sum
53. **"Pivot the data to show average sales per region"** → Tests aggregation=average
54. **"Count how many sales per region"** → Tests aggregation=count
55. **"Show me max sales by region and name"** → Tests multiple row fields

---

## 🎨 CONDITIONAL FORMATTING (1 tool)

56. **"Add color scale formatting to C2:C11"** → Tests `add_conditional_format` ColorScale
57. **"Highlight cells in C2:C11 that are greater than 3000 with red background"** → Tests CellValue rule
58. **"Add data bars to the sales column"** → Tests DataBar
59. **"Add icon set formatting to column C"** → Tests IconSet
60. **"Highlight values below 2000 in yellow"** → Tests conditional with criteria

---

## 🏷️ NAMED RANGES (2 tools)

61. **"Create a named range called 'SalesData' for C2:C11"** → Tests `create_named_range`
62. **"Name the range A1:D11 as 'AllData'"** → Tests named range creation
63. **"What named ranges exist in this workbook?"** → Tests `get_named_ranges`
64. **"Create named ranges for each column: Names, Regions, Sales, Dates"** → Tests multiple named ranges

---

## 🧠 INTELLIGENT BEHAVIOR TESTS

### Context Awareness
65. **"Format the header row"** (without specifying which row) → Tests auto-detection
66. **"Sort by sales"** (without specifying column) → Tests column name recognition
67. **"What's in the selected range?"** → Tests selection context
68. **"Add a total at the bottom"** → Tests understanding "bottom"

### Smart Defaults
69. **"Create a chart"** (no details) → Tests default chart selection
70. **"Format this as a table"** → Tests selection + table creation
71. **"Sort this data"** → Tests default sort (first column, ascending)
72. **"Give me statistics"** → Tests default to current selection

### Pattern Recognition
73. **"Continue this pattern"** (select 1,2,3) → Tests pattern detection
74. **"Fill this down"** (select any pattern) → Tests auto-fill intelligence
75. **"Extend these dates"** → Tests date pattern detection

### Modern Formulas
76. **"Add a formula to lookup values"** → Tests if AI suggests XLOOKUP
77. **"Create a formula to filter data"** → Tests if AI suggests FILTER()
78. **"Make a formula to remove duplicates"** → Tests if AI suggests UNIQUE()
79. **"Add a formula to sort this data"** → Tests if AI suggests SORT()

### Multi-Step Operations
80. **"Create a sales report with headers formatted blue, data sorted by sales, and a chart"** → Tests 3+ steps
81. **"Analyze my data, format the headers, and create a summary table"** → Tests complex workflow
82. **"Clean up this sheet by removing empty rows, formatting headers, and adding totals"** → Tests 4+ steps
83. **"Convert to table, sort by region, add conditional formatting to sales"** → Tests 3 tools

---

## 🚨 EDGE CASES & ERROR HANDLING

### Empty Data
84. **"Analyze an empty range"** → Tests empty data handling
85. **"Create a chart from A100:B105"** (empty cells) → Tests no-data scenario
86. **"What's the average of an empty column?"** → Tests empty column stats

### Invalid Ranges
87. **"Read range ZZZ1:ZZZ100"** → Tests out-of-bounds range
88. **"Sort a single cell A1"** → Tests invalid sort range

### Type Mismatches
89. **"Calculate average of text column B"** → Tests type detection
90. **"Create a pivot with non-numeric values"** → Tests aggregation on text

### Large Operations
91. **"Fill pattern from A1 to A10000"** → Tests large range performance
92. **"Analyze the entire sheet"** → Tests performance limits

---

## 🎯 REAL-WORLD SCENARIOS

### Sales Dashboard
93. **"I have sales data. Format it nicely, add totals, and create a chart"**
94. **"Create a professional sales dashboard with formatting and charts"**
95. **"Make this sales data presentation-ready"**

### Data Cleaning
96. **"Clean up this messy data: format headers, remove duplicates, sort alphabetically"**
97. **"Prepare this data for analysis"**
98. **"Organize this spreadsheet professionally"**

### Analysis & Insights
99. **"Analyze my sales data and give me insights"**
100. **"What trends do you see in this data?"**
101. **"Give me a statistical summary of this dataset"**

### Formatting & Polish
102. **"Make this spreadsheet look professional"**
103. **"Format this data as a nice table with colors"**
104. **"Add conditional formatting to highlight high and low values"**

### Complex Workflows
105. **"Create a complete monthly report: format headers, calculate totals, add averages, create charts, and highlight outliers"**
106. **"Build a dashboard: convert to table, add pivot summary, create 2 charts, format professionally"**
107. **"Analyze this dataset: detect data types, calculate statistics, identify patterns, create visualizations"**

---

## 📝 Testing Checklist

Use this checklist to track your testing:

- [ ] **Basic Operations** (Questions 6-12) - 7 tests
- [ ] **Tables** (Questions 13-15) - 3 tests
- [ ] **Formatting** (Questions 16-19) - 4 tests
- [ ] **Charts** (Questions 20-23) - 4 tests
- [ ] **Row/Column Ops** (Questions 24-27) - 4 tests
- [ ] **Sorting** (Questions 28-30) - 3 tests
- [ ] **Filtering** (Questions 31-33) - 3 tests
- [ ] **Auto-Fill** (Questions 34-37) - 4 tests
- [ ] **Cell Merging** (Questions 38-40) - 3 tests
- [ ] **Column Stats** (Questions 41-44) - 4 tests
- [ ] **Data Analysis** (Questions 45-48) - 4 tests
- [ ] **Header Detection** (Questions 49-51) - 3 tests
- [ ] **Pivot** (Questions 52-55) - 4 tests
- [ ] **Conditional Format** (Questions 56-60) - 5 tests
- [ ] **Named Ranges** (Questions 61-64) - 4 tests
- [ ] **Context Awareness** (Questions 65-68) - 4 tests
- [ ] **Smart Defaults** (Questions 69-72) - 4 tests
- [ ] **Pattern Recognition** (Questions 73-75) - 3 tests
- [ ] **Modern Formulas** (Questions 76-79) - 4 tests
- [ ] **Multi-Step** (Questions 80-83) - 4 tests
- [ ] **Edge Cases** (Questions 84-92) - 9 tests
- [ ] **Real-World** (Questions 93-107) - 15 tests

**Total: 107 Test Questions**

---

## 🎬 Suggested Testing Order

### Phase 1: Basic Functionality (30 min)
1. Setup test data (Questions 1-5)
2. Test basic operations (Questions 6-12)
3. Test tables and basic formatting (Questions 13-19)

### Phase 2: Data Manipulation (30 min)
4. Test charts (Questions 20-23)
5. Test row/column operations (Questions 24-27)
6. Test sorting and filtering (Questions 28-33)
7. Test auto-fill (Questions 34-37)

### Phase 3: Advanced Analysis (30 min)
8. Test statistics and analysis (Questions 41-55)
9. Test conditional formatting (Questions 56-60)
10. Test named ranges (Questions 61-64)

### Phase 4: Intelligence Tests (30 min)
11. Test context awareness (Questions 65-68)
12. Test smart defaults (Questions 69-72)
13. Test pattern recognition (Questions 73-75)
14. Test modern formulas (Questions 76-79)
15. Test multi-step operations (Questions 80-83)

### Phase 5: Edge Cases & Real-World (30 min)
16. Test edge cases (Questions 84-92)
17. Test real-world scenarios (Questions 93-107)

---

## 📊 Expected Results

After testing, you should see:
- ✅ All 27 tools working correctly
- ✅ AI making smart decisions based on context
- ✅ Multi-step operations completing successfully
- ✅ Modern Excel formulas being recommended
- ✅ Proper error handling for edge cases
- ✅ Professional formatting and presentation
- ✅ Accurate data analysis and insights

**Happy Testing! 🚀**
