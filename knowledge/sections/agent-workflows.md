# Agent Workflows

## Workflow: Create a Report from Scratch

### Steps:
1. **getWorkbookOverview()** — check current state
2. Ask user confirmation if clearing existing data
3. **clearWorksheet()** or **clearRange()** — only after confirmation
4. **setValues()** for headers row
5. **setValues()** for data rows
6. **createTable()** with hasHeaders: true
7. **applyFormat()** for numbers
8. **mergeCells()** for title if needed
9. **setCellFormat()** for bold headers

### Example:
User: "Create a sales report with columns Product, Qty, Price, Total"
```
→ getWorkbookOverview()
→ setValues("A1", [["Sales Report","","",""]])  // title row
→ setValues("A3", [["Product","Qty","Price","Total"]]) // headers
→ setValues("A4", [["Widget A",10,25.99,0], ["Widget B",5,15.50,0]])
→ createTable("A3:D6", true, "SalesTable")
→ fillFormula("D4", "D5:D6", "B4*C4")
→ applyFormat("C4:D6", "#,##0.00")
```

## Workflow: Add Calculations to Existing Table

### Steps:
1. **getWorkbookOverview()** — find table and columns
2. **getRange()** — read data to understand columns
3. **fillFormula()** or **setFormula()** — add formula column
4. **verifyRange()** — read back to confirm

### Example:
User: "Add a column that calculates profit margin"
```
→ getWorkbookOverview() → see table with Revenue and Cost columns
→ getRange("SalesTable[Revenue]")
→ getRange("SalesTable[Cost]")
→ Need to add column after Cost
→ First setValues for the new header in the next column
→ fillFormula for profit margin: "=(Revenue-Cost)/Revenue*100"
→ applyFormat with "0.00%"
```

## Workflow: Analyze and Summarize Data

### Steps:
1. **getWorkbookOverview()** — identify data location
2. **getRange()** — read the full data range
3. **getKnowledge("excel-formulas")** if needed for functions
4. Determine analysis type: sum, average, count, pivot
5. Execute appropriate tools

### Example:
User: "Show me total sales by region"
```
→ getWorkbookOverview() → see "Sales" table with Region and Amount
→ getRange("SalesTable[#All]")
→ Use SUMIF or create a pivot table
→ For simple: setFormula for each region total
→ For complex: addPivotTable with Region as row, Amount as value (sum)
```

## Workflow: Format and Beautify a Table

### Steps:
1. **getWorkbookOverview()** — see current structure
2. **getRange()** — read data to understand columns/data types
3. Identify which columns are numeric, which are text, which are dates
4. **applyFormat()** per column for appropriate format
5. **createTable()** if not already a table (applies banded rows)
6. Set bold headers, colors, alignment

### Example:
User: "Make this spreadsheet look professional"
```
→ getWorkbookOverview() → see data in Sheet1
→ getRange("Sheet1!A1:E20") → read all data
→ Identify: A=text, B=numbers, C=currency, D=dates, E=percentage
→ applyFormat("B:B", "#,##0")
→ applyFormat("C:C", "$#,##0.00")
→ applyFormat("D:D", "DD.MM.YYYY")
→ applyFormat("E:E", "0%")
→ createTable with banded rows
```

## Workflow: Creating Charts

### Steps:
1. **getWorkbookOverview()** — find data
2. **getRange()** — read data including headers
3. **addChart(type, sourceAddress)** — create chart
4. **configureChart()** — title, labels, legend
5. Report chart name to user

### Chart Types:
- `columnClustered` — side-by-side comparison
- `columnStacked` — part-to-whole
- `line` — trends over time
- `pie` — proportions
- `barClustered` — horizontal comparison
- `scatter` — correlation/relationship
- `area` — volume over time

### Example:
User: "Create a chart of sales by month"
```
→ getRange("A1:B13") → months and sales data
→ addChart("line", "Sheet1!A1:B13", "SalesTrend")
→ configureChart("SalesTrend", "Monthly Sales 2026", "Sales")
```

## Workflow: Fix Errors in Existing Data

### Steps:
1. **getWorkbookOverview()** — see what exists
2. **getRange()** — read data and formulas
3. Analyze the errors (see Error Handling section)
4. Fix with appropriate tools
5. Verify the fix

### Example:
User: "Fix the calculation errors"
```
→ getRange("Sheet1!A1:D20") → read data and formulas
→ See #NAME? in formula =B8B6
→ setFormula("D3", "B8*B6") → fix the missing operator
→ fillFormula("D3", "D4:D20", "B8*B6") → fix all rows
→ getRange("D1:D20") → verify fix
```

## Workflow: Multi-Sheet Operations

### Steps:
1. **getWorkbookOverview()** — identify all sheets and their data
2. Process each sheet
3. For cross-sheet references, use SheetName!CellRef syntax

### Example:
User: "Combine all monthly sheets into annual summary"
```
→ getWorkbookOverview() → see Jan, Feb, Mar sheets
→ Create summary sheet
→ Use formulas referencing other sheets: =SUM(Jan!D3:D10, Feb!D3:D10, Mar!D3:D10)
```
