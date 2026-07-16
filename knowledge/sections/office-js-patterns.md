# Office.js Patterns for Excel AI Agent

## Excel.run Pattern
All Office.js operations must be wrapped in `Excel.run()`:
```typescript
await Excel.run(async (context) => {
  const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  range.load("values");
  await context.sync();
  // use range.values
});
```

## Load + Sync Pattern
Every property you read must be loaded before syncing:
```typescript
range.load("address, values, formulas, rowCount, columnCount");
await context.sync();
```

## Getting Ranges
- `getRange("A1:C10")` — by address string
- `getRangeByIndexes(row, col, rows, cols)` — by position (0-based)
- `getCell(row, col)` — single cell relative to parent range
- `getResizedRange(deltaRows, deltaCols)` — expand/shrink range
- `getAbsoluteResizedRange(rows, cols)` — absolute sizing

## Writing Values
```typescript
range.values = [["Header1", "Header2"], [1, 2], [3, 4]];
await context.sync();
```

## Writing Formulas
```typescript
range.formulas = [["=SUM(A1:A10)", "=AVERAGE(B1:B10)"]];
await context.sync();
```
Note: formulas must start with `=` in the array.

## AutoFill
Used to fill formulas down a column:
```typescript
source.autoFill(destination, Excel.AutoFillType.fillDefault);
await context.sync();
```
- `fillDefault` — copies values, formats, formulas (adjusts relative refs)
- `fillCopy` — copies exact values
- `fillFormats` — copies only formatting

## Creating Tables
```typescript
const table = sheet.tables.add(range, true); // true = has headers
table.name = "MyTable";
await context.sync();
```

## Working with Tables
```typescript
const table = sheet.tables.getItem("MyTable");
table.rows.add(undefined, [values]); // add at end
table.columns.getItemAt(0); // get first column
table.sort.apply([{ key: 0, ascending: true }], true);
column.filter.applyValuesFilter(["Value1", "Value2"]);
```

## Number Formats
```typescript
range.numberFormat = [["#,##0.00"]]; // Applies to all cells in range
```
Common formats:
- `#,##0.00` — thousands separator, 2 decimals
- `$#,##0.00` — currency USD
- `0%` — percentage (0.5 → 50%)
- `0.00%` — percentage with decimals
- `#,##0` — whole numbers with separator
- `DD.MM.YYYY` — date format
- `@` — text format

## Merging Cells
```typescript
const range = sheet.getRange("A1:E1");
range.merge(true); // true = across columns in the row
```
Note: merged cells cannot be inside a table. Merge FIRST, then create table below.

## Cell Formatting
```typescript
const range = sheet.getRange("A1:E1");
const format = range.format;
format.font.bold = true;
format.font.color = "white";
format.fill.color = "#4472C4";
format.horizontalAlignment = "Center";
```

## Pivot Tables
```typescript
const pivot = sheet.pivotTables.add("PivotName", sourceRange, targetRange);
pivot.pivotHierarchies.add(pivot, "RowField"); // row
pivot.pivotHierarchies.add(pivot, "ColField"); // column
pivot.pivotHierarchies.add(pivot, "ValueField"); // value
// For data field:
const dataHierarchy = pivot.pivotHierarchies.getItem("FieldName");
dataHierarchy.pivotFields.getItem("FieldName").summarizeBy = Excel.AggregationFunction.sum;
```

## Charts
```typescript
const chart = sheet.charts.add(Excel.ChartType.columnClustered, sourceRange, Excel.ChartSeriesBy.auto);
chart.setTitle("Chart Title");
chart.axes.getItem("Primary").title.text = "Axis Title";
chart.legend.position = Excel.ChartLegendPosition.bottom;
```

## Performance Tips
- Minimize `context.sync()` calls — batch loads where possible
- Use `getUsedRangeOrNullObject()` instead of `getUsedRange()` for empty sheets
- Load only the properties you need, not everything
- For large data reads, use `getRange()` with specific address, not whole sheet

## Common Pitfalls
1. **PropertyNotLoaded** — accessing a property before load() + sync()
2. **NullObject** — calling getItem() on non-existent item; use getItemOrNullObject()
3. **Range size mismatch** — setting values array that doesn't match range dimensions
4. **Table overlap** — creating a table that overlaps existing tables
5. **AutoFill source** — source cell must be within the destination range
