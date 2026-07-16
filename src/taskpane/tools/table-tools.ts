import { toolRegistry } from "./registry";
import { withPerformanceGuard } from "./performance-guard";

toolRegistry.register(
  "listTables",
  "List all tables in the active worksheet.",
  { type: "object", properties: {} },
  async () => {
    return Excel.run(async (context) => {
      const tables = context.workbook.worksheets.getActiveWorksheet().tables;
      tables.load("items/name");
      await context.sync();
      const info = tables.items.map((t) => ({
        name: t.name,
      }));
      return JSON.stringify(info);
    });
  },
);

toolRegistry.register(
  "createTable",
  "Create a table from a range with headers.",
  {
    type: "object",
    properties: {
      address: {
        type: "string",
        description: 'Range address like "A1:D20"',
      },
      hasHeaders: {
        type: "boolean",
        description: "Whether the first row contains headers",
      },
      tableName: {
        type: "string",
        description: "Optional name for the table",
      },
    },
    required: ["address"],
  },
  async (args) => {
    const address = args.address as string;
    const hasHeaders =
      args.hasHeaders !== undefined ? (args.hasHeaders as boolean) : true;
    const tableName = args.tableName as string | undefined;
    return withPerformanceGuard(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      if (tableName) {
        try {
          const existing = sheet.tables.getItem(tableName);
          existing.delete();
          await context.sync();
        } catch {
          // table doesn't exist, nothing to delete
        }
      }

      const range = sheet.getRange(address);
      const table = sheet.tables.add(range, hasHeaders);
      if (tableName) {
        table.name = tableName;
      }
      table.load("name");
      await context.sync();
      return JSON.stringify({ success: true, name: table.name });
    });
  },
);

toolRegistry.register(
  "addTableRow",
  "Add a row to an existing table.",
  {
    type: "object",
    properties: {
      tableName: { type: "string", description: "Name of the table" },
      values: {
        type: "array",
        items: { type: "string" },
        description: "Array of cell values for the new row (one per column)",
      },
    },
    required: ["tableName", "values"],
  },
  async (args) => {
    const tableName = args.tableName as string;
    const values = args.values as unknown[];
    return Excel.run(async (context) => {
      const table = context.workbook.worksheets
        .getActiveWorksheet()
        .tables.getItem(tableName);
      table.rows.add(undefined, [values as (string | number | boolean)[]]);
      await context.sync();
      return JSON.stringify({ success: true, addedTo: tableName });
    });
  },
);

toolRegistry.register(
  "sortTable",
  "Sort a table by a column.",
  {
    type: "object",
    properties: {
      tableName: { type: "string", description: "Name of the table" },
      columnIndex: {
        type: "number",
        description: "Zero-based column index to sort by",
      },
      ascending: {
        type: "boolean",
        description: "Sort ascending (true) or descending (false)",
      },
    },
    required: ["tableName", "columnIndex"],
  },
  async (args) => {
    const tableName = args.tableName as string;
    const columnIndex = args.columnIndex as number;
    const ascending =
      args.ascending !== undefined ? (args.ascending as boolean) : true;
    return Excel.run(async (context) => {
      const table = context.workbook.worksheets
        .getActiveWorksheet()
        .tables.getItem(tableName);
      const sort = table.sort;
      sort.apply([{ key: columnIndex, ascending }], true);
      await context.sync();
      return JSON.stringify({
        success: true,
        sortedBy: columnIndex,
        ascending,
      });
    });
  },
);

toolRegistry.register(
  "filterTable",
  "Apply a filter to a table column.",
  {
    type: "object",
    properties: {
      tableName: { type: "string", description: "Name of the table" },
      columnIndex: {
        type: "number",
        description: "Zero-based column index to filter",
      },
      filterType: {
        type: "string",
        enum: ["values", "custom", "topN", "dynamic"],
        description: "Type of filter to apply",
      },
      values: {
        type: "array",
        items: { type: "string" },
        description: "Values to filter by (for values filter type)",
      },
    },
    required: ["tableName", "columnIndex", "filterType"],
  },
  async (args) => {
    const tableName = args.tableName as string;
    const columnIndex = args.columnIndex as number;
    const filterType = args.filterType as string;
    const values = args.values as string[] | undefined;

    return Excel.run(async (context) => {
      const table = context.workbook.worksheets
        .getActiveWorksheet()
        .tables.getItem(tableName);
      const column = table.columns.getItemAt(columnIndex);
      const filter = column.filter;

      if (filterType === "values" && values) {
        filter.applyValuesFilter(values);
      }

      await context.sync();
      return JSON.stringify({ success: true, filtered: tableName });
    });
  },
);
