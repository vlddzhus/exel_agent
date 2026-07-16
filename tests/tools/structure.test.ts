import { toolRegistry } from "../../src/taskpane/tools/registry";

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

import "../../src/taskpane/tools/structure";

function setupExcelMock() {
  const syncMock = jest.fn().mockResolvedValue(undefined);
  const mockRange = {
    address: "Sheet1!A1:E10",
    load: jest.fn(),
    rowCount: 10,
    columnCount: 5,
  };
  const mockTable = {
    name: "TestTable",
    style: "TableStyleLight1",
    load: jest.fn(),
    delete: jest.fn(),
    getDataBodyRange: jest
      .fn()
      .mockReturnValue({ load: jest.fn(), clear: jest.fn() }),
  };
  const mockTableItem = {
    name: "Existing",
    range: { address: "Sheet1!A1:D5" },
  };
  const mockWorksheet = {
    name: "Sheet1",
    load: jest.fn(),
    position: 0,
    getRange: jest.fn().mockReturnValue(mockRange),
    getUsedRangeOrNullObject: jest.fn().mockReturnValue({
      isNullObject: false,
      address: "A1:E10",
      load: jest.fn(),
    }),
    tables: {
      add: jest.fn().mockReturnValue(mockTable),
      getItem: jest.fn().mockReturnValue(mockTable),
      items: [mockTableItem],
      load: jest.fn(),
    },
    pivotTables: {
      add: jest.fn().mockReturnValue({
        name: "Pivot1",
        load: jest.fn(),
        hierarchies: { getItem: jest.fn().mockReturnValue({}) },
        rowHierarchies: { add: jest.fn() },
        columnHierarchies: { add: jest.fn() },
        dataFields: { add: jest.fn().mockReturnValue({}) },
        filterHierarchies: { add: jest.fn() },
      }),
      load: jest.fn(),
    },
    charts: {
      add: jest.fn().mockReturnValue({
        id: "chart-1",
        name: "Chart1",
        title: { text: "", visible: false },
        setPosition: jest.fn(),
        load: jest.fn(),
      }),
      getItem: jest.fn().mockReturnValue({ delete: jest.fn() }),
      load: jest.fn(),
    },
    freezePanes: {
      freezeAt: jest.fn(),
      unfreeze: jest.fn(),
      freezeRows: jest.fn(),
      freezeColumns: jest.fn(),
    },
    delete: jest.fn(),
    copy: jest.fn().mockReturnValue({ name: "CopyOfSheet" }),
  };

  (globalThis as { Excel?: unknown }).Excel = {
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        const ctx = {
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue(mockWorksheet),
              add: jest.fn().mockReturnValue(mockWorksheet),
              getItem: jest.fn().mockReturnValue(mockWorksheet),
            },
            tables: { load: jest.fn() },
          },
          sync: syncMock,
        };
        return fn(ctx);
      }),
  };

  return { syncMock, mockWorksheet, mockRange, mockTable };
}

// ===========================================================================
// S1 manageSheets
// ===========================================================================

describe("S1 manageSheets", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("add — создаёт лист", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageSheets", {
        action: "add",
        name: "NewSheet",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("rename — переименовывает", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageSheets", {
        action: "rename",
        name: "Sheet1",
        newName: "Renamed",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("delete — удаляет", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageSheets", {
        action: "delete",
        name: "Sheet1",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("без action → MISSING_ARGS", async () => {
    const r = JSON.parse(await toolRegistry.execute("manageSheets", {}));
    expect(r.error?.code).toBe("MISSING_ARGS");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("manageSheets")).toBe("moderate");
    expect(toolRegistry.requiresUndo("manageSheets")).toBe(true);
  });
});

// ===========================================================================
// S2 manageTable
// ===========================================================================

describe("S2 manageTable", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });

  test("list — возвращает таблицы", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageTable", { action: "list" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.tables).toBeDefined();
  });

  test("create — создаёт таблицу", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageTable", {
        action: "create",
        address: "A1:E10",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("delete — удаляет таблицу", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageTable", {
        action: "delete",
        tableName: "TestTable",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("create без address → MISSING_ADDRESS", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageTable", { action: "create" }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("manageTable")).toBe("moderate");
    expect(toolRegistry.requiresUndo("manageTable")).toBe(true);
  });
});

// ===========================================================================
// S3 createPivotTable
// ===========================================================================

describe("S3 createPivotTable", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("создаёт сводную таблицу", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("createPivotTable", {
        name: "MyPivot",
        sourceAddress: "A1:E100",
        destinationAddress: "G1",
        rows: ["Category"],
        values: [{ column: "Amount", agg: "sum" }],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.name).toBe("Pivot1");
  });

  test("без name → MISSING_ARGS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("createPivotTable", {
        sourceAddress: "A1:E100",
        destinationAddress: "G1",
      }),
    );
    expect(r.error?.code).toBe("MISSING_ARGS");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("createPivotTable")).toBe("moderate");
    expect(toolRegistry.requiresUndo("createPivotTable")).toBe(true);
  });
});

// ===========================================================================
// S4 createChart
// ===========================================================================

describe("S4 createChart", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("создаёт диаграмму", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("createChart", {
        chartType: "ColumnClustered",
        address: "A1:D20",
        title: "Sales Chart",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.name).toBe("Chart1");
  });

  test("без chartType → MISSING_ARGS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("createChart", { address: "A1:D20" }),
    );
    expect(r.error?.code).toBe("MISSING_ARGS");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("createChart")).toBe("moderate");
    expect(toolRegistry.requiresUndo("createChart")).toBe(true);
  });
});

// ===========================================================================
// S5 freezePanes
// ===========================================================================

describe("S5 freezePanes", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("firstRow — фиксирует строку", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("freezePanes", { target: "firstRow" }),
    );
    expect(r.ok).toBe(true);
  });

  test("none — снимает фиксацию", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("freezePanes", { target: "none" }),
    );
    expect(r.ok).toBe(true);
  });

  test("без target → MISSING_TARGET", async () => {
    const r = JSON.parse(await toolRegistry.execute("freezePanes", {}));
    expect(r.error?.code).toBe("MISSING_TARGET");
  });

  test("риск safe, не требует undo", () => {
    expect(toolRegistry.riskLevel("freezePanes")).toBe("safe");
    expect(toolRegistry.requiresUndo("freezePanes")).toBe(false);
  });
});
