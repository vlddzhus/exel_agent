/**
 * Тесты для write-инструментов W1-W5 через defineTool API.
 */
import { toolRegistry } from "../../src/taskpane/tools/registry";

// ===========================================================================
// Mocks
// ===========================================================================

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
const mockWithPerformanceGuard = jest.fn();

jest.mock("../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

jest.mock("../../src/taskpane/tools/_shared/performance", () => {
  const actual = jest.requireActual(
    "../../src/taskpane/tools/_shared/performance",
  );
  return {
    ...actual,
    withPerformanceGuard: (callback: (ctx: unknown) => Promise<unknown>) =>
      mockWithPerformanceGuard(callback),
  };
});

import "../../src/taskpane/tools/write";

// ===========================================================================
// Helper: setup Excel mock
// ===========================================================================

interface RangeConfig {
  rowCount?: number;
  columnCount?: number;
  address?: string;
}

function makeRange(cfg?: RangeConfig) {
  const rowCount = cfg?.rowCount ?? 2;
  const columnCount = cfg?.columnCount ?? 2;
  return {
    rowCount,
    columnCount,
    address: cfg?.address ?? "Test!A1:B2",
    load: jest.fn(),
    clear: jest.fn(),
    values: [] as unknown[][],
    formulas: [] as string[][],
    getCell: (_r: number, _c: number) => ({
      getResizedRange: (_dr: number, _dc: number) =>
        makeRange({ rowCount: _dr + 1, columnCount: _dc + 1 }),
    }),
    getResizedRange: (_dr: number, _dc: number) =>
      makeRange({ rowCount: _dr + 1, columnCount: _dc + 1 }),
  };
}

function setupExcelMock(rangeCfg?: RangeConfig) {
  const syncMock = jest.fn().mockResolvedValue(undefined);
  const range = makeRange(rangeCfg);

  const sheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockReturnValue(range),
    getUsedRangeOrNullObject: jest
      .fn()
      .mockReturnValue(
        Object.assign(
          makeRange({
            rowCount: 5,
            columnCount: 3,
            address: "TestSheet!A1:E5",
          }),
          { isNullObject: false },
        ),
      ),
  };

  const mockExcel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        const ctx = {
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue(sheet),
              getItem: jest.fn().mockReturnValue(sheet),
            },
          },
          sync: syncMock,
          runtime: { enableEvents: false },
          application: { calculationMode: "" },
        };
        mockWithPerformanceGuard.mockImplementation(
          async (cb: (ctx: unknown) => Promise<unknown>) => {
            await cb(ctx);
          },
        );
        return fn(ctx);
      }),
  };

  (globalThis as { Excel?: unknown }).Excel = mockExcel;
  return { syncCount: syncMock, sheet, range };
}

// ===========================================================================
// W1. setValues
// ===========================================================================

describe("W1 setValues", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
    mockWithPerformanceGuard.mockReset();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("успешная запись 2×3", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "A1",
        values: [
          [1, 2, 3],
          [4, 5, 6],
        ],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(6);
  });

  test("без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setValues", { values: [[1]] }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("пустой values → EMPTY_VALUES", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setValues", { address: "A1", values: [] }),
    );
    expect(r.error?.code).toBe("EMPTY_VALUES");
  });

  test("создаётся undo-снапшот", async () => {
    setupExcelMock();
    await toolRegistry.execute("setValues", { address: "A1", values: [[1]] });
    expect(mockCreateBackup).toHaveBeenCalledWith(
      "A1",
      "setValues",
      expect.any(Object),
    );
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("setValues")).toBe("moderate");
    expect(toolRegistry.requiresUndo("setValues")).toBe(true);
  });
});

// ===========================================================================
// W2. setFormula
// ===========================================================================

describe("W2 setFormula", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
    mockWithPerformanceGuard.mockReset();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("успешная запись", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "C5",
        formulas: [["=SUM(A1:A3)"], ["=B1*2"]],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(2);
  });

  test("невалидная формула → FORMULA_INVALID", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "C5",
        formulas: [["SUM(A1:A3"]],
      }),
    );
    expect(r.error?.code).toBe("FORMULA_INVALID");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("setFormula")).toBe("moderate");
    expect(toolRegistry.requiresUndo("setFormula")).toBe(true);
  });
});

// ===========================================================================
// W3. fillRange
// ===========================================================================

describe("W3 fillRange", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
    mockWithPerformanceGuard.mockReset();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("заполнение progression 5×1", async () => {
    setupExcelMock({ rowCount: 5, columnCount: 1, address: "A1:A5" });
    const r = JSON.parse(
      await toolRegistry.execute("fillRange", {
        address: "A1:A5",
        fillType: "progression",
        startValue: 10,
        step: 5,
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(5);
  });

  test("неверный fillType → INVALID_FILL_TYPE", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("fillRange", {
        address: "A1",
        fillType: "invalid",
      }),
    );
    expect(r.error?.code).toBe("INVALID_FILL_TYPE");
  });

  test("без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("fillRange", { fillType: "progression" }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("fillRange")).toBe("moderate");
    expect(toolRegistry.requiresUndo("fillRange")).toBe(true);
  });
});

// ===========================================================================
// W4. appendRows
// ===========================================================================

describe("W4 appendRows", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
    mockWithPerformanceGuard.mockReset();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("добавление 2 строк", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("appendRows", {
        values: [
          [1, "Иван"],
          [2, "Мария"],
        ],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(4);
  });

  test("пустой values → EMPTY_VALUES", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("appendRows", { values: [] }),
    );
    expect(r.error?.code).toBe("EMPTY_VALUES");
  });

  test("с указанием листа", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("appendRows", {
        sheetName: "TestSheet",
        values: [["a"]],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.sheetName).toBe("TestSheet");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("appendRows")).toBe("moderate");
    expect(toolRegistry.requiresUndo("appendRows")).toBe(true);
  });
});

// ===========================================================================
// W5. clearRange
// ===========================================================================

describe("W5 clearRange", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
    mockWithPerformanceGuard.mockReset();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("очистка всех данных", async () => {
    const { range } = setupExcelMock({ address: "A1:D10" });
    const r = JSON.parse(
      await toolRegistry.execute("clearRange", { address: "A1:D10" }),
    );
    expect(r.ok).toBe(true);
    expect(range.clear).toHaveBeenCalledWith();
  });

  test("очистка только значений", async () => {
    const { range } = setupExcelMock();
    await toolRegistry.execute("clearRange", {
      address: "A1:D10",
      clearWhat: "values",
    });
    expect(range.clear).toHaveBeenCalledWith("Contents");
  });

  test("очистка только форматов", async () => {
    const { range } = setupExcelMock();
    await toolRegistry.execute("clearRange", {
      address: "A1:D10",
      clearWhat: "formats",
    });
    expect(range.clear).toHaveBeenCalledWith("Formats");
  });

  test("без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(await toolRegistry.execute("clearRange", {}));
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("риск dangerous + требует undo + requiresConfirmation", () => {
    expect(toolRegistry.riskLevel("clearRange")).toBe("dangerous");
    expect(toolRegistry.requiresUndo("clearRange")).toBe(true);
    expect(toolRegistry.requiresConfirmation("clearRange")).toBe(true);
  });
});
