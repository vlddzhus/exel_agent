/**
 * Тесты для transform.ts — T1-T6 инструменты.
 */
import { toolRegistry } from "../../src/taskpane/tools/registry";

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

import "../../src/taskpane/tools/transform";

function setupExcelMock() {
  const syncMock = jest.fn().mockResolvedValue(undefined);

  const mockRange = {
    rowCount: 5,
    columnCount: 3,
    address: "Test!A1:C5",
    load: jest.fn(),
    values: [
      ["H", "V1", "V2"],
      ["a", 1, 2],
      ["b", 3, 4],
      ["c", 5, 6],
      ["a", 1, 2],
    ] as unknown[][],
    sort: { apply: jest.fn() },
    removeDuplicates: jest
      .fn()
      .mockReturnValue({ removed: 2, load: jest.fn() }),
    textToColumns: jest.fn(),
  };

  const mockSheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockReturnValue(mockRange),
    getUsedRangeOrNullObject: jest
      .fn()
      .mockReturnValue(Object.assign({}, mockRange, { isNullObject: false })),
    autoFilter: {
      apply: jest.fn(),
      clear: jest.fn(),
    },
  };

  (globalThis as { Excel?: unknown }).Excel = {
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        const ctx = {
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue(mockSheet),
            },
          },
          sync: syncMock,
        };
        return fn(ctx);
      }),
  };

  return { syncCount: syncMock, mockRange, mockSheet };
}

// ===========================================================================
// T1 sortData
// ===========================================================================

describe("T1 sortData", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("сортировка по одной колонке", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("sortData", {
        address: "A1:C5",
        sortColumns: [{ column: 1, order: "desc" }],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("сортировка по двум колонкам", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("sortData", {
        address: "A1:C5",
        sortColumns: [
          { column: 0, order: "asc" },
          { column: 1, order: "desc" },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("sortData", { sortColumns: [{ column: 0 }] }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("пустые sortColumns → EMPTY_COLUMNS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("sortData", {
        address: "A1:C5",
        sortColumns: [],
      }),
    );
    expect(r.error?.code).toBe("EMPTY_COLUMNS");
  });

  test("создаётся undo-снапшот", async () => {
    setupExcelMock();
    await toolRegistry.execute("sortData", {
      address: "A1:C5",
      sortColumns: [{ column: 0 }],
    });
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("sortData")).toBe("moderate");
    expect(toolRegistry.requiresUndo("sortData")).toBe(true);
  });
});

// ===========================================================================
// T2 filterData
// ===========================================================================

describe("T2 filterData", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("включение фильтра без колоночных фильтров", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("filterData", { address: "A1:C5" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.active).toBe(true);
  });

  test("отключение фильтра", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("filterData", {
        address: "A1:C5",
        clear: true,
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.active).toBe(false);
  });

  test("без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(await toolRegistry.execute("filterData", {}));
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("filterData")).toBe("moderate");
    expect(toolRegistry.requiresUndo("filterData")).toBe(true);
  });
});

// ===========================================================================
// T3 removeDuplicates
// ===========================================================================

describe("T3 removeDuplicates", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("удаление дублей", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("removeDuplicates", {
        address: "A1:C5",
        columns: [0, 1],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.removed).toBe(2);
  });

  test("без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("removeDuplicates", { columns: [0] }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("пустые columns → EMPTY_COLUMNS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("removeDuplicates", {
        address: "A1:C5",
        columns: [],
      }),
    );
    expect(r.error?.code).toBe("EMPTY_COLUMNS");
  });

  test("создаётся undo-снапшот", async () => {
    setupExcelMock();
    await toolRegistry.execute("removeDuplicates", {
      address: "A1:C5",
      columns: [0],
    });
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("removeDuplicates")).toBe("moderate");
    expect(toolRegistry.requiresUndo("removeDuplicates")).toBe(true);
  });
});

// ===========================================================================
// T4 splitTextToColumns
// ===========================================================================

describe("T4 splitTextToColumns", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("разбивка с разделителем comma", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("splitTextToColumns", {
        address: "A2:A100",
        delimiter: "comma",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("разбивка с targetStartCell", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("splitTextToColumns", {
        address: "A2:A100",
        delimiter: "space",
        targetStartCell: "B2",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("splitTextToColumns", { delimiter: "comma" }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("неизвестный delimiter → INVALID_DELIMITER", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("splitTextToColumns", {
        address: "A2:A100",
        delimiter: "bad",
      }),
    );
    expect(r.error?.code).toBe("INVALID_DELIMITER");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("splitTextToColumns")).toBe("moderate");
    expect(toolRegistry.requiresUndo("splitTextToColumns")).toBe(true);
  });
});

// ===========================================================================
// T5 normalizeText
// ===========================================================================

describe("T5 normalizeText", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("uppercase + trim", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "B2:B5",
        operations: ["uppercase", "trim"],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("cleanWhitespace", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "B2:B5",
        operations: ["cleanWhitespace"],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", { operations: ["trim"] }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("пустые operations → EMPTY_OPS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "B2:B5",
        operations: [],
      }),
    );
    expect(r.error?.code).toBe("EMPTY_OPS");
  });

  test("создаётся undo-снапшот", async () => {
    setupExcelMock();
    await toolRegistry.execute("normalizeText", {
      address: "B2:B5",
      operations: ["trim"],
    });
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("normalizeText")).toBe("moderate");
    expect(toolRegistry.requiresUndo("normalizeText")).toBe(true);
  });

  test("propercase: pure функция", () => {
    const { normalizeValue } = jest.requireActual(
      "../../src/taskpane/tools/transform",
    );
    expect(normalizeValue("hello world", ["propercase"])).toBe("Hello World");
    expect(normalizeValue("ivan petrov", ["propercase"])).toBe("Ivan Petrov");
  });

  test("чистка пробелов", () => {
    const { normalizeValue } = jest.requireActual(
      "../../src/taskpane/tools/transform",
    );
    expect(normalizeValue("  много   пробелов  ", ["cleanWhitespace"])).toBe(
      "много пробелов",
    );
  });

  test("числовые значения не трогаем", () => {
    const { normalizeValue } = jest.requireActual(
      "../../src/taskpane/tools/transform",
    );
    expect(normalizeValue(42, ["trim"])).toBe(42);
  });
});

// ===========================================================================
// T6 lookup
// ===========================================================================

describe("T6 lookup", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("успешный поиск", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("lookup", {
        lookupAddress: "A1:C5",
        lookupColumn: 0,
        resultColumn: 1,
        lookupValue: "b",
        writeTo: "D2",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.result).toBe(3);
  });

  test("значение не найдено", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("lookup", {
        lookupAddress: "A1:C5",
        lookupColumn: 0,
        resultColumn: 1,
        lookupValue: "NOTFOUND",
        writeTo: "D2",
      }),
    );
    expect(r.error?.code).toBe("NOT_FOUND");
  });

  test("без lookupAddress → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("lookup", {
        lookupColumn: 0,
        resultColumn: 1,
        lookupValue: "x",
        writeTo: "D2",
      }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("без writeTo → MISSING_TARGET", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("lookup", {
        lookupAddress: "A1:C5",
        lookupColumn: 0,
        resultColumn: 1,
        lookupValue: "x",
      }),
    );
    expect(r.error?.code).toBe("MISSING_TARGET");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("lookup")).toBe("moderate");
    expect(toolRegistry.requiresUndo("lookup")).toBe(true);
  });
});
