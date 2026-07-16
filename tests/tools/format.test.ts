/**
 * Тесты для format.ts — F1-F5 инструменты.
 *
 * Паттерн: мок globalThis.Excel + мок undoManager.
 */
import { toolRegistry } from "../../src/taskpane/tools/registry";

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

import "../../src/taskpane/tools/format";

function setupExcelMock() {
  const syncMock = jest.fn().mockResolvedValue(undefined);
  const mockRange = {
    rowCount: 5,
    columnCount: 3,
    address: "Test!A1:C5",
    load: jest.fn(),
    numberFormat: [["General"]],
    getCell: jest.fn(),
    format: {
      font: { bold: false, italic: false, size: 11, color: "#000000" },
      fill: { color: "#FFFFFF" },
      horizontalAlignment: "General",
      verticalAlignment: "Bottom",
      wrapText: false,
      autofitColumns: jest.fn(),
      borders: {
        getItem: jest.fn().mockReturnValue({ style: "None", color: "#000000" }),
      },
    },
    conditionalFormats: {
      add: jest.fn().mockReturnValue({
        colorScale: { criteria: {} },
        bar: { fill: { color: "" }, showBarOnly: false },
        top10: { rank: 10, bottom: false, percent: false },
        cellValue: {
          rule: {},
          format: { fill: { color: "" }, font: { color: "" } },
        },
      }),
    },
  };

  const mockSheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockReturnValue(mockRange),
    getUsedRangeOrNullObject: jest.fn().mockReturnValue(
      Object.assign({}, mockRange, {
        isNullObject: false,
        format: mockRange.format,
      }),
    ),
    tables: {
      add: jest.fn().mockReturnValue({
        name: "Table1",
        style: "TableStyleLight1",
        load: jest.fn(),
      }),
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
// F1 applyCellFormat
// ===========================================================================

describe("F1 applyCellFormat", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("форматирование шрифта", async () => {
    const { mockRange } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyCellFormat", {
        address: "A1:C5",
        format: {
          bold: true,
          italic: true,
          fontSize: 14,
          fontColor: "#FF0000",
        },
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("заливка и выравнивание", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyCellFormat", {
        address: "A1:C5",
        format: {
          fillColor: "#FFFF00",
          horizontalAlignment: "center",
          wrapText: true,
        },
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyCellFormat", { format: { bold: true } }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("applyCellFormat")).toBe("moderate");
    expect(toolRegistry.requiresUndo("applyCellFormat")).toBe(true);
  });
});

// ===========================================================================
// F2 applyNumberFormat
// ===========================================================================

describe("F2 applyNumberFormat", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("предустановка currency", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyNumberFormat", {
        address: "A1:C5",
        format: "currency",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.format).toBe("currency");
  });

  test("произвольный формат", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyNumberFormat", {
        address: "A1:C5",
        format: "#,##0.00",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("без format → MISSING_FORMAT", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyNumberFormat", {
        address: "A1:C5",
        format: "",
      }),
    );
    expect(r.error?.code).toBe("MISSING_FORMAT");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("applyNumberFormat")).toBe("moderate");
    expect(toolRegistry.requiresUndo("applyNumberFormat")).toBe(true);
  });
});

// ===========================================================================
// F3 applyConditionalFormat
// ===========================================================================

describe("F3 applyConditionalFormat", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("colorScale правило", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "A1:C5",
        rules: [
          { type: "colorScale", minColor: "#F8696B", maxColor: "#63BE7B" },
        ],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.rulesCount).toBe(1);
  });

  test("highlightCell правило", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "A1:C5",
        rules: [
          {
            type: "highlightCell",
            operator: "greaterThan",
            value1: 1000,
            fillColor: "#FF0000",
          },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("без rules → MISSING_RULES", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "A1:C5",
        rules: [],
      }),
    );
    expect(r.error?.code).toBe("MISSING_RULES");
  });

  test("риск moderate, не требует undo", () => {
    expect(toolRegistry.riskLevel("applyConditionalFormat")).toBe("moderate");
    expect(toolRegistry.requiresUndo("applyConditionalFormat")).toBe(false);
  });
});

// ===========================================================================
// F4 formatAsTable
// ===========================================================================

describe("F4 formatAsTable", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("создание таблицы", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("formatAsTable", { address: "A1:D20" }),
    );
    expect(r.ok).toBe(true);
  });

  test("с именем и стилем", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("formatAsTable", {
        address: "A1:D20",
        tableName: "MyTable",
        style: "TableStyleMedium9",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("formatAsTable")).toBe("moderate");
    expect(toolRegistry.requiresUndo("formatAsTable")).toBe(true);
  });
});

// ===========================================================================
// F5 autoFitColumns
// ===========================================================================

describe("F5 autoFitColumns", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("авто-ширина по диапазону", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("autoFitColumns", { address: "A1:C5" }),
    );
    expect(r.ok).toBe(true);
  });

  test("авто-ширина по всему листу", async () => {
    setupExcelMock();
    const r = JSON.parse(await toolRegistry.execute("autoFitColumns", {}));
    expect(r.ok).toBe(true);
  });

  test("риск safe, не требует undo", () => {
    expect(toolRegistry.riskLevel("autoFitColumns")).toBe("safe");
    expect(toolRegistry.requiresUndo("autoFitColumns")).toBe(false);
  });
});
