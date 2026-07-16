/**
 * Тесты для format-extended.ts — F7-F11 инструменты + applyAutoDesign.
 *
 * Итерация «Профессиональное форматирование».
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

import "../../src/taskpane/tools/format-extended";
import "../../src/taskpane/tools/auto-designer";

function setupExcelMock(options?: {
  values?: unknown[][];
  rowCount?: number;
  colCount?: number;
}) {
  const syncMock = jest.fn().mockResolvedValue(undefined);
  const data =
    options?.values ??
    [
      ["Товар", "Цена", "Кол-во", "Сумма"],
      ["Ноутбук", 45000, 5, 225000],
      ["Мышь", 1200, 10, 12000],
      ["Монитор", 18000, 3, 54000],
    ];

  const makeRange = (addr: string): any => {
    const range: any = {
      rowCount: options?.rowCount ?? data.length,
      columnCount: options?.colCount ?? (data[0]?.length ?? 4),
      address: addr,
      values: data,
      numberFormat: [["General"]],
      style: "Normal",
      load: jest.fn(),
      getCell: jest.fn().mockImplementation(function (this: any, r: number, c: number) {
        return makeRange(`${addr}!${r}x${c}`);
      }),
      getResizedRange: jest.fn().mockImplementation(function (
        this: any,
        dr: number,
        dc: number,
      ) {
        const sub = makeRange(`${addr}~${dr}x${dc}`);
        sub.rowCount = Math.abs(dr) + 1;
        sub.columnCount = Math.abs(dc) + 1;
        return sub;
      }),
      format: {
        font: {
          bold: false,
          italic: false,
          size: 11,
          color: "#000000",
          name: "Calibri",
          underline: "None",
          strikethrough: false,
        },
        fill: { color: "#FFFFFF" },
        horizontalAlignment: "General",
        verticalAlignment: "Bottom",
        wrapText: false,
        indentLevel: 0,
        rowHeight: 15,
        columnWidth: 8.43,
        protection: { locked: true },
        autofitColumns: jest.fn(),
        autofitRows: jest.fn(),
        borders: {
          getItem: jest.fn().mockReturnValue({ style: "None", color: "#000000" }),
        },
      },
    };
    return range;
  };

  const mockRange = makeRange("Test!A1:D4");

  const mockSheet = {
    name: "TestSheet",
    load: jest.fn(),
    tabColor: "",
    getRange: jest.fn().mockImplementation((addr: string) => makeRange(addr)),
    getUsedRangeOrNullObject: jest.fn().mockReturnValue(
      Object.assign(mockRange, { isNullObject: false }),
    ),
  };

  (globalThis as { Excel?: unknown }).Excel = {
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        const ctx = {
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue(mockSheet),
              getItem: jest.fn().mockReturnValue(mockSheet),
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
// F7 setRowHeights
// ===========================================================================

describe("F7 setRowHeights", () => {
  beforeEach(() => {
    setupExcelMock();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("установка высоты одной строки", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setRowHeights", {
        address: "A1:A1",
        height: 25,
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("установка высоты диапазона строк", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setRowHeights", {
        address: "A1:A5",
        height: 30,
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setRowHeights", { height: 25 }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("некорректная высота → INVALID_HEIGHT", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setRowHeights", {
        address: "A1:A1",
        height: -5,
      }),
    );
    expect(r.error?.code).toBe("INVALID_HEIGHT");
  });

  test("риск safe, не требует undo", () => {
    expect(toolRegistry.riskLevel("setRowHeights")).toBe("safe");
    expect(toolRegistry.requiresUndo("setRowHeights")).toBe(false);
  });
});

// ===========================================================================
// F8 autoFitRows
// ===========================================================================

describe("F8 autoFitRows", () => {
  beforeEach(() => {
    setupExcelMock();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("авто-высота по диапазону", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("autoFitRows", { address: "A1:D10" }),
    );
    expect(r.ok).toBe(true);
  });

  test("авто-высота по всему листу", async () => {
    const r = JSON.parse(await toolRegistry.execute("autoFitRows", {}));
    expect(r.ok).toBe(true);
  });

  test("риск safe, не требует undo", () => {
    expect(toolRegistry.riskLevel("autoFitRows")).toBe("safe");
    expect(toolRegistry.requiresUndo("autoFitRows")).toBe(false);
  });
});

// ===========================================================================
// F9 copyFormat
// ===========================================================================

describe("F9 copyFormat", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
    setupExcelMock();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("копирование всего стиля (по умолчанию)", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("copyFormat", {
        sourceAddress: "A1:A1",
        targetAddress: "B1:B1",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.aspects).toEqual([
      "font",
      "fill",
      "border",
      "numberFormat",
      "alignment",
    ]);
  });

  test("выборочное копирование (font + fill)", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("copyFormat", {
        sourceAddress: "A1:A1",
        targetAddress: "B1:B1",
        what: ["font", "fill"],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.aspects).toEqual(["font", "fill"]);
  });

  test("без source/target → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("copyFormat", {
        sourceAddress: "A1:A1",
      }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("copyFormat")).toBe("moderate");
    expect(toolRegistry.requiresUndo("copyFormat")).toBe(true);
  });
});

// ===========================================================================
// F10 applyNamedStyle
// ===========================================================================

describe("F10 applyNamedStyle", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
    setupExcelMock();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("применение стиля Good", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyNamedStyle", {
        address: "A1:A1",
        style: "Good",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("применение стиля Bad", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyNamedStyle", {
        address: "A1:A1",
        style: "Bad",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("применение стиля Input", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyNamedStyle", {
        address: "A1:A1",
        style: "Input",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("без style → MISSING_STYLE", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyNamedStyle", {
        address: "A1:A1",
        style: "",
      }),
    );
    expect(r.error?.code).toBe("MISSING_STYLE");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("applyNamedStyle")).toBe("moderate");
    expect(toolRegistry.requiresUndo("applyNamedStyle")).toBe(true);
  });
});

// ===========================================================================
// F11 setSheetTabColor
// ===========================================================================

describe("F11 setSheetTabColor", () => {
  beforeEach(() => {
    setupExcelMock();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("установка цвета ярлычка активного листа", async () => {
    const { mockSheet } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("setSheetTabColor", { color: "#FF0000" }),
    );
    expect(r.ok).toBe(true);
  });

  test("установка цвета по имени листа", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setSheetTabColor", {
        sheetName: "TestSheet",
        color: "#0000FF",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("сброс цвета (none)", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setSheetTabColor", { color: "none" }),
    );
    expect(r.ok).toBe(true);
  });

  test("без color → MISSING_COLOR", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("setSheetTabColor", { color: "" }),
    );
    expect(r.error?.code).toBe("MISSING_COLOR");
  });

  test("риск safe, не требует undo", () => {
    expect(toolRegistry.riskLevel("setSheetTabColor")).toBe("safe");
    expect(toolRegistry.requiresUndo("setSheetTabColor")).toBe(false);
  });
});

// ===========================================================================
// applyAutoDesign
// ===========================================================================

describe("applyAutoDesign", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
    setupExcelMock();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("auto intent — автоопределение финансового профиля", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyAutoDesign", {
        address: "A1:D4",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.intent).toBeDefined();
    expect(r.data.columnTypes).toEqual(
      expect.arrayContaining([
        expect.stringMatching(/text|number|currency|percent|date|boolean|id/),
      ]),
    );
  });

  test("financial intent — принудительный финансовый стиль", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyAutoDesign", {
        address: "A1:D4",
        intent: "financial",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.intent).toBe("financial");
    expect(r.data.detectedIntent).toBeUndefined();
  });

  test("professional intent", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyAutoDesign", {
        address: "A1:D4",
        intent: "professional",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.intent).toBe("professional");
  });

  test("dashboard intent", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyAutoDesign", {
        address: "A1:D4",
        intent: "dashboard",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.palette.headerFill).toBe("#27AE60");
  });

  test("minimal intent — без чередования", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyAutoDesign", {
        address: "A1:D4",
        intent: "minimal",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.intent).toBe("minimal");
  });

  test("без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(await toolRegistry.execute("applyAutoDesign", {}));
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("hasHeader=false — обрабатывает таблицу без заголовка", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyAutoDesign", {
        address: "A1:D4",
        hasHeader: false,
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("applyAutoDesign")).toBe("moderate");
    expect(toolRegistry.requiresUndo("applyAutoDesign")).toBe(true);
  });

  test("определяет типы колонок в данных", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyAutoDesign", {
        address: "A1:D4",
        intent: "professional",
      }),
    );
    expect(r.ok).toBe(true);
    // Первая колонка "Товар" — text
    expect(r.data.columnTypes[0]).toBe("text");
  });
});
