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
    style: "Normal",
    getCell: jest.fn().mockImplementation(function (
      this: any,
      _r: number,
      _c: number,
    ) {
      const sub = Object.assign(Object.create(this), {
        rowCount: 1,
        columnCount: 1,
        address: `${this.address}!${_r}x${_c}`,
      });
      return sub;
    }),
    getResizedRange: jest.fn().mockImplementation(function (
      this: any,
      _dr: number,
      _dc: number,
    ) {
      const sub = Object.assign(Object.create(this), {
        rowCount: Math.abs(_dr) + 1,
        columnCount: Math.abs(_dc) + 1,
        address: `${this.address}~${_dr}x${_dc}`,
      });
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
    conditionalFormats: {
      add: jest.fn().mockReturnValue({
        colorScale: { criteria: {} },
        bar: { fill: { color: "" }, showBarOnly: false },
        top10: { rank: 10, bottom: false, percent: false },
        cellValue: {
          rule: {},
          format: { fill: { color: "" }, font: { color: "" } },
        },
        iconSet: {
          iconSet: "ThreeTrafficLights1",
          reverseIconOrder: false,
          showIconOnly: false,
          format: { fill: { color: "" }, font: { color: "" } },
        },
        customRule: {
          formula: "",
          format: { fill: { color: "" }, font: { color: "" } },
        },
        presetCriteria: {
          rule: { type: "DuplicateValues" },
          format: { fill: { color: "" }, font: { color: "" } },
        },
      }),
    },
  };

  const mockSheet = {
    name: "TestSheet",
    load: jest.fn(),
    tabColor: "",
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
    expect(mockRange.format.font.bold).toBe(true);
    expect(mockRange.format.font.italic).toBe(true);
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

  // ── Новые поля (итерация «Профессиональное форматирование») ──

  test("fontName — установка шрифта Arial", async () => {
    const { mockRange } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyCellFormat", {
        address: "A1:C5",
        format: { fontName: "Arial" },
      }),
    );
    expect(r.ok).toBe(true);
    expect(mockRange.format.font.name).toBe("Arial");
  });

  test("underline + strikethrough", async () => {
    const { mockRange } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyCellFormat", {
        address: "A1:C5",
        format: { underline: "single", strikethrough: true },
      }),
    );
    expect(r.ok).toBe(true);
    expect(mockRange.format.font.underline).toBe("single");
    expect(mockRange.format.font.strikethrough).toBe(true);
  });

  test("indentLevel — отступ 2", async () => {
    const { mockRange } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyCellFormat", {
        address: "A1:C5",
        format: { indentLevel: 2 },
      }),
    );
    expect(r.ok).toBe(true);
    expect(mockRange.format.indentLevel).toBe(2);
  });

  test("indentLevel клипится к диапазону 0-15", async () => {
    const { mockRange } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyCellFormat", {
        address: "A1:C5",
        format: { indentLevel: 50 },
      }),
    );
    expect(r.ok).toBe(true);
    expect(mockRange.format.indentLevel).toBe(15);
  });

  test("locked — блокировка ячеек", async () => {
    const { mockRange } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyCellFormat", {
        address: "A1:C5",
        format: { locked: true },
      }),
    );
    expect(r.ok).toBe(true);
    expect(mockRange.format.protection.locked).toBe(true);
  });

  test("граница со стилем Double (расширенный border)", async () => {
    const { mockRange } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyCellFormat", {
        address: "A1:C5",
        format: {
          borderTop: { style: "Double", color: "#000000" },
        },
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("комбинация всех новых полей", async () => {
    const { mockRange } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyCellFormat", {
        address: "A1:C5",
        format: {
          fontName: "Segoe UI",
          underline: "double",
          strikethrough: false,
          indentLevel: 1,
          locked: false,
          bold: true,
        },
      }),
    );
    expect(r.ok).toBe(true);
    expect(mockRange.format.font.name).toBe("Segoe UI");
    expect(mockRange.format.font.bold).toBe(true);
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

  // ── Новые типы CF (итерация «Профессиональное форматирование») ──

  test("iconSet — светофор", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "B2:B6",
        rules: [
          {
            type: "iconSet",
            iconSet: "threeTrafficLights1",
            reverseIconOrder: false,
          },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("iconSet — стрелки с showIconOnly", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "B2:B6",
        rules: [
          {
            type: "iconSet",
            iconSet: "threeArrows",
            showIconOnly: true,
          },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("customFormula — просроченные даты", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "A2:A100",
        rules: [
          {
            type: "customFormula",
            formula: "A1<TODAY()-7",
            fillColor: "#FFC7CE",
            fontColor: "#9C0006",
          },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("customFormula с ведущим = — корректно обрезается", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "A2:A100",
        rules: [
          {
            type: "customFormula",
            formula: "=AND(A1>0,A1<100)",
          },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("duplicates — подсветка дубликатов", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "A2:A100",
        rules: [
          {
            type: "duplicates",
            criteria: "duplicateValues",
            fillColor: "#FFC7CE",
          },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("duplicates — uniqueValues", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "A2:A100",
        rules: [
          {
            type: "duplicates",
            criteria: "uniqueValues",
            fillColor: "#C6EFCE",
          },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("несколько правил разного типа в одном вызове", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "A1:D10",
        rules: [
          { type: "colorScale", minColor: "#63BE7B", maxColor: "#F8696B" },
          { type: "iconSet", iconSet: "threeArrows" },
          { type: "duplicates", criteria: "duplicateValues" },
        ],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.rulesCount).toBe(3);
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
