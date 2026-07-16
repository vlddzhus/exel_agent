/**
 * Сценарий №16: «Условный формат: красным выше среднего»
 *
 * Инструменты: R3 getRangeStats, F3 applyConditionalFormat
 *
 * Сценарий:
 *   1. Получить статистику (среднее) через getRangeStats
 *   2. Применить условное форматирование (highlightCell: больше среднего)
 */
import { toolRegistry } from "../../../src/taskpane/tools/registry";

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

jest.mock("../../../src/taskpane/tools/_shared/performance", () => {
  const actual = jest.requireActual(
    "../../../src/taskpane/tools/_shared/performance",
  );
  return {
    ...actual,
    withPerformanceGuard: jest
      .fn()
      .mockImplementation(async (cb: (ctx: unknown) => Promise<unknown>) => {
        return Excel.run(async (context) => cb(context));
      }),
  };
});

import "../../../src/taskpane/tools/read";
import "../../../src/taskpane/tools/format";
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

const SALES_DATA = [
  ["Товар", "Продажи"],
  ["Ноутбук", 120000],
  ["Мышь", 1500],
  ["Монитор", 35000],
  ["Клавиатура", 5000],
  ["Наушники", 8000],
];

function setupCfMock(state: ScenarioState, testData?: unknown[][]) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });
  const data = testData ?? SALES_DATA;

  const sheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockImplementation((addr: string) => {
      const r = createMockRange(addr, state);
      Object.defineProperty(r, "values", {
        get: () => data,
        set: () => {},
        configurable: true,
      });
      r.rowCount = data.length;
      r.columnCount = (data[0] ?? []).length;
      return r;
    }),
    getUsedRangeOrNullObject: jest.fn().mockReturnValue(
      Object.assign({}, createMockRange("TestSheet!A1:B10", state), {
        isNullObject: false,
      }),
    ),
    conditionalFormats: {
      add: jest.fn().mockReturnValue({ load: jest.fn(), getName: jest.fn() }),
    },
  };

  (globalThis as { Excel?: unknown }).Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    CalculationMode: { manual: "manual", automatic: "automatic" },
    ConditionalFormatType: {
      cellValue: "CellValue",
      colorScale: "ColorScale",
      dataBar: "DataBar",
    },
    ConditionalCellValueOperator: { greaterThan: "GreaterThan" },
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        return fn({
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue(sheet),
              getItem: jest.fn().mockReturnValue(sheet),
            },
          },
          sync: syncMock,
          runtime: { enableEvents: false },
          application: { calculationMode: "" },
        });
      }),
  };
}

describe("Сценарий 16: Условный формат: красным выше среднего", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: getRangeStats — получить статистику", async () => {
    setupCfMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("getRangeStats", { address: "A1:B6" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
  });

  test("Шаг 2: applyConditionalFormat — highlightCell greater than", async () => {
    setupCfMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "B2:B6",
        rules: [
          {
            type: "highlightCell",
            operator: "greaterThan",
            value: 10000,
            fillColor: "#FF0000",
            fontColor: "#FFFFFF",
          },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("applyConditionalFormat — colorScale", async () => {
    setupCfMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "B2:B6",
        rules: [
          { type: "colorScale", minColor: "#FFFFFF", maxColor: "#FF0000" },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("applyConditionalFormat — dataBar", async () => {
    setupCfMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        address: "B2:B6",
        rules: [{ type: "dataBar", fillColor: "#5B9BD5" }],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("applyConditionalFormat без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("applyConditionalFormat", {
        rule: { type: "highlightCell", operator: "greaterThan", value: 100 },
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });
});
