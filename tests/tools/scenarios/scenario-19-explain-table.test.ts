/**
 * Сценарий №19: «Объяснить что в таблице»
 *
 * Инструменты: R1 getWorkbookOverview, R3 getRangeStats, R5 findAnomalies
 *
 * Сценарий:
 *   1. Получить обзор книги (структура, листы, таблицы)
 *   2. Получить статистику данных (min/max/sum/avg)
 *   3. Найти аномалии (выбросы, пустоты)
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
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

const ANALYSIS_DATA = [
  ["Товар", "Цена", "Количество", "Сумма"],
  ["Ноутбук", 75000, 10, 750000],
  ["Мышь", 1500, 200, 300000],
  ["Монитор", 25000, 5, 125000],
  ["Клавиатура", 3500, 50, 175000],
  ["Наушники", 5000, null, 0],
  ["Колонки", -1, -1, 1],
];

function setupAnalysisMock(state: ScenarioState) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });
  const tableMock = {
    name: "Table1",
    range: { address: "A1:D7" },
    load: jest.fn(),
  };
  const usedRangeMock = Object.assign(
    {},
    createMockRange("Data!A1:D7", state, { rows: 7, cols: 4 }),
    {
      isNullObject: false,
      getRow: jest.fn().mockReturnValue({
        load: jest.fn(),
        values: [["Товар", "Цена", "Количество", "Сумма"]],
      }),
      getColumn: jest.fn().mockReturnValue({
        load: jest.fn(),
        values: [
          [
            "Товар",
            "Ноутбук",
            "Мышь",
            "Монитор",
            "Клавиатура",
            "Наушники",
            "Колонки",
          ],
        ],
      }),
    },
  );

  const sheetMock = {
    name: "Data",
    load: jest.fn(),
    getUsedRangeOrNullObject: jest.fn().mockReturnValue(usedRangeMock),
    getRange: jest.fn().mockImplementation((addr: string) => {
      const r = createMockRange(addr, state);
      Object.defineProperty(r, "values", {
        get: () => ANALYSIS_DATA,
        set: () => {},
        configurable: true,
      });
      r.rowCount = 7;
      r.columnCount = 4;
      return r;
    }),
    tables: { load: jest.fn() },
  };

  const sheets = {
    load: jest.fn(),
    items: [sheetMock],
    getActiveWorksheet: jest.fn().mockReturnValue(sheetMock),
    getItem: jest.fn().mockReturnValue(sheetMock),
    getCount: jest.fn().mockReturnValue({ load: jest.fn() }),
  };

  (globalThis as { Excel?: unknown }).Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    CalculationMode: { manual: "manual", automatic: "automatic" },
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        return fn({
          workbook: {
            worksheets: sheets,
            tables: { load: jest.fn(), items: [] },
          },
          sync: syncMock,
          runtime: { enableEvents: false },
          application: { calculationMode: "" },
        });
      }),
  };
}

describe("Сценарий 19: Объяснить что в таблице", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: getWorkbookOverview", async () => {
    setupAnalysisMock(state);
    const r = JSON.parse(await toolRegistry.execute("getWorkbookOverview", {}));
    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
  });

  test("Шаг 2: getRangeStats — статистика колонок", async () => {
    setupAnalysisMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("getRangeStats", { address: "A1:D7" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
  });

  test("Шаг 3: findAnomalies", async () => {
    setupAnalysisMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("findAnomalies", { address: "A1:D7" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
  });

  test("ПОЛНЫЙ ФЛОУ: overview → stats → anomalies", async () => {
    setupAnalysisMock(state);

    const step1 = JSON.parse(
      await toolRegistry.execute("getWorkbookOverview", {}),
    );
    expect(step1.ok).toBe(true);

    const step2 = JSON.parse(
      await toolRegistry.execute("getRangeStats", { address: "A1:D7" }),
    );
    expect(step2.ok).toBe(true);

    const step3 = JSON.parse(
      await toolRegistry.execute("findAnomalies", { address: "A1:D7" }),
    );
    expect(step3.ok).toBe(true);
    expect(step3.data).toBeDefined();
  });
});
