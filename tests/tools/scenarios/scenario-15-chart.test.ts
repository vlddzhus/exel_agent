/**
 * Сценарий №15: «Столбчатый график по месяцам»
 *
 * Инструменты: S4 createChart
 *
 * Сценарий:
 *   1. Создать столбчатую диаграмму (ColumnClustered)
 *   2. Создать круговую диаграмму (Pie)
 *   3. Создать на отдельном листе (newSheet)
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

import "../../../src/taskpane/tools/structure";
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

function setupChartMock(state: ScenarioState) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });
  const chartMock = {
    id: "chart_1",
    name: "Chart1",
    load: jest.fn(),
    title: { text: "", visible: false },
    delete: jest.fn(),
  };
  const newSheet = {
    name: "График_123",
    load: jest.fn(),
    charts: {
      add: jest.fn().mockReturnValue({
        ...chartMock,
        id: "chart_2",
        name: "Chart2",
        load: jest.fn(),
      }),
    },
  };

  const sheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest
      .fn()
      .mockImplementation((addr: string) => createMockRange(addr, state)),
    getUsedRangeOrNullObject: jest.fn(),
    charts: {
      add: jest.fn().mockReturnValue(chartMock),
      getItem: jest.fn().mockReturnValue(chartMock),
      load: jest.fn(),
    },
  };

  const sheets = {
    getActiveWorksheet: jest.fn().mockReturnValue(sheet),
    getItem: jest.fn().mockReturnValue(sheet),
    add: jest.fn().mockReturnValue(newSheet),
  };

  (globalThis as { Excel?: unknown }).Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    CalculationMode: { manual: "manual", automatic: "automatic" },
    ChartType: { columnClustered: "ColumnClustered", line: "Line", pie: "Pie" },
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        return fn({
          workbook: { worksheets: sheets },
          sync: syncMock,
          runtime: { enableEvents: false },
          application: { calculationMode: "" },
        });
      }),
  };
}

describe("Сценарий 15: Столбчатый график по месяцам", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: createChart — ColumnClustered", async () => {
    setupChartMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("createChart", {
        address: "A1:D6",
        chartType: "ColumnClustered",
        title: "Продажи по месяцам",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
  });

  test("Шаг 2: createChart — Pie", async () => {
    setupChartMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("createChart", {
        address: "A1:B6",
        chartType: "Pie",
        title: "Доли рынка",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("Шаг 3: createChart — на отдельном листе (newSheet)", async () => {
    setupChartMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("createChart", {
        address: "A1:D6",
        chartType: "ColumnClustered",
        position: "newSheet",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("createChart без обязательных аргументов → MISSING_ARGS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("createChart", { address: "A1" }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ARGS");
  });
});
