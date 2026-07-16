/**
 * Сценарий №14: «Сводная: продажи × менеджеры × месяцы»
 *
 * Инструменты: R1 getWorkbookOverview, S3 createPivotTable
 *
 * Сценарий:
 *   1. Получить обзор книги
 *   2. Создать сводную таблицу продаж по менеджерам и месяцам
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
import "../../../src/taskpane/tools/structure";
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

function setupPivotMock(state: ScenarioState) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });
  const hierarchyMock = {
    name: "mock",
    load: jest.fn(),
  };
  const dataFieldMock = {
    name: "mock",
    summarizeBy: "Sum",
    load: jest.fn(),
  };
  const pivotTableMock = {
    name: "Сводная",
    load: jest.fn(),
    rowCount: 10,
    columnCount: 5,
    hierarchies: {
      getItem: jest.fn().mockReturnValue(hierarchyMock),
    },
    rowHierarchies: { add: jest.fn() },
    columnHierarchies: { add: jest.fn() },
    dataFields: { add: jest.fn().mockReturnValue(dataFieldMock) },
    filterHierarchies: { add: jest.fn() },
  };

  const sheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest
      .fn()
      .mockImplementation((addr: string) => createMockRange(addr, state)),
    getUsedRangeOrNullObject: jest.fn().mockReturnValue(
      Object.assign({}, createMockRange("TestSheet!A1:E50", state), {
        isNullObject: false,
      }),
    ),
    pivotTables: {
      add: jest.fn().mockReturnValue(pivotTableMock),
      load: jest.fn(),
    },
    tables: { load: jest.fn() },
  };

  const sheet2 = {
    name: "Data",
    load: jest.fn(),
    getRange: jest.fn(),
    getUsedRangeOrNullObject: jest.fn(),
  };

  (globalThis as { Excel?: unknown }).Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    CalculationMode: { manual: "manual", automatic: "automatic" },
    AggregationFunction: {
      sum: "Sum",
      count: "Count",
      average: "Average",
      max: "Max",
      min: "Min",
    },
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        return fn({
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue(sheet),
              getItem: jest.fn().mockReturnValue(sheet2),
              getCount: jest.fn().mockReturnValue({ load: jest.fn() }),
            },
          },
          sync: syncMock,
          runtime: { enableEvents: false },
          application: { calculationMode: "" },
        });
      }),
  };
}

describe("Сценарий 14: Сводная: продажи × менеджеры × месяцы", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: createPivotTable — базовая сводная", async () => {
    setupPivotMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("createPivotTable", {
        sourceAddress: "A1:D50",
        destinationAddress: "F1",
        name: "СводнаяПродажи",
        rows: ["Менеджер"],
        values: [{ column: "Сумма", agg: "sum" }],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
  });

  test("Шаг 2: createPivotTable с колонками (месяцы)", async () => {
    setupPivotMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("createPivotTable", {
        sourceAddress: "A1:D50",
        destinationAddress: "F1",
        name: "Сводная2",
        rows: ["Менеджер"],
        columns: ["Месяц"],
        values: [{ column: "Сумма", agg: "sum" }],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("createPivotTable с average и count", async () => {
    setupPivotMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("createPivotTable", {
        sourceAddress: "A1:D50",
        destinationAddress: "F1",
        name: "СредниеПродажи",
        rows: ["Менеджер"],
        values: [
          { column: "Сумма", agg: "average" },
          { column: "Количество", agg: "count" },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("createPivotTable без обязательных аргументов → MISSING_ARGS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("createPivotTable", {
        sourceAddress: "A1:D50",
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ARGS");
  });
});
