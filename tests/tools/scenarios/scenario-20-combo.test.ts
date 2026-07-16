/**
 * Сценарий №20: «Комбо: очисти → посчитай → сводная → график»
 *
 * Инструменты: R4 detectDataTypes, T3 removeDuplicates, T5 normalizeText,
 *              R3 getRangeStats, S3 createPivotTable, S4 createChart
 *
 * Сценарий:
 *   1. Очистить данные: определить типы, удалить дубликаты, нормализовать
 *   2. Посчитать статистику
 *   3. Создать сводную таблицу
 *   4. Построить график
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
import "../../../src/taskpane/tools/transform";
import "../../../src/taskpane/tools/structure";
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

const COMBO_DATA = [
  ["Менеджер", "Месяц", "Сумма"],
  ["  Иван  ", "Янв", 100000],
  ["  Мария  ", "Янв", 150000],
  ["Иван", "Янв", 100000],
  ["  Пётр  ", "Фев", 80000],
  ["  Мария  ", "Фев", 200000],
  ["Пётр", "Фев", 80000],
];

function setupComboMock(state: ScenarioState) {
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
    rowCount: 5,
    columnCount: 4,
    hierarchies: { getItem: jest.fn().mockReturnValue(hierarchyMock) },
    rowHierarchies: { add: jest.fn() },
    columnHierarchies: { add: jest.fn() },
    dataFields: { add: jest.fn().mockReturnValue(dataFieldMock) },
    filterHierarchies: { add: jest.fn() },
  };

  const sheet = {
    name: "Data",
    load: jest.fn(),
    getRange: jest.fn().mockImplementation((addr: string) => {
      const r = createMockRange(addr, state);
      Object.defineProperty(r, "values", {
        get: () => COMBO_DATA,
        set: (v: unknown[][]) => {
          state.writtenValues = v;
        },
        configurable: true,
      });
      r.rowCount = COMBO_DATA.length;
      r.columnCount = (COMBO_DATA[0] ?? []).length;
      return r;
    }),
    getUsedRangeOrNullObject: jest
      .fn()
      .mockReturnValue(
        Object.assign(
          {},
          createMockRange("Data!A1:C7", state, { rows: 7, cols: 3 }),
          { isNullObject: false },
        ),
      ),
    charts: {
      add: jest.fn().mockReturnValue({
        id: "chart_c",
        name: "ChartC",
        load: jest.fn(),
        title: { text: "", visible: false },
        delete: jest.fn(),
      }),
    },
    pivotTables: {
      add: jest.fn().mockReturnValue(pivotTableMock),
      load: jest.fn(),
    },
    tables: { load: jest.fn() },
  };

  (globalThis as { Excel?: unknown }).Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    CalculationMode: { manual: "manual", automatic: "automatic" },
    ChartType: { columnClustered: "ColumnClustered", pie: "Pie" },
    AggregationFunction: { sum: "Sum", count: "Count", average: "Average" },
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

describe("Сценарий 20: Комбо: очисти → посчитай → сводная → график", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: detectDataTypes", async () => {
    setupComboMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("detectDataTypes", { address: "A1:C7" }),
    );
    expect(r.ok).toBe(true);
  });

  test("Шаг 2: removeDuplicates", async () => {
    setupComboMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("removeDuplicates", {
        address: "A1:C7",
        columns: [0, 1, 2],
        hasHeaders: true,
      }),
    );
    expect(r.ok).toBe(true);
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("Шаг 3: normalizeText (trim менеджеров)", async () => {
    setupComboMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "A1:A7",
        operations: ["trim"],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("Шаг 4: createPivotTable", async () => {
    setupComboMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("createPivotTable", {
        sourceAddress: "A1:C7",
        destinationAddress: "E1",
        name: "СводнаяПродажи",
        rows: ["Менеджер"],
        columns: ["Месяц"],
        values: [{ column: "Сумма", agg: "sum" }],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("Шаг 5: createChart на основе данных", async () => {
    setupComboMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("createChart", {
        address: "A1:C7",
        chartType: "ColumnClustered",
        title: "Продажи",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("ПОЛНЫЙ ФЛОУ: 5 шагов последовательно", async () => {
    setupComboMock(state);

    const step1 = JSON.parse(
      await toolRegistry.execute("detectDataTypes", { address: "A1:C7" }),
    );
    expect(step1.ok).toBe(true);

    const step2 = JSON.parse(
      await toolRegistry.execute("removeDuplicates", {
        address: "A1:C7",
        columns: [0, 1, 2],
      }),
    );
    expect(step2.ok).toBe(true);

    const step3 = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "A1:A7",
        operations: ["trim"],
      }),
    );
    expect(step3.ok).toBe(true);

    const step4 = JSON.parse(
      await toolRegistry.execute("createPivotTable", {
        sourceAddress: "A1:C7",
        destinationAddress: "E1",
        name: "Итог",
        rows: ["Менеджер"],
        values: [{ column: "Сумма", agg: "sum" }],
      }),
    );
    expect(step4.ok).toBe(true);

    const step5 = JSON.parse(
      await toolRegistry.execute("createChart", {
        address: "A1:C7",
        chartType: "ColumnClustered",
      }),
    );
    expect(step5.ok).toBe(true);
  });
});
