/**
 * Сценарий №12: «Отсортировать по убыванию суммы»
 *
 * Инструменты: T1 sortData
 *
 * Сценарий:
 *   1. Таблица с суммами
 *   2. Сортировка по колонке суммы по убыванию (desc)
 *   3. Сортировка по 2 колонкам
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

import "../../../src/taskpane/tools/transform";
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

function setupSortMock(state: ScenarioState) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });

  (globalThis as { Excel?: unknown }).Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    CalculationMode: { manual: "manual", automatic: "automatic" },
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        return fn({
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue({
                name: "TestSheet",
                load: jest.fn(),
                getRange: jest.fn().mockImplementation((addr: string) => {
                  const r = createMockRange(addr, state);
                  r.rowCount = 6;
                  r.columnCount = 3;
                  return r;
                }),
                getUsedRangeOrNullObject: jest.fn(),
              }),
              getItem: jest.fn(),
            },
          },
          sync: syncMock,
          runtime: { enableEvents: false },
          application: { calculationMode: "" },
        });
      }),
  };
}

describe("Сценарий 12: Отсортировать по убыванию суммы", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: sortData по одной колонке desc", async () => {
    setupSortMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("sortData", {
        address: "A1:C10",
        sortColumns: [{ column: 2, order: "desc" }],
      }),
    );
    expect(r.ok).toBe(true);
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("Шаг 2: sortData по 2 колонкам", async () => {
    setupSortMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("sortData", {
        address: "A1:C10",
        sortColumns: [
          { column: 1, order: "asc" },
          { column: 2, order: "desc" },
        ],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("sortData без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("sortData", { sortColumns: [{ column: 0 }] }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("sortData с пустым sortColumns → EMPTY_COLUMNS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("sortData", {
        address: "A1:C10",
        sortColumns: [],
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("EMPTY_COLUMNS");
  });

  test("sortData с hasHeaders=false", async () => {
    setupSortMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("sortData", {
        address: "A1:C10",
        sortColumns: [{ column: 0, order: "asc" }],
        hasHeaders: false,
      }),
    );
    expect(r.ok).toBe(true);
  });
});
