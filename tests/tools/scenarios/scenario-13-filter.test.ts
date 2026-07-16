/**
 * Сценарий №13: «Отфильтровать «только Оплачено»»
 *
 * Инструменты: T2 filterData
 *
 * Сценарий:
 *   1. Таблица со статусами (Оплачено, Ожидание, Отмена)
 *   2. Включить фильтр по статусу = Оплачено
 *   3. Отключить фильтр (показать всё)
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

function setupFilterMock(state: ScenarioState) {
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
                  r.rowCount = 10;
                  r.columnCount = 4;
                  return r;
                }),
                getUsedRangeOrNullObject: jest.fn(),
                autoFilter: { apply: jest.fn(), clear: jest.fn() },
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

describe("Сценарий 13: Отфильтровать «только Оплачено»", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: filterData — включить фильтр по статусу", async () => {
    setupFilterMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("filterData", {
        address: "A1:D10",
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("Шаг 2: filterData — отключить фильтр (clear=true)", async () => {
    setupFilterMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("filterData", {
        address: "A1:D10",
        clear: true,
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("filterData без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(await toolRegistry.execute("filterData", {}));
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });
});
