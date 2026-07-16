/**
 * Сценарий №18: «Авто-ширина колонок»
 *
 * Инструменты: F5 autoFitColumns
 *
 * Сценарий:
 *   1. Авто-подбор ширины колонок по содержимому
 */
import { toolRegistry } from "../../../src/taskpane/tools/registry";

jest.mock("../../../src/taskpane/tools/backup", () => ({
  undoManager: { createBackup: jest.fn().mockResolvedValue(undefined) },
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

import "../../../src/taskpane/tools/format";
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

function setupAutoFitMock(state: ScenarioState) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });
  const rangeMock = createMockRange("TestSheet!A1:Z100", state);
  Object.assign(rangeMock, {
    autofitColumns: jest.fn(),
    format: { autofitColumns: jest.fn() },
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
                getRange: jest.fn().mockReturnValue(rangeMock),
                getUsedRangeOrNullObject: jest
                  .fn()
                  .mockReturnValue(
                    Object.assign({}, rangeMock, { isNullObject: false }),
                  ),
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
  return rangeMock;
}

describe("Сценарий 18: Авто-ширина колонок", () => {
  let state: ScenarioState;

  beforeEach(() => {
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: autoFitColumns на диапазон", async () => {
    const r = setupAutoFitMock(state);
    const res = JSON.parse(
      await toolRegistry.execute("autoFitColumns", { address: "A1:Z100" }),
    );
    expect(res.ok).toBe(true);
    expect(res.data).toBeDefined();
  });

  test("autoFitColumns без address — использует used range", async () => {
    setupAutoFitMock(state);
    const res = JSON.parse(await toolRegistry.execute("autoFitColumns", {}));
    expect(res.ok).toBe(true);
  });
});
