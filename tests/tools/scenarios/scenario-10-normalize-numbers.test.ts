/**
 * Сценарий №10: «Числа «1 234,5» → 1234.5»
 *
 * Инструменты: R4 detectDataTypes, T5 normalizeText
 *
 * Сценарий:
 *   1. Определить колонки с числами-как-текстом
 *   2. Через normalizeText (cleanWhitespace) убрать пробелы
 *   3. Записать очищенные числа
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
import "../../../src/taskpane/tools/read";
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

const NUM_DATA = [
  ["Сумма"],
  ["1 234,5"],
  ["567 890,00"],
  ["12 345"],
  ["1 000 000"],
];

function setupNumMock(state: ScenarioState) {
  const syncMock = jest.fn().mockImplementation(async () => {
    state.syncCalls++;
  });
  const range = createMockRange("TestSheet!A1:A5", state, { rows: 5, cols: 1 });

  (globalThis as { Excel?: unknown }).Excel = {
    ClearApplyTo: { contents: "Contents", formats: "Formats" },
    CalculationMode: { manual: "manual", automatic: "automatic" },
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        const ctx = {
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue({
                name: "TestSheet",
                load: jest.fn(),
                getRange: jest.fn().mockImplementation((addr: string) => {
                  const r = createMockRange(addr, state);
                  Object.defineProperty(r, "values", {
                    get: () => NUM_DATA,
                    set: (v: unknown[][]) => {
                      state.writtenValues = v;
                    },
                    configurable: true,
                  });
                  r.rowCount = 5;
                  r.columnCount = 1;
                  return r;
                }),
                getUsedRangeOrNullObject: jest
                  .fn()
                  .mockReturnValue(
                    Object.assign(
                      {},
                      createMockRange("TestSheet!A1:E10", state),
                      { isNullObject: false },
                    ),
                  ),
              }),
              getItem: jest.fn(),
            },
          },
          sync: syncMock,
          runtime: { enableEvents: false },
          application: { calculationMode: "" },
        };
        return fn(ctx);
      }),
  };
}

describe("Сценарий 10: Числа «1 234,5» → 1234.5", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: normalizeText с cleanWhitespace", async () => {
    setupNumMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "A1:A5",
        operations: ["cleanWhitespace"],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data).toBeDefined();
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("normalizeText с trim+cleanWhitespace", async () => {
    setupNumMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "A1:A5",
        operations: ["trim", "cleanWhitespace"],
      }),
    );
    expect(r.ok).toBe(true);
  });

  test("normalizeText без address → MISSING_ADDRESS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", { operations: ["trim"] }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("normalizeText с пустыми operations → EMPTY_OPS", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("normalizeText", {
        address: "A1",
        operations: [],
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("EMPTY_OPS");
  });
});
