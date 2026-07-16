/**
 * Сценарий №17: «Закрепить первую строку»
 *
 * Инструменты: S5 freezePanes
 *
 * Сценарий:
 *   1. Закрепить первую строку (freezeRow)
 *   2. Закрепить первую колонку (freezeColumn)
 *   3. Отменить закрепление
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

import "../../../src/taskpane/tools/structure";
import {
  createScenarioState,
  cleanupExcelMock,
  createMockRange,
  type ScenarioState,
} from "./_mock";

describe("Сценарий 17: Закрепить первую строку", () => {
  let state: ScenarioState;

  beforeEach(() => {
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  function setupFreezeMock() {
    const syncMock = jest.fn().mockImplementation(async () => {
      state.syncCalls++;
    });
    const fpMock = {
      freezeRows: jest.fn(),
      freezeColumns: jest.fn(),
      freezeAt: jest.fn(),
      unfreeze: jest.fn(),
    };

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
                  getRange: jest
                    .fn()
                    .mockImplementation((addr: string) =>
                      createMockRange(addr, state),
                    ),
                  freezePanes: fpMock,
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
    return fpMock;
  }

  test("Шаг 1: freezePanes firstRow", async () => {
    setupFreezeMock();
    const r = JSON.parse(
      await toolRegistry.execute("freezePanes", { target: "firstRow" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data).toMatchObject({ target: "firstRow" });
  });

  test("Шаг 2: freezePanes firstColumn", async () => {
    setupFreezeMock();
    const r = JSON.parse(
      await toolRegistry.execute("freezePanes", { target: "firstColumn" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data).toMatchObject({ target: "firstColumn" });
  });

  test("Шаг 3: freezePanes none (отмена)", async () => {
    const fp = setupFreezeMock();
    const r = JSON.parse(
      await toolRegistry.execute("freezePanes", { target: "none" }),
    );
    expect(r.ok).toBe(true);
    expect(fp.unfreeze).toHaveBeenCalled();
  });

  test("freezePanes без target → MISSING_TARGET", async () => {
    const r = JSON.parse(await toolRegistry.execute("freezePanes", {}));
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_TARGET");
  });
});
