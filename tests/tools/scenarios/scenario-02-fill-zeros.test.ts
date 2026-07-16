/**
 * Эталонный сценарий №2: «Заполнить колонку нулями»
 *
 * Инструменты: W3 fillRange
 *
 * Сценарий:
 *   1. Создать колонку (заголовок + данные через setValues)
 *   2. Заполнить остаток колонки нулями через fillRange с fillType="value" / fillValue=0
 */
import { toolRegistry } from "../../../src/taskpane/tools/registry";

// ---------------------------------------------------------------------------
// Mocks
// ---------------------------------------------------------------------------

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

// Mock withPerformanceGuard for setValues
jest.mock("../../../src/taskpane/tools/_shared/performance", () => {
  const actual = jest.requireActual(
    "../../../src/taskpane/tools/_shared/performance",
  );
  return {
    ...actual,
    withPerformanceGuard: jest
      .fn()
      .mockImplementation(
        async (callback: (ctx: unknown) => Promise<unknown>) => {
          return Excel.run(async (context) => {
            return callback(context);
          });
        },
      ),
  };
});

import "../../../src/taskpane/tools/write";

import {
  createScenarioState,
  setupExcelMock,
  cleanupExcelMock,
  type ScenarioState,
} from "./_mock";

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("Сценарий 2: Заполнить колонку нулями", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });

  afterEach(() => {
    cleanupExcelMock();
  });

  test("fillRange value — заполнить колонку A1:A10 нулями", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("fillRange", {
        address: "A1:A10",
        fillType: "value",
        fillValue: 0,
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(10);
    expect(r.summary).toContain("10");
    expect(r.data).toMatchObject({ fillType: "value" });
  });

  test("fillRange copy — альтернативный тип для заполнения", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("fillRange", {
        address: "B1:B8",
        fillType: "copy",
        fillValue: 0,
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(8);
  });

  test("fillRange value — горизонтальный диапазон A1:F1", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("fillRange", {
        address: "A1:F1",
        fillType: "value",
        fillValue: 0,
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(6);
  });

  test("ПОЛНЫЙ ФЛОУ: заголовок → данные → дозаполнить нулями", async () => {
    setupExcelMock(state);

    // Step 1: Write header
    const step1 = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "A1",
        values: [["Остаток"]],
      }),
    );
    expect(step1.ok).toBe(true);

    // Step 2: Write actual data rows
    const step2 = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "A2:A5",
        values: [[100], [200], [150], [300]],
      }),
    );
    expect(step2.ok).toBe(true);

    // Step 3: Fill remaining rows with zeros
    const step3 = JSON.parse(
      await toolRegistry.execute("fillRange", {
        address: "A6:A100",
        fillType: "value",
        fillValue: 0,
      }),
    );
    expect(step3.ok).toBe(true);
    expect(step3.cellsAffected).toBe(95);

    // All 3 write operations should have created undo snapshots
    expect(mockCreateBackup).toHaveBeenCalledTimes(3);
  });

  test("Ошибка: fillRange без address", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("fillRange", {
        fillType: "value",
        fillValue: 0,
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("Ошибка: fillRange с неизвестным fillType", async () => {
    setupExcelMock(state);
    const r = JSON.parse(
      await toolRegistry.execute("fillRange", {
        address: "A1:A10",
        fillType: "invalid_type",
      }),
    );
    expect(r.ok).toBe(false);
    expect(r.error?.code).toBe("INVALID_FILL_TYPE");
  });

  test("fillRange возвращает правильные метаданные в data", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("fillRange", {
        address: "C1:C50",
        fillType: "value",
        fillValue: 0,
      }),
    );

    expect(r.ok).toBe(true);
    expect(r.data).toMatchObject({
      rows: 50,
      cols: 1,
      cellCount: 50,
      fillType: "value",
    });
  });
});
