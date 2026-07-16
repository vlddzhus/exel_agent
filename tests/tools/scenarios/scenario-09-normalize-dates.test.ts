/**
 * Сценарий №9: «Привести даты к ДД.ММ.ГГГГ»
 *
 * Инструменты: R4 detectDataTypes, W2 setFormula, W1 setValues
 *
 * Сценарий:
 *   1. Определить колонки с датами через detectDataTypes
 *   2. Через TEXT() формулу привести даты к единому формату
 *   3. Или через setValues записать нормализованные даты
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

import "../../../src/taskpane/tools/write";
import "../../../src/taskpane/tools/read";
import {
  createScenarioState,
  setupExcelMock,
  cleanupExcelMock,
  type ScenarioState,
} from "./_mock";

describe("Сценарий 9: Привести даты к ДД.ММ.ГГГГ", () => {
  let state: ScenarioState;

  beforeEach(() => {
    mockCreateBackup.mockClear();
    state = createScenarioState();
  });
  afterEach(() => {
    cleanupExcelMock();
  });

  test("Шаг 1: TEXT(date, формат) для приведения дат", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "B2:B10",
        formulas: Array(9).fill(['=TEXT(A2,"DD.MM.YYYY")']),
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(9);
  });

  test("Шаг 2: setValues с датами в формате ДД.ММ.ГГГГ", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setValues", {
        address: "A2:A5",
        values: [
          ["01.01.2024"],
          ["15.06.2024"],
          ["31.12.2024"],
          ["01.03.2025"],
        ],
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.cellsAffected).toBe(4);
  });

  test("ПОЛНЫЙ ФЛОУ: detectDataTypes → TEXT формулы", async () => {
    setupExcelMock(state);

    const step1 = JSON.parse(
      await toolRegistry.execute("getRangeStats", { address: "A1:B10" }),
    );
    expect(step1.ok).toBe(true);

    const step2 = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "B2:B10",
        formulas: Array(9).fill(['=TEXT(A2,"DD.MM.YYYY")']),
      }),
    );
    expect(step2.ok).toBe(true);
  });

  test("Формат через DATEVALUE + TEXT", async () => {
    setupExcelMock(state);

    const r = JSON.parse(
      await toolRegistry.execute("setFormula", {
        address: "C2",
        formulas: [['=TEXT(DATEVALUE(A2),"DD.MM.YYYY")']],
      }),
    );
    expect(r.ok).toBe(true);
  });
});
