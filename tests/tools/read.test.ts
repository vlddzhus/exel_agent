/**
 * Тесты для getWorkbookOverview v2 — критический фикс N+1 sync.
 *
 * Главный ассерт: на книге с N листами context.sync вызывается ОГРАНИЧЕННОЕ
 * число раз (3: базовая инфа → usedRange → headers/samples + tables),
 * а не 5*N как было раньше.
 */
import { getWorkbookOverviewTool } from "../../src/taskpane/tools/read";

// ---------------------------------------------------------------------------
// Mock Excel для overview
// ---------------------------------------------------------------------------

interface MockSheet {
  name: string;
  position: number;
  visible: boolean;
  usedRange: { address: string; rowCount: number; columnCount: number; values: unknown[][] } | null;
  headers: string[];
}

function setupMockExcel(sheets: MockSheet[], activeName: string): { syncCount: number } {
  const state = { syncCount: 0 };

  function makeRange(addr: string, values: unknown[][]) {
    return {
      address: addr,
      rowCount: values.length,
      columnCount: values[0]?.length ?? 0,
      values,
      isNullObject: false,
      load: () => {},
      getRow: (_n: number) => makeRange(addr, [values[0] ?? []]),
      getCell: (r: number, c: number) => ({
        getResizedRange: (dr: number, dc: number) => {
          const endRow = r + dr + 1;
          const endCol = c + dc + 1;
          const sub: unknown[][] = [];
          for (let i = r; i < endRow && i < values.length; i++) {
            sub.push((values[i] ?? []).slice(c, endCol));
          }
          return makeRange(addr, sub);
        },
      }),
    };
  }

  const nullRange = {
    address: "",
    rowCount: 0,
    columnCount: 0,
    values: [],
    isNullObject: true,
    load: () => {},
    getRow: () => null,
    getCell: () => null,
  };

  const worksheetsMock = {
    items: sheets.map((s) => ({
      name: s.name,
      position: s.position,
      visibility: s.visible ? "Visible" : "Hidden",
      getUsedRangeOrNullObject: () => (s.usedRange === null ? nullRange : makeRange(s.usedRange.address, s.usedRange.values)),
    })),
    getActiveWorksheet: () => ({
      name: activeName,
      load: () => {},
    }),
    load: () => {},
  };

  const tablesMock = {
    items: [],
    load: () => {},
  };

  (globalThis as { Excel?: unknown }).Excel = {
    run: async (fn: (ctx: unknown) => Promise<unknown>) => {
      const ctx = {
        workbook: { worksheets: worksheetsMock, tables: tablesMock },
        sync: async () => {
          state.syncCount++;
        },
      };
      return fn(ctx);
    },
  };

  return state;
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

describe("getWorkbookOverview v2 — фикс N+1 sync", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("возвращает корректную структуру для книги с одним листом", async () => {
    setupMockExcel(
      [
        {
          name: "Продажи",
          position: 0,
          visible: true,
          usedRange: {
            address: "Продажи!A1:C3",
            rowCount: 3,
            columnCount: 3,
            values: [
              ["Месяц", "Менеджер", "Сумма"],
              ["Янв", "Иван", 100],
              ["Фев", "Мария", 200],
            ],
          },
          headers: ["Месяц", "Менеджер", "Сумма"],
        },
      ],
      "Продажи",
    );

    const result = await getWorkbookOverviewTool.execute({});

    expect(result.ok).toBe(true);
    expect(result.data).toBeDefined();
    const data = result.data as { sheets: unknown[]; totalSheets: number };
    expect(data.totalSheets).toBe(1);
    expect(data.sheets.length).toBe(1);
  });

  test("КРИТИЧНО: на книге с 10 листами sync вызывается НЕ больше 5 раз", async () => {
    const sheets: MockSheet[] = Array.from({ length: 10 }, (_, i) => ({
      name: `Лист${i + 1}`,
      position: i,
      visible: true,
      usedRange: {
        address: `Лист${i + 1}!A1:B5`,
        rowCount: 5,
        columnCount: 2,
        values: [
          ["a", "b"],
          ["1", "2"],
          ["3", "4"],
          ["5", "6"],
          ["7", "8"],
        ],
      },
      headers: ["a", "b"],
    }));

    const state = setupMockExcel(sheets, "Лист1");

    await getWorkbookOverviewTool.execute({});

    // До фикса: ~4 sync на лист = 40+ sync.
    // После фикса: 3 sync (базовая инфа, usedRanges, headers+samples+tables).
    expect(state.syncCount).toBeLessThanOrEqual(5);
  });

  test("КРИТИЧНО: на книге с 30 листами sync вызывается НЕ больше 5 раз", async () => {
    const sheets: MockSheet[] = Array.from({ length: 30 }, (_, i) => ({
      name: `S${i + 1}`,
      position: i,
      visible: true,
      usedRange: {
        address: `S${i + 1}!A1:C3`,
        rowCount: 3,
        columnCount: 3,
        values: [["x", "y", "z"], ["1", "2", "3"], ["4", "5", "6"]],
      },
      headers: ["x", "y", "z"],
    }));

    const state = setupMockExcel(sheets, "S1");

    await getWorkbookOverviewTool.execute({});

    // На 30 листах раньше было 120+ sync. Теперь ≤ 5.
    expect(state.syncCount).toBeLessThanOrEqual(5);
  });

  test("summary содержит имя активного листа", async () => {
    setupMockExcel(
      [
        {
          name: "Данные",
          position: 0,
          visible: true,
          usedRange: null,
          headers: [],
        },
      ],
      "Данные",
    );

    const result = await getWorkbookOverviewTool.execute({});
    expect(result.summary).toContain("Данные");
  });

  test("пустые листы (без usedRange) корректно учитываются", async () => {
    setupMockExcel(
      [
        {
          name: "Пустой",
          position: 0,
          visible: true,
          usedRange: null,
          headers: [],
        },
        {
          name: "СДанными",
          position: 1,
          visible: true,
          usedRange: {
            address: "СДанными!A1:B2",
            rowCount: 2,
            columnCount: 2,
            values: [["a", "b"], ["1", "2"]],
          },
          headers: ["a", "b"],
        },
      ],
      "СДанными",
    );

    const result = await getWorkbookOverviewTool.execute({});
    const data = result.data as { sheets: { name: string; rowCount: number }[] };
    const empty = data.sheets.find((s) => s.name === "Пустой");
    const withData = data.sheets.find((s) => s.name === "СДанными");
    expect(empty?.rowCount).toBe(0);
    expect(withData?.rowCount).toBe(2);
  });

  test("riskLevel = safe (не меняет данные)", () => {
    expect(getWorkbookOverviewTool.riskLevel).toBe("safe");
    expect(getWorkbookOverviewTool.requiresUndo).toBe(false);
  });
});
