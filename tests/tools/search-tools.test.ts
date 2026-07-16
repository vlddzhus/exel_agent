/**
 * Тесты для search-tools.ts — findAndReplace.
 *
 * Итерация «Расширение инструментов». Паттерн как в format-extended.test.ts:
 * мок globalThis.Excel + мок undoManager.
 */
import { toolRegistry } from "../../src/taskpane/tools/registry";

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

import "../../src/taskpane/tools/search-tools";

/**
 * Mock Excel: range.find() возвращает RangeAreas с адресами совпадений,
 * заданных через `matches`. Для замены — range.values читается/пишется.
 */
function setupExcelMock(opts?: { matches?: string[]; cellValues?: Record<string, unknown> }) {
  const syncMock = jest.fn().mockResolvedValue(undefined);
  const matches = opts?.matches ?? ["Test!A1", "Test!A5"];
  const cellValues = opts?.cellValues ?? {};

  const makeRangeAreas = (addr: string): any => ({
    address: addr,
    isNullObject: false,
    load: jest.fn(),
  });

  const makeRange = (addr: string): any => {
    let currentVal = cellValues[addr];
    return {
      address: addr,
      rowCount: 10,
      columnCount: 4,
      isNullObject: false,
      load: jest.fn(),
      values: currentVal !== undefined ? [[currentVal]] : [[""]],
      getCell: jest.fn().mockImplementation((_r: number, _c: number) => makeRange(`${addr}!${_r}x${_c}`)),
      getResizedRange: jest.fn().mockImplementation(() => makeRange(`${addr}~resized`)),
      // find: возвращает следующий RangeAreas по списку matches
      find: jest.fn().mockImplementation((text: string) => {
        // Простой последовательный возврат адресов совпадений
        const idx = (makeRange as any)._findIdx ?? 0;
        if (idx < matches.length) {
          (makeRange as any)._findIdx = idx + 1;
          return makeRangeAreas(matches[idx]);
        }
        return { address: "", isNullObject: true, load: jest.fn() };
      }),
      // setter для values через defineProperty ниже
    };
  };

  // Сброс счётчика find между тестами
  (makeRange as any)._findIdx = 0;

  const baseRange = makeRange("Test!A1:D10");

  // Per-cell range с читаемыми/перезаписываемыми values (для замены)
  const cellRangeFactory = (addr: string) => {
    const r: any = makeRange(addr);
    let stored: any = cellValues[addr] ?? "";
    Object.defineProperty(r, "values", {
      get: () => [[stored]],
      set: (v: unknown[][]) => {
        stored = v[0]?.[0];
        cellValues[addr] = stored;
      },
      configurable: true,
    });
    return r;
  };

  const mockSheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockImplementation((addr: string) => cellRangeFactory(addr)),
    getUsedRangeOrNullObject: jest.fn().mockReturnValue(baseRange),
  };

  (globalThis as { Excel?: unknown }).Excel = {
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        const ctx = {
          workbook: {
            worksheets: {
              getActiveWorksheet: jest.fn().mockReturnValue(mockSheet),
              getItem: jest.fn().mockReturnValue(mockSheet),
            },
          },
          sync: syncMock,
        };
        return fn(ctx);
      }),
  };

  return { syncMock, mockSheet, cellValues };
}

describe("findAndReplace", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
    setupExcelMock({ matches: ["Test!A1", "Test!A3"], cellValues: { "Test!A1": "hello", "Test!A3": "hello" } });
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("findAll — возвращает список адресов совпадений", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("findAndReplace", {
        action: "findAll",
        findWhat: "hello",
        address: "A1:D10",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.matches).toBeGreaterThan(0);
    expect(Array.isArray(r.data.addresses)).toBe(true);
  });

  test("find — возвращает первое совпадение", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("findAndReplace", {
        action: "find",
        findWhat: "hello",
        address: "A1:D10",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.matches).toBe(1);
  });

  test("replaceAll — заменяет значения и создаёт undo-снапшот", async () => {
    const { cellValues } = setupExcelMock({
      matches: ["Test!A1"],
      cellValues: { "Test!A1": "old" },
    });
    const r = JSON.parse(
      await toolRegistry.execute("findAndReplace", {
        action: "replaceAll",
        findWhat: "old",
        replaceWith: "new",
        address: "A1:D10",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.replaced).toBe(1);
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("replace без replaceWith → MISSING_REPLACE_WITH", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("findAndReplace", {
        action: "replace",
        findWhat: "old",
        address: "A1:D10",
      }),
    );
    expect(r.error?.code).toBe("MISSING_REPLACE_WITH");
  });

  test("невалидный action → INVALID_ACTION", async () => {
    const r = JSON.parse(
      await toolRegistry.execute("findAndReplace", {
        action: "unknown",
        findWhat: "x",
      }),
    );
    expect(r.error?.code).toBe("INVALID_ACTION");
  });

  test("риск dangerous + требует undo", () => {
    expect(toolRegistry.riskLevel("findAndReplace")).toBe("dangerous");
    expect(toolRegistry.requiresUndo("findAndReplace")).toBe(true);
  });
});
