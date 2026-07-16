/**
 * Тесты для named-ranges.ts — manageNamedRanges.
 * Паттерн как в format-extended.test.ts: мок globalThis.Excel + мок undoManager.
 */
import { toolRegistry } from "../../src/taskpane/tools/registry";

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

import "../../src/taskpane/tools/named-ranges";

/**
 * Mock коллекции именованных диапазонов workbook.
 * items хранится в замыкании, делится между add/getItem/delete/list.
 */
function setupExcelMock(opts?: { existing?: Array<{ name: string; value: string; type?: string }> }) {
  const syncMock = jest.fn().mockResolvedValue(undefined);
  const store = new Map<string, { value: string; type: string; comment?: string }>();
  for (const it of opts?.existing ?? []) {
    store.set(it.name, { value: it.value, type: it.type ?? "Range", comment: undefined });
  }

  const namesCollection: any = {
    load: jest.fn().mockImplementation(function (this: any, _sel: string) {
      // Эмуляция Office.js load: после sync — items готовы
      this.items = Array.from(store.entries()).map(([name, v]) => ({
        name,
        value: v.value,
        type: v.type,
        comment: v.comment,
      }));
    }),
    add: jest.fn().mockImplementation((name: string, refersTo: string, comment?: string) => {
      if (store.has(name)) throw new Error(`Name already exists: ${name}`);
      const value = refersTo;
      store.set(name, { value, type: "Range", comment });
      const item: any = {
        name,
        value,
        type: "Range",
        comment,
        load: jest.fn(),
        delete: jest.fn().mockImplementation(() => store.delete(name)),
      };
      return item;
    }),
    getItem: jest.fn().mockImplementation((name: string) => {
      if (!store.has(name)) throw new Error(`Item not found: ${name}`);
      const v = store.get(name)!;
      return {
        name,
        value: v.value,
        type: v.type,
        comment: v.comment,
        load: jest.fn(),
        delete: jest.fn().mockImplementation(() => store.delete(name)),
      };
    }),
  };

  (globalThis as { Excel?: unknown }).Excel = {
    run: jest
      .fn()
      .mockImplementation(async (fn: (ctx: unknown) => Promise<unknown>) => {
        const ctx = {
          workbook: { names: namesCollection },
          sync: syncMock,
        };
        return fn(ctx);
      }),
  };

  return { syncMock, store };
}

describe("manageNamedRanges", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("add — создаёт именованный диапазон и делает undo-снапшот", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageNamedRanges", {
        action: "add",
        name: "SalesData",
        refersTo: "Sheet1!A1:D10",
        comment: "Продажи Q1",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.name).toBe("SalesData");
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("add с нормализацией refersTo (без ведущего =)", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageNamedRanges", {
        action: "add",
        name: "Total",
        refersTo: "A1",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.value).toContain("=A1");
  });

  test("add без refersTo → MISSING_REFERS_TO", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageNamedRanges", {
        action: "add",
        name: "X",
      }),
    );
    expect(r.error?.code).toBe("MISSING_REFERS_TO");
  });

  test("list — возвращает существующие имена", async () => {
    setupExcelMock({ existing: [{ name: "A", value: "=Sheet1!A1" }, { name: "B", value: "=Sheet1!B1" }] });
    const r = JSON.parse(await toolRegistry.execute("manageNamedRanges", { action: "list" }));
    expect(r.ok).toBe(true);
    expect(r.data.count).toBe(2);
  });

  test("list — пустая книга", async () => {
    setupExcelMock({ existing: [] });
    const r = JSON.parse(await toolRegistry.execute("manageNamedRanges", { action: "list" }));
    expect(r.ok).toBe(true);
    expect(r.data.count).toBe(0);
  });

  test("get — возвращает детали имени", async () => {
    setupExcelMock({ existing: [{ name: "Tax", value: "=Sheet1!B2" }] });
    const r = JSON.parse(
      await toolRegistry.execute("manageNamedRanges", { action: "get", name: "Tax" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.name).toBe("Tax");
  });

  test("get несуществующего → NAME_NOT_FOUND", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageNamedRanges", { action: "get", name: "Nope" }),
    );
    expect(r.error?.code).toBe("NAME_NOT_FOUND");
  });

  test("delete — удаляет имя и делает undo-снапшот", async () => {
    setupExcelMock({ existing: [{ name: "Old", value: "=Sheet1!A1" }] });
    const r = JSON.parse(
      await toolRegistry.execute("manageNamedRanges", { action: "delete", name: "Old" }),
    );
    expect(r.ok).toBe(true);
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("add/get/delete без name → MISSING_NAME", async () => {
    setupExcelMock();
    const r = JSON.parse(await toolRegistry.execute("manageNamedRanges", { action: "get" }));
    expect(r.error?.code).toBe("MISSING_NAME");
  });

  test("невалидный action → INVALID_ACTION", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageNamedRanges", { action: "rename", name: "x" }),
    );
    expect(r.error?.code).toBe("INVALID_ACTION");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("manageNamedRanges")).toBe("moderate");
    expect(toolRegistry.requiresUndo("manageNamedRanges")).toBe(true);
  });
});
