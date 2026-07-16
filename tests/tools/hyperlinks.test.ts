/**
 * Тесты для hyperlinks.ts — manageHyperlinks.
 * Паттерн как в format-extended.test.ts.
 */
import { toolRegistry } from "../../src/taskpane/tools/registry";

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

import "../../src/taskpane/tools/hyperlinks";

function setupExcelMock(opts?: { initialHyperlink?: any }) {
  const syncMock = jest.fn().mockResolvedValue(undefined);
  let hyperlink = opts?.initialHyperlink ?? null;

  const makeRange = (addr: string): any => {
    const range: any = {
      address: addr,
      load: jest.fn(),
      get hyperlink() {
        return hyperlink;
      },
      set hyperlink(value: any) {
        hyperlink = value;
      },
    };
    return range;
  };

  const mockSheet = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockImplementation((addr: string) => makeRange(addr)),
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

  return { syncMock, getHyperlink: () => hyperlink };
}

describe("manageHyperlinks", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("add — добавляет внешнюю ссылку (URL)", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageHyperlinks", {
        action: "add",
        address: "A1",
        target: "https://example.com",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.type).toBe("external");
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("add — добавляет внутреннюю ссылку (на лист)", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageHyperlinks", {
        action: "add",
        address: "B2",
        target: "Sheet2!A1",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.type).toBe("internal");
  });

  test("add mailto — определяется как внешняя", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageHyperlinks", {
        action: "add",
        address: "A1",
        target: "mailto:info@example.com",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.type).toBe("external");
  });

  test("add без target → MISSING_TARGET", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageHyperlinks", { action: "add", address: "A1" }),
    );
    expect(r.error?.code).toBe("MISSING_TARGET");
  });

  test("get — возвращает существующую ссылку", async () => {
    setupExcelMock({ initialHyperlink: { address: "https://x.com" } });
    const r = JSON.parse(
      await toolRegistry.execute("manageHyperlinks", { action: "get", address: "A1" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.hyperlink).not.toBeNull();
  });

  test("get — без ссылки возвращает hyperlink=null", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageHyperlinks", { action: "get", address: "A1" }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.hyperlink).toBeNull();
  });

  test("remove — очищает ссылку", async () => {
    setupExcelMock({ initialHyperlink: { address: "https://x.com" } });
    const r = JSON.parse(
      await toolRegistry.execute("manageHyperlinks", { action: "remove", address: "A1" }),
    );
    expect(r.ok).toBe(true);
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("без address → MISSING_ADDRESS", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageHyperlinks", { action: "get" }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("невалидный action → INVALID_ACTION", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageHyperlinks", { action: "update", address: "A1" }),
    );
    expect(r.error?.code).toBe("INVALID_ACTION");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("manageHyperlinks")).toBe("moderate");
    expect(toolRegistry.requiresUndo("manageHyperlinks")).toBe(true);
  });
});
