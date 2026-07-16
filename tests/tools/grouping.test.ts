/**
 * Тесты для grouping.ts — manageGrouping.
 * Паттерн как в format-extended.test.ts.
 */
import { toolRegistry } from "../../src/taskpane/tools/registry";

const mockCreateBackup = jest.fn().mockResolvedValue(undefined);
jest.mock("../../src/taskpane/tools/backup", () => ({
  undoManager: {
    createBackup: (...args: unknown[]) => mockCreateBackup(...args),
  },
}));

import "../../src/taskpane/tools/grouping";

function setupExcelMock() {
  const syncMock = jest.fn().mockResolvedValue(undefined);
  const groupSpy = jest.fn();
  const ungroupSpy = jest.fn();
  const clearMethodsSpy = jest.fn();

  const makeRange = (addr: string): any => ({
    address: addr,
    rowCount: 10,
    columnCount: 4,
    isNullObject: false,
    load: jest.fn(),
    group: groupSpy,
    ungroup: ungroupSpy,
    getCell: jest.fn(),
    getResizedRange: jest.fn(),
  });

  const mockSheet: any = {
    name: "TestSheet",
    load: jest.fn(),
    getRange: jest.fn().mockImplementation((addr: string) => makeRange(addr)),
    getUsedRangeOrNullObject: jest.fn().mockReturnValue(makeRange("used")),
    outline: {
      summaryBelow: true,
      summaryRight: true,
      clearMethods: clearMethodsSpy,
    },
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

  return { syncMock, mockSheet, groupSpy, ungroupSpy, clearMethodsSpy };
}

describe("manageGrouping", () => {
  beforeEach(() => {
    mockCreateBackup.mockClear();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  test("groupRows — вызывает range.group(ByRows)", async () => {
    const { groupSpy } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageGrouping", {
        action: "groupRows",
        address: "A5:A10",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.axis).toBe("rows");
    expect(groupSpy).toHaveBeenCalledWith("ByRows");
    expect(mockCreateBackup).toHaveBeenCalled();
  });

  test("groupColumns — вызывает range.group(ByColumns)", async () => {
    const { groupSpy } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageGrouping", {
        action: "groupColumns",
        address: "C1:F10",
      }),
    );
    expect(r.ok).toBe(true);
    expect(r.data.axis).toBe("columns");
    expect(groupSpy).toHaveBeenCalledWith("ByColumns");
  });

  test("ungroupRows — вызывает range.ungroup()", async () => {
    const { ungroupSpy } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageGrouping", {
        action: "ungroupRows",
        address: "A5:A10",
      }),
    );
    expect(r.ok).toBe(true);
    expect(ungroupSpy).toHaveBeenCalled();
  });

  test("clearOutline — не требует address", async () => {
    const { clearMethodsSpy } = setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageGrouping", { action: "clearOutline" }),
    );
    expect(r.ok).toBe(true);
    expect(clearMethodsSpy).toHaveBeenCalled();
  });

  test("groupRows без address → MISSING_ADDRESS", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageGrouping", { action: "groupRows" }),
    );
    expect(r.error?.code).toBe("MISSING_ADDRESS");
  });

  test("summaryBelow/Right — передаются в outline", async () => {
    const { mockSheet } = setupExcelMock();
    await toolRegistry.execute("manageGrouping", {
      action: "groupRows",
      address: "A1:A5",
      summaryBelow: false,
      summaryRight: false,
    });
    expect(mockSheet.outline.summaryBelow).toBe(false);
    expect(mockSheet.outline.summaryRight).toBe(false);
  });

  test("невалидный action → INVALID_ACTION", async () => {
    setupExcelMock();
    const r = JSON.parse(
      await toolRegistry.execute("manageGrouping", { action: "hide", address: "A1:A5" }),
    );
    expect(r.error?.code).toBe("INVALID_ACTION");
  });

  test("риск moderate + требует undo", () => {
    expect(toolRegistry.riskLevel("manageGrouping")).toBe("moderate");
    expect(toolRegistry.requiresUndo("manageGrouping")).toBe(true);
  });
});
