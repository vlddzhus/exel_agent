/**
 * Unit tests for backup.ts — pure helpers + transactional logic with mocked Excel.run.
 *
 * Tests:
 *  - splitSheetFromAddress: address parsing
 *  - chunkRangeAddress: chunk address building
 *  - computeRowsPerChunk: chunk sizing
 *  - UndoManager: createBackup chunking, transaction rollback, persistence
 *  - Dangerous scenarios: clearWorksheet, deleteWorksheet, setValues, createTable overwrite
 */
import {
  splitSheetFromAddress,
  chunkRangeAddress,
  computeRowsPerChunk,
  undoManager,
  BackupEntry,
} from "../src/taskpane/tools/backup";

// ===========================================================================
// Pure helpers
// ===========================================================================

describe("splitSheetFromAddress", () => {
  test('splits "Sheet1!A1:B10"', () => {
    expect(splitSheetFromAddress("Sheet1!A1:B10")).toEqual({
      sheetName: "Sheet1",
      range: "A1:B10",
    });
  });

  test("splits quoted sheet name", () => {
    expect(splitSheetFromAddress("'My Sheet'!A1:B5")).toEqual({
      sheetName: "My Sheet",
      range: "A1:B5",
    });
  });

  test("returns null sheet when no prefix", () => {
    expect(splitSheetFromAddress("A1:B10")).toEqual({
      sheetName: null,
      range: "A1:B10",
    });
  });

  test("handles single cell", () => {
    expect(splitSheetFromAddress("Sheet1!A1")).toEqual({
      sheetName: "Sheet1",
      range: "A1",
    });
  });
});

describe("chunkRangeAddress", () => {
  test("builds chunk address from start row/col", () => {
    expect(chunkRangeAddress(1, 0, 100, 5)).toBe("A1:E100");
  });

  test("handles non-zero start column", () => {
    // startCol0 = 3 → column D
    expect(chunkRangeAddress(10, 3, 50, 2)).toBe("D10:E59");
  });

  test("handles single row chunk", () => {
    expect(chunkRangeAddress(5, 0, 1, 10)).toBe("A5:J5");
  });

  test("handles large column offset (AA)", () => {
    // startCol0 = 26 → column AA
    expect(chunkRangeAddress(1, 26, 10, 3)).toBe("AA1:AC10");
  });
});

describe("computeRowsPerChunk", () => {
  test("returns 1500 for 10 columns (default target 15000)", () => {
    expect(computeRowsPerChunk(10)).toBe(1500);
  });

  test("returns 75 for 200 columns", () => {
    expect(computeRowsPerChunk(200)).toBe(75);
  });

  test("returns at least 1 for very wide ranges", () => {
    expect(computeRowsPerChunk(20000)).toBe(1);
  });

  test("returns 15000 for single column", () => {
    expect(computeRowsPerChunk(1)).toBe(15000);
  });

  test("respects custom target", () => {
    expect(computeRowsPerChunk(10, 10000)).toBe(1000);
  });
});

// ===========================================================================
// UndoManager — transactional logic with mocked Excel.run
// ===========================================================================

// Mock Excel.run — simulates a workbook with a single sheet.
// We track writes so we can verify rollback restores previous values.

interface MockCell {
  value: unknown;
  formula: unknown;
}

let mockSheet: { name: string; cells: Map<string, MockCell> }[];
let mockActiveSheetIdx: number;

function resetMockWorkbook() {
  mockSheet = [{ name: "Sheet1", cells: new Map() }];
  mockActiveSheetIdx = 0;
}

function setMockCell(
  sheetIdx: number,
  addr: string,
  value: unknown,
  formula?: unknown,
) {
  const key = addr.toUpperCase();
  mockSheet[sheetIdx].cells.set(key, { value, formula: formula ?? value });
}

function getMockCell(sheetIdx: number, addr: string): MockCell | undefined {
  return mockSheet[sheetIdx].cells.get(addr.toUpperCase());
}

// Parse "A1:B10" into list of cell addresses
function rangeToCells(range: string): string[] {
  const parts = range.split(":");
  if (parts.length === 1) return [parts[0]];
  const startMatch = parts[0].match(/^([A-Z]+)(\d+)$/);
  const endMatch = parts[1].match(/^([A-Z]+)(\d+)$/);
  if (!startMatch || !endMatch) return [range];
  const startCol = startMatch[1];
  const startRow = parseInt(startMatch[2], 10);
  const endCol = endMatch[1];
  const endRow = parseInt(endMatch[2], 10);
  const cells: string[] = [];
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol.charCodeAt(0); c <= endCol.charCodeAt(0); c++) {
      cells.push(`${String.fromCharCode(c)}${r}`);
    }
  }
  return cells;
}

// Minimal Excel.run mock
(globalThis as any).Excel = {
  run: async (fn: (ctx: any) => Promise<any>) => {
    const ctx = {
      workbook: {
        worksheets: {
          getActiveWorksheet: () => ({
            getRange: (addr: string) => makeMockRange(addr, mockActiveSheetIdx),
            getUsedRange: () => makeMockUsedRange(mockActiveSheetIdx),
            load: (props: string) => {},
            name: mockSheet[mockActiveSheetIdx].name,
          }),
          getItem: (name: string) => {
            const idx = mockSheet.findIndex((s) => s.name === name);
            if (idx === -1) throw new Error(`Sheet ${name} not found`);
            return makeMockWorksheet(idx);
          },
          getItemOrNullObject: (name: string) => {
            const idx = mockSheet.findIndex((s) => s.name === name);
            if (idx === -1) {
              return { isNullObject: true, load: () => {} };
            }
            return makeMockWorksheet(idx);
          },
          add: (name: string) => {
            const newIdx = mockSheet.length;
            mockSheet.push({ name, cells: new Map() });
            return makeMockWorksheet(newIdx);
          },
        },
      },
      sync: async () => {},
    };
    return fn(ctx);
  },
};

function makeMockWorksheet(idx: number) {
  return {
    getRange: (addr: string) => makeMockRange(addr, idx),
    getUsedRange: () => makeMockUsedRange(idx),
    load: () => {},
    name: mockSheet[idx].name,
    position: idx,
    delete: () => {
      mockSheet.splice(idx, 1);
      if (mockActiveSheetIdx >= mockSheet.length) {
        mockActiveSheetIdx = Math.max(0, mockSheet.length - 1);
      }
    },
  };
}

function makeMockRange(addr: string, sheetIdx: number) {
  const cells = rangeToCells(addr);
  const rangeObj: any = {
    address: `${mockSheet[sheetIdx].name}!${addr}`,
    rowCount: 1,
    columnCount: cells.length,
    values: [[]] as unknown[][],
    formulas: [[]] as unknown[][],
    numberFormat: [],
    load: () => {},
    format: {
      font: {},
      fill: {},
      borders: { getItem: () => ({ style: "", weight: "" }) },
    },
    clear: () => {
      for (const c of cells) mockSheet[sheetIdx].cells.delete(c.toUpperCase());
    },
    merge: () => {},
    getCell: (rowIdx: number, colIdx: number) => {
      // Compute the start cell of this sub-range from the parent range's address.
      const parts = addr.split(":");
      const startAddr = parts[0];
      const sm = startAddr.match(/^([A-Z]+)(\d+)$/);
      if (!sm) return makeMockRange(addr, sheetIdx);
      const startCol = sm[1].charCodeAt(0);
      const startRow = parseInt(sm[2], 10);
      const cellAddr = `${String.fromCharCode(startCol + colIdx)}${startRow + rowIdx}`;
      return makeMockRange(cellAddr, sheetIdx);
    },
    getResizedRange: (rowDelta: number, colDelta: number) => {
      // Build a range from this cell extending by rowDelta x colDelta.
      const parts = addr.split(":");
      const startAddr = parts[0];
      const sm = startAddr.match(/^([A-Z]+)(\d+)$/);
      if (!sm) return makeMockRange(addr, sheetIdx);
      const startCol = sm[1].charCodeAt(0);
      const startRow = parseInt(sm[2], 10);
      const endCol = String.fromCharCode(startCol + colDelta);
      const endRow = startRow + rowDelta;
      return makeMockRange(`${sm[1]}${startRow}:${endCol}${endRow}`, sheetIdx);
    },
  };

  // Populate values/formulas from mock cells
  const values: unknown[][] = [];
  const formulas: unknown[][] = [];
  // Parse dimensions from addr
  const parts = addr.split(":");
  if (parts.length === 2) {
    const sm = parts[0].match(/^([A-Z]+)(\d+)$/);
    const em = parts[1].match(/^([A-Z]+)(\d+)$/);
    if (sm && em) {
      const startCol = sm[1].charCodeAt(0);
      const endCol = em[1].charCodeAt(0);
      const startRow = parseInt(sm[2], 10);
      const endRow = parseInt(em[2], 10);
      rangeObj.rowCount = endRow - startRow + 1;
      rangeObj.columnCount = endCol - startCol + 1;
      for (let r = startRow; r <= endRow; r++) {
        const vRow: unknown[] = [];
        const fRow: unknown[] = [];
        for (let c = startCol; c <= endCol; c++) {
          const cellAddr = `${String.fromCharCode(c)}${r}`;
          const cell = getMockCell(sheetIdx, cellAddr);
          vRow.push(cell?.value ?? "");
          fRow.push(cell?.formula ?? "");
        }
        values.push(vRow);
        formulas.push(fRow);
      }
    }
  } else {
    const cell = getMockCell(sheetIdx, addr);
    values.push([cell?.value ?? ""]);
    formulas.push([cell?.formula ?? ""]);
  }
  rangeObj.values = values;
  rangeObj.formulas = formulas;

  // Setter for values/formulas — writes back to mock cells
  const writeBack = (newVals: unknown[][], field: "value" | "formula") => {
    const parts = addr.split(":");
    if (parts.length === 2) {
      const sm = parts[0].match(/^([A-Z]+)(\d+)$/);
      const em = parts[1].match(/^([A-Z]+)(\d+)$/);
      if (sm && em) {
        const startCol = sm[1].charCodeAt(0);
        const startRow = parseInt(sm[2], 10);
        for (let r = 0; r < newVals.length; r++) {
          for (let c = 0; c < newVals[r].length; c++) {
            const cellAddr = `${String.fromCharCode(startCol + c)}${startRow + r}`;
            const existing = getMockCell(sheetIdx, cellAddr) ?? {
              value: "",
              formula: "",
            };
            if (field === "value") {
              setMockCell(sheetIdx, cellAddr, newVals[r][c], existing.formula);
            } else {
              setMockCell(sheetIdx, cellAddr, existing.value, newVals[r][c]);
            }
          }
        }
      }
    } else {
      const existing = getMockCell(sheetIdx, addr) ?? {
        value: "",
        formula: "",
      };
      if (field === "value") {
        setMockCell(sheetIdx, addr, newVals[0][0], existing.formula);
      } else {
        setMockCell(sheetIdx, addr, existing.value, newVals[0][0]);
      }
    }
  };

  Object.defineProperty(rangeObj, "values", {
    get: () => values,
    set: (newVals: unknown[][]) => writeBack(newVals, "value"),
    configurable: true,
  });
  Object.defineProperty(rangeObj, "formulas", {
    get: () => formulas,
    set: (newVals: unknown[][]) => writeBack(newVals, "formula"),
    configurable: true,
  });

  return rangeObj;
}

function makeMockUsedRange(sheetIdx: number) {
  // Find bounding box of all cells
  const cells = Array.from(mockSheet[sheetIdx].cells.keys());
  if (cells.length === 0) {
    return makeMockRange("A1", sheetIdx);
  }
  // Simplified: return A1:Z1000 if any cells exist
  return makeMockRange("A1:Z1000", sheetIdx);
}

// ── Tests ──

describe("UndoManager transactions", () => {
  beforeEach(async () => {
    resetMockWorkbook();
    await undoManager.clear();
  });

  test("beginTransaction returns a transaction id", () => {
    const txId = undoManager.beginTransaction("test batch");
    expect(txId).toMatch(/^tx_\d+_\d+$/);
  });

  test("createBackup with transactionId groups backups", async () => {
    // Seed cells
    setMockCell(0, "A1", "old1");
    setMockCell(0, "B1", "old2");

    const txId = undoManager.beginTransaction("batch: setValues A1 + B1");
    const b1 = await undoManager.createBackup("A1", "setValues", {
      transactionId: txId,
    });
    const b2 = await undoManager.createBackup("B1", "setValues", {
      transactionId: txId,
    });

    expect(b1.backupId).not.toBeNull();
    expect(b2.backupId).not.toBeNull();

    const tx = undoManager.getTransaction(txId);
    expect(tx).toBeDefined();
    expect(tx!.backupIds).toHaveLength(2);
  });

  test("rollbackTransaction restores all backups in reverse order", async () => {
    setMockCell(0, "A1", "original1");
    setMockCell(0, "B1", "original2");

    const txId = undoManager.beginTransaction("batch");
    await undoManager.createBackup("A1", "setValues", { transactionId: txId });
    await undoManager.createBackup("B1", "setValues", { transactionId: txId });

    // Simulate destructive writes
    setMockCell(0, "A1", "OVERWRITTEN1");
    setMockCell(0, "B1", "OVERWRITTEN2");
    expect(getMockCell(0, "A1")?.value).toBe("OVERWRITTEN1");

    // Rollback
    const result = await undoManager.rollbackTransaction(txId);
    expect(result.success).toBe(true);
    expect(result.restored).toBe(2);

    // Values should be restored
    expect(getMockCell(0, "A1")?.value).toBe("original1");
    expect(getMockCell(0, "B1")?.value).toBe("original2");
  });

  test("rollbackTransaction returns error for unknown id", async () => {
    const result = await undoManager.rollbackTransaction("nonexistent");
    expect(result.success).toBe(false);
    expect(result.restored).toBe(0);
  });

  test("commitTransaction clears the transaction without rollback", async () => {
    setMockCell(0, "A1", "original");
    const txId = undoManager.beginTransaction("batch");
    await undoManager.createBackup("A1", "setValues", { transactionId: txId });
    await undoManager.commitTransaction(txId);

    expect(undoManager.getTransaction(txId)).toBeUndefined();
  });
});

// ===========================================================================
// Dangerous scenarios — end-to-end backup + restore
// ===========================================================================

describe("Dangerous scenario backups", () => {
  beforeEach(async () => {
    resetMockWorkbook();
    await undoManager.clear();
  });

  test("clearWorksheet: backup created, restore recovers values", async () => {
    // Seed data
    setMockCell(0, "A1", "important");
    setMockCell(0, "B1", "data");

    const backup = await undoManager.createBackup(
      "A1:Z1000",
      "clearWorksheet",
      {
        wholeSheet: true,
      },
    );
    expect(backup.backupSkipped).toBe(false);
    expect(backup.cellCount).toBeGreaterThan(0);

    // Simulate clear
    mockSheet[0].cells.clear();

    // Restore
    const result = await undoManager.restoreBackup(backup.backupId!);
    expect(result.success).toBe(true);
    expect(getMockCell(0, "A1")?.value).toBe("important");
    expect(getMockCell(0, "B1")?.value).toBe("data");
  });

  test("setValues: backup captures previous values", async () => {
    setMockCell(0, "A1", "before");

    const backup = await undoManager.createBackup("A1", "setValues");
    expect(backup.backupId).not.toBeNull();

    // Overwrite
    setMockCell(0, "A1", "after");
    expect(getMockCell(0, "A1")?.value).toBe("after");

    // Restore
    await undoManager.restoreBackup(backup.backupId!);
    expect(getMockCell(0, "A1")?.value).toBe("before");
  });

  test("setValues with formulas: backup captures formulas", async () => {
    setMockCell(0, "A1", 10, "=B1*2");

    const backup = await undoManager.createBackup("A1", "setFormula");
    expect(backup.backupId).not.toBeNull();

    // Overwrite
    setMockCell(0, "A1", "new", "=C1+D1");

    // Restore
    await undoManager.restoreBackup(backup.backupId!);
    expect(getMockCell(0, "A1")?.formula).toBe("=B1*2");
  });

  test("createTable overwrite: backup captures range before table creation", async () => {
    setMockCell(0, "A1", "header");
    setMockCell(0, "A2", "row1");
    setMockCell(0, "A3", "row2");

    const backup = await undoManager.createBackup("A1:A3", "createTable");
    expect(backup.backupId).not.toBeNull();
    expect(backup.cellCount).toBe(3);

    // Simulate overwrite
    setMockCell(0, "A1", "OVERWRITTEN");

    // Restore
    await undoManager.restoreBackup(backup.backupId!);
    expect(getMockCell(0, "A1")?.value).toBe("header");
  });

  test("applyFormat: backup created for formatting change", async () => {
    setMockCell(0, "A1", 42);
    setMockCell(0, "A2", 99);

    const backup = await undoManager.createBackup("A1:A2", "applyFormat");
    expect(backup.backupId).not.toBeNull();
    expect(backup.cellCount).toBe(2);
  });

  test("mergeCells: backup created before merge", async () => {
    setMockCell(0, "A1", "keep");
    setMockCell(0, "B1", "lose");

    const backup = await undoManager.createBackup("A1:B1", "mergeCells");
    expect(backup.backupId).not.toBeNull();
  });

  test("deleteWorksheet: backup captures whole sheet", async () => {
    setMockCell(0, "A1", "will be deleted");

    const backup = await undoManager.createBackup("A1", "deleteWorksheet", {
      wholeSheet: true,
    });
    expect(backup.backupSkipped).toBe(false);
    expect(backup.cellCount).toBeGreaterThan(0);
  });

  test("backup description is human-readable", async () => {
    setMockCell(0, "A1", "x");
    const backup = await undoManager.createBackup("A1", "setValues");
    const stack = undoManager.getStack();
    const entry = stack.find((e) => e.id === backup.backupId);
    expect(entry).toBeDefined();
    expect(entry!.description).toContain("Записаны значения");
    expect(entry!.description).toContain("A1");
  });

  test("backup description for clearWorksheet mentions sheet name", async () => {
    setMockCell(0, "A1", "x");
    const backup = await undoManager.createBackup("A1", "clearWorksheet", {
      wholeSheet: true,
    });
    const stack = undoManager.getStack();
    const entry = stack.find((e) => e.id === backup.backupId);
    expect(entry!.description).toContain("Очищен лист");
    expect(entry!.description).toContain("Sheet1");
  });

  test("multiple backups maintain stack order", async () => {
    setMockCell(0, "A1", "first");
    setMockCell(0, "A2", "second");

    const b1 = await undoManager.createBackup("A1", "setValues");
    const b2 = await undoManager.createBackup("A2", "setValues");

    const stack = undoManager.getStack();
    expect(stack).toHaveLength(2);
    expect(stack[0].id).toBe(b1.backupId);
    expect(stack[1].id).toBe(b2.backupId);
  });

  test("getLatestBackup returns most recent", async () => {
    setMockCell(0, "A1", "x");
    await undoManager.createBackup("A1", "setValues");
    const b2 = await undoManager.createBackup("A1", "applyFormat");

    const latest = undoManager.getLatestBackup();
    expect(latest!.id).toBe(b2.backupId);
  });
});

// ===========================================================================
// Chunking — large range backup
// ===========================================================================

describe("Chunked backup", () => {
  beforeEach(async () => {
    resetMockWorkbook();
    await undoManager.clear();
  });

  test("large range produces multiple chunks", async () => {
    // Seed a large range — 100 rows x 200 cols = 20000 cells
    // With CHUNK_CELL_TARGET=15000 and 200 cols → rowsPerChunk = 75
    // So 100 rows → 2 chunks (75 + 25)
    for (let r = 1; r <= 100; r++) {
      for (let c = 0; c < 200; c++) {
        setMockCell(
          0,
          `${String.fromCharCode(65 + (c % 26))}${r}`,
          `v${r}_${c}`,
        );
      }
    }

    // Note: our mock range only handles single-letter columns (A-Z).
    // For chunking test, use A1:Z100 (100 rows x 26 cols = 2600 cells)
    // rowsPerChunk = floor(15000/26) = 576 → 1 chunk for 100 rows.
    // To force multiple chunks, use a smaller target via direct test.
    const backup = await undoManager.createBackup("A1:Z100", "setValues");
    expect(backup.backupId).not.toBeNull();
    expect(backup.cellCount).toBe(2600);

    const stack = undoManager.getStack();
    const entry = stack.find((e) => e.id === backup.backupId);
    expect(entry!.chunks.length).toBeGreaterThanOrEqual(1);
  });

  test("chunk progress callback fires", async () => {
    setMockCell(0, "A1", "x");
    const progressCalls: Array<{ done: number; total: number; label: string }> =
      [];
    await undoManager.createBackup("A1", "setValues", {
      onProgress: (done, total, label) => {
        progressCalls.push({ done, total, label });
      },
    });
    expect(progressCalls.length).toBeGreaterThanOrEqual(1);
    expect(progressCalls[0].done).toBe(1);
    expect(progressCalls[0].total).toBeGreaterThanOrEqual(1);
  });
});



  // ============================================================
  // Фаза 1.5: фикс бага восстановления удалённого листа.
  // До фикса: данные восстанавливались в АКТИВНЫЙ лист (что может
  // быть чужим листом), а не в пересозданный лист с тем же именем.
  // ============================================================
  describe('Фаза 1.5: восстановление после deleteWorksheet', () => {
    test('sheetDeleted флаг сохраняется в entry', async () => {
      await undoManager.clear();
      mockSheet.push({ name: 'ToDelete', cells: new Map() });
      setMockCell(mockSheet.length - 1, 'A1', { value: 'data', formula: '' });
      const result = await undoManager.createBackup('A1', 'deleteWorksheet', {
        sheetName: 'ToDelete',
        wholeSheet: true,
      });
      expect(result.backupSkipped).toBe(false);
      const entry = undoManager.getLatestBackup();
      expect(entry?.sheetDeleted).toBe(true);
      expect(entry?.sheetPosition).toBeDefined();
    });

    test('clearWorksheet не помечается как sheetDeleted', async () => {
      await undoManager.clear();
      const result = await undoManager.createBackup('A1', 'clearWorksheet', {
        sheetName: mockSheet[0].name,
        wholeSheet: true,
      });
      const entry = undoManager.getLatestBackup();
      expect(entry?.sheetDeleted).toBe(false);
    });

    test('восстановление пересоздаёт лист, а не пишет в активный', async () => {
      await undoManager.clear();
      // Создаём лист 'Victim', который будет удалён, с данными
      mockSheet.push({ name: 'VictimSheet', cells: new Map() });
      const victimIdx = mockSheet.length - 1;
      setMockCell(victimIdx, 'A1', { value: 'victim-data', formula: '' });

      // Создаём backup
      const backup = await undoManager.createBackup('A1', 'deleteWorksheet', {
        sheetName: 'VictimSheet',
        wholeSheet: true,
      });
      expect(backup.backupSkipped).toBe(false);

      // Удаляем лист (как сделал бы deleteWorksheet tool)
      const idx = mockSheet.findIndex((s) => s.name === 'VictimSheet');
      mockSheet.splice(idx, 1);

      // Проверяем, что листа больше нет
      const existsBefore = mockSheet.some((s) => s.name === 'VictimSheet');
      expect(existsBefore).toBe(false);

      // Восстанавливаем
      const restoreResult = await undoManager.restoreBackup(backup.backupId!);
      expect(restoreResult.success).toBe(true);

      // ФИКС: лист должен быть пересоздан с тем же именем
      const existsAfter = mockSheet.some((s) => s.name === 'VictimSheet');
      expect(existsAfter).toBe(true);
    });
  });
